#!/bin/bash
# Fetch technology and company logos from multiple sources.
#
# Sources (in priority order):
#   1. gilbarbara/logos — 1400+ tech logos, AWS services (SVG, no auth)
#   2. LF AI Landscape — 470+ data/AI/ML logos (SVG, no auth)
#   3. Brandfetch search — company/brand logos (PNG, needs BRANDFETCH_CLIENT_ID)
#
# Usage:
#   ~/.claude/skills/powerpoint-mcp/scripts/fetch-logo.sh [OPTIONS] QUERY [QUERY...]
#
# Options:
#   --icon                    Prefer icon variant (no text, just symbol)
#   --size N                  Max width/height for Brandfetch (default: 400)
#   --format png|svg          Preferred format (default: svg for gilbarbara, png for brandfetch)
#   --outdir DIR              Output directory (default: ~/.cache/powerpoint-mcp/logos/)
#   --search-only             List matches without downloading
#   --source gl|lfai|bf|auto  Force source: gl=gilbarbara, lfai=LF AI, bf=brandfetch, auto=try all (default: auto)
#
# Examples:
#   ~/.claude/skills/powerpoint-mcp/scripts/fetch-logo.sh dbt
#   ~/.claude/skills/powerpoint-mcp/scripts/fetch-logo.sh --icon Airflow
#   ~/.claude/skills/powerpoint-mcp/scripts/fetch-logo.sh "AWS MSK" Kubernetes Snowflake
#   ~/.claude/skills/powerpoint-mcp/scripts/fetch-logo.sh --search-only dataflow
#   ~/.claude/skills/powerpoint-mcp/scripts/fetch-logo.sh --source bf "dbt Labs"

set -euo pipefail

# Global cache — shared across all projects
CACHE_DIR="$HOME/.cache/powerpoint-mcp"

# Defaults
ICON_MODE=false
SIZE="400"
OUTDIR="$CACHE_DIR/logos"
SEARCH_ONLY=false
SOURCE="auto"
QUERIES=()

# gilbarbara index URL (cached locally after first fetch)
GL_INDEX_URL="https://raw.githubusercontent.com/gilbarbara/logos/main/logos.json"
GL_RAW_URL="https://raw.githubusercontent.com/gilbarbara/logos/main/logos"
GL_CACHE="$CACHE_DIR/indexes/gilbarbara-logos.json"

# LF AI Landscape (data/AI/ML logos)
LFAI_YML_URL="https://raw.githubusercontent.com/lfai/landscape/master/landscape.yml"
LFAI_RAW_URL="https://raw.githubusercontent.com/lfai/landscape/master/hosted_logos"
LFAI_CACHE="$CACHE_DIR/indexes/lfai-landscape-logos.json"

# --- Argument parsing ---

while [[ $# -gt 0 ]]; do
    case "$1" in
        --icon)
            ICON_MODE=true; shift ;;
        --size)
            SIZE="$2"; shift 2 ;;
        --outdir)
            OUTDIR="$2"; shift 2 ;;
        --search-only)
            SEARCH_ONLY=true; shift ;;
        --source)
            SOURCE="$2"; shift 2
            if [[ ! "$SOURCE" =~ ^(gl|lfai|bf|auto)$ ]]; then
                echo "Error: --source must be gl, lfai, bf, or auto"; exit 1
            fi ;;
        --help|-h)
            sed -n '2,/^$/p' "$0" | sed 's/^# \?//'; exit 0 ;;
        -*)
            echo "Error: unknown option $1"; exit 1 ;;
        *)
            QUERIES+=("$1"); shift ;;
    esac
done

# --- Preflight ---

for cmd in curl jq python3; do
    if ! command -v "$cmd" &>/dev/null; then
        echo "Error: $cmd is required but not installed."; exit 1
    fi
done

if [[ ${#QUERIES[@]} -eq 0 ]]; then
    echo "Error: provide at least one search query."
    echo "Run with --help for usage."
    exit 1
fi

mkdir -p "$OUTDIR"

# --- gilbarbara/logos functions ---

gl_ensure_index() {
    if [[ -f "$GL_CACHE" ]]; then
        # Refresh if older than 7 days
        local is_fresh
        is_fresh=$(find "$GL_CACHE" -mtime -7 2>/dev/null)
        if [[ -n "$is_fresh" ]]; then
            return 0
        fi
    fi
    mkdir -p "$(dirname "$GL_CACHE")"
    echo "  Updating logo index..." >&2
    curl -fsSL "$GL_INDEX_URL" -o "$GL_CACHE" 2>/dev/null || {
        echo "  Warning: could not fetch gilbarbara index" >&2
        return 1
    }
}

gl_search() {
    local query="$1"
    gl_ensure_index || return 1

    # Split query into words for multi-word matching (e.g. "AWS MSK" matches "AWS MSK (...)")
    # Also match shortname with hyphens (e.g. "AWS MSK" → "aws-msk")
    local query_lower
    query_lower=$(echo "$query" | tr '[:upper:]' '[:lower:]')
    local query_hyphenated
    query_hyphenated=$(echo "$query_lower" | tr ' ' '-')

    # Search and sort: exact name/shortname matches first, then partial
    local results
    results=$(jq --arg q "$query_lower" --arg qh "$query_hyphenated" '
        [.[] | select(
            (.name | ascii_downcase | contains($q)) or
            (.shortname | ascii_downcase | contains($q)) or
            (.shortname | ascii_downcase | contains($qh)) or
            (.files[] | ascii_downcase | contains($qh))
        )]
        | sort_by(
            if (.name | ascii_downcase) == $q then 0
            elif (.shortname | ascii_downcase) == $q then 0
            elif (.shortname | ascii_downcase) == $qh then 0
            else 1 end
        )
        | .[0:8]
    ' "$GL_CACHE")

    local count
    count=$(echo "$results" | jq 'length')

    if [[ "$count" == "0" ]]; then
        return 1
    fi

    echo "  gilbarbara matches:"
    echo "$results" | jq -r 'to_entries[] | "    [\(.key + 1)] \(.value.name) → \(.value.files | join(", "))"'

    # Pick the best file from first match
    local files
    files=$(echo "$results" | jq -r '.[0].files[]')

    local selected_file=""
    if [[ "$ICON_MODE" == true ]]; then
        # Prefer -icon variant
        selected_file=$(echo "$files" | grep -i '\-icon' | head -1)
    fi
    if [[ -z "$selected_file" ]]; then
        # Prefer non-icon (full logo with text)
        selected_file=$(echo "$files" | grep -iv '\-icon' | head -1)
    fi
    if [[ -z "$selected_file" ]]; then
        selected_file=$(echo "$files" | head -1)
    fi

    echo "$selected_file"
}

gl_download() {
    local file="$1"
    local outfile="${OUTDIR}/${file}"
    local url="${GL_RAW_URL}/${file}"

    curl -fsSL "$url" -o "$outfile" 2>/dev/null || {
        echo "  Error: download failed for $file"
        return 1
    }

    local size_bytes
    size_bytes=$(wc -c < "$outfile" | tr -d ' ')
    echo "  Saved: $outfile ($size_bytes bytes)"
}

# --- LF AI Landscape functions ---

lfai_ensure_index() {
    if [[ -f "$LFAI_CACHE" ]]; then
        local is_fresh
        is_fresh=$(find "$LFAI_CACHE" -mtime -7 2>/dev/null)
        if [[ -n "$is_fresh" ]]; then
            return 0
        fi
    fi
    mkdir -p "$(dirname "$LFAI_CACHE")"
    echo "  Updating LF AI index..." >&2
    # Parse landscape.yml: extract name → logo pairs from items
    curl -fsSL "$LFAI_YML_URL" 2>/dev/null | python3 -c "
import sys, json, re
content = sys.stdin.read()
entries = []
lines = content.split('\n')
current_name = None
for line in lines:
    name_match = re.match(r'\s+name:\s+(.+)', line)
    logo_match = re.match(r'\s+logo:\s+(.+)', line)
    if name_match:
        current_name = name_match.group(1).strip()
    elif logo_match and current_name:
        logo = logo_match.group(1).strip()
        entries.append({'name': current_name, 'logo': logo})
        current_name = None
json.dump(entries, sys.stdout)
" > "$LFAI_CACHE" || {
        echo "  Warning: could not fetch LF AI index" >&2
        return 1
    }
}

lfai_search() {
    local query="$1"
    lfai_ensure_index || return 1

    local query_lower
    query_lower=$(echo "$query" | tr '[:upper:]' '[:lower:]')
    # Normalize: spaces/underscores/hyphens all equivalent
    local query_normalized
    query_normalized=$(echo "$query_lower" | tr ' _' '--')

    local results
    results=$(jq --arg q "$query_lower" --arg qn "$query_normalized" '
        [.[] | select(
            (.name | ascii_downcase | contains($q)) or
            (.name | ascii_downcase | gsub("[_ ]"; "-") | contains($qn)) or
            (.logo | ascii_downcase | gsub("[_]"; "-") | contains($qn))
        )]
        | sort_by(
            if (.name | ascii_downcase) == $q then 0
            elif (.name | ascii_downcase | gsub("[_ ]"; "-")) == $qn then 0
            else 1 end
        )
        | unique_by(.logo)
        | .[0:8]
    ' "$LFAI_CACHE")

    local count
    count=$(echo "$results" | jq 'length')

    if [[ "$count" == "0" ]]; then
        return 1
    fi

    echo "  LF AI matches:"
    echo "$results" | jq -r 'to_entries[] | "    [\(.key + 1)] \(.value.name) → \(.value.logo)"'

    # Return logo filename of first match
    echo "$results" | jq -r '.[0].logo'
}

lfai_download() {
    local file="$1"
    local outfile="${OUTDIR}/${file}"
    local url="${LFAI_RAW_URL}/${file}"

    curl -fsSL "$url" -o "$outfile" 2>/dev/null || {
        echo "  Error: download failed for $file"
        return 1
    }

    local size_bytes
    size_bytes=$(wc -c < "$outfile" | tr -d ' ')
    echo "  Saved: $outfile ($size_bytes bytes)"
}

# --- Brandfetch functions ---

bf_check() {
    if [[ -z "${BRANDFETCH_CLIENT_ID:-}" ]]; then
        echo "  Brandfetch: BRANDFETCH_CLIENT_ID not set (skipping)"
        echo "  To enable: export BRANDFETCH_CLIENT_ID=your_client_id"
        return 1
    fi
    return 0
}

bf_search() {
    local query="$1"
    bf_check || return 1

    local encoded
    encoded=$(python3 -c "import urllib.parse, sys; print(urllib.parse.quote(sys.argv[1]))" "$query")
    local url="https://api.brandfetch.io/v2/search/${encoded}?c=${BRANDFETCH_CLIENT_ID}"

    local response
    response=$(curl -fsSL "$url" 2>/dev/null) || {
        echo "  Brandfetch: search failed"
        return 1
    }

    local count
    count=$(echo "$response" | jq 'length')

    if [[ "$count" == "0" ]]; then
        return 1
    fi

    echo "  Brandfetch matches:"
    echo "$response" | jq -r 'to_entries[] | "    [\(.key + 1)] \(.value.name) (\(.value.domain))"'

    # Return first match JSON
    echo "$response" | jq -c '.[0]'
}

bf_download() {
    local brand_json="$1"
    local brand_id domain name icon_url token type

    brand_id=$(echo "$brand_json" | jq -r '.brandId')
    domain=$(echo "$brand_json" | jq -r '.domain')
    name=$(echo "$brand_json" | jq -r '.name')
    icon_url=$(echo "$brand_json" | jq -r '.icon')
    token=$(echo "$icon_url" | sed 's/.*c=//; s/&.*//')

    if [[ "$ICON_MODE" == true ]]; then
        type="icon"
    else
        type="logo"
    fi

    local url="https://cdn.brandfetch.io/${brand_id}/w/${SIZE}/h/${SIZE}/fallback/lettermark/${type}.png?c=${token}"
    local outfile="${OUTDIR}/${domain}-${type}.png"

    curl -fsSL "$url" -o "$outfile" 2>/dev/null || {
        echo "  Error: download failed for $name"
        return 1
    }

    # Verify it's a real image
    local filetype
    filetype=$(file -b "$outfile")
    if echo "$filetype" | grep -qi "html\|text"; then
        echo "  Error: got HTML instead of image for $name"
        rm -f "$outfile" 2>/dev/null
        return 1
    fi

    local size_bytes
    size_bytes=$(wc -c < "$outfile" | tr -d ' ')
    echo "  Saved: $outfile ($size_bytes bytes)"
}

# --- Main ---

total=${#QUERIES[@]}
failures=0

for i in "${!QUERIES[@]}"; do
    n=$((i + 1))
    query="${QUERIES[$i]}"
    echo "[${n}/${total}] Searching: $query"

    found=false

    # Try gilbarbara first (unless forced to brandfetch)
    if [[ "$found" == false && ("$SOURCE" == "auto" || "$SOURCE" == "gl") ]]; then
        if output=$(gl_search "$query"); then
            # Print matches (all lines except last = filename)
            echo "$output" | head -n -1

            if [[ "$SEARCH_ONLY" == false ]]; then
                selected_file=$(echo "$output" | tail -1)
                echo "  Downloading: $selected_file"
                gl_download "$selected_file" && found=true
            else
                found=true
            fi
        fi
    fi

    # Try LF AI Landscape (data/AI/ML logos)
    if [[ "$found" == false && ("$SOURCE" == "auto" || "$SOURCE" == "lfai") ]]; then
        if output=$(lfai_search "$query"); then
            echo "$output" | head -n -1

            if [[ "$SEARCH_ONLY" == false ]]; then
                selected_file=$(echo "$output" | tail -1)
                echo "  Downloading: $selected_file"
                lfai_download "$selected_file" && found=true
            else
                found=true
            fi
        fi
    fi

    # Fallback to Brandfetch (company logos)
    if [[ "$found" == false && ("$SOURCE" == "auto" || "$SOURCE" == "bf") ]]; then
        if output=$(bf_search "$query"); then
            echo "$output" | head -n -1

            if [[ "$SEARCH_ONLY" == false ]]; then
                brand_json=$(echo "$output" | tail -1)
                name=$(echo "$brand_json" | jq -r '.name')
                echo "  Downloading: $name (brandfetch)"
                bf_download "$brand_json" && found=true
            else
                found=true
            fi
        fi
    fi

    if [[ "$found" == false ]]; then
        echo "  No results found in any source."
        ((failures++))
    fi

    echo ""
done

if [[ "$SEARCH_ONLY" == true ]]; then
    exit 0
fi

if [[ "$failures" -gt 0 ]]; then
    echo "Done. $((total - failures))/${total} succeeded."
    exit 1
else
    echo "Done. ${total} logo(s) saved to ${OUTDIR}/"
fi
