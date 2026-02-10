---
status: testing
phase: 04-mcp-tools
source: [04-01-SUMMARY.md]
started: 2026-02-08T06:45:06Z
updated: 2026-02-08T06:45:06Z
---

## Current Test
<!-- OVERWRITE each test - shows where we are -->

number: 1
name: MCP Tools Discovery
expected: |
  Claude Code lists the three MCP tools: get_presentation, get_slide, and execute_officejs.
  Each tool appears with its description and parameters.
awaiting: user response

## Tests

### 1. MCP Tools Discovery
expected: Claude Code lists the three MCP tools: get_presentation, get_slide, and execute_officejs. Each tool appears with its description and parameters.
result: [pending]

### 2. Get Presentation Overview
expected: Calling get_presentation returns JSON with all slides in the open presentation, including slide IDs and shape summaries (count, names, types for each slide).
result: [pending]

### 3. Get Slide Details
expected: Calling get_slide with a slide index returns detailed shape data for that slide, including text content, positions (left/top in points), sizes (width/height in points), and fill colors.
result: [pending]

### 4. Execute Office.js Code
expected: Calling execute_officejs with Office.js code (e.g., adding a shape) modifies the live presentation. The change appears immediately in PowerPoint.
result: [pending]

### 5. Round-Trip Verification
expected: Reading presentation state, then modifying it via execute_officejs, then reading again shows the modification reflected in the updated state.
result: [pending]

## Summary

total: 5
passed: 0
issues: 0
pending: 5
skipped: 0

## Gaps

[none yet]
