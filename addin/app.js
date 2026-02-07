/* PowerPoint Bridge — WebSocket client */

var ws = null;
var reconnectAttempt = 0;
var BASE_DELAY = 500;
var MAX_DELAY = 30000;
var officeReady = false;

/* Office.js initialization */
Office.onReady(function(info) {
  officeReady = true;
  console.log('Office.js ready:', info.host, info.platform);
  updateStatus('connecting');
  initWebSocket();
});

/* Fallback for browser testing (no Office.js host) */
setTimeout(function() {
  if (!officeReady) {
    console.log('Office.js not detected — standalone mode');
    updateStatus('connecting');
    initWebSocket();
  }
}, 3000);

/* WebSocket connection with exponential backoff */
function initWebSocket() {
  connect();
}

function connect() {
  try {
    ws = new WebSocket('wss://localhost:8443');
  } catch (err) {
    console.error('WebSocket constructor error:', err);
    scheduleReconnect();
    return;
  }

  ws.onopen = function() {
    reconnectAttempt = 0;
    updateStatus('connected');
    console.log('WebSocket connected');
  };

  ws.onclose = function() {
    updateStatus('disconnected');
    console.log('WebSocket closed');
    scheduleReconnect();
  };

  ws.onerror = function(err) {
    console.error('WebSocket error:', err);
  };

  ws.onmessage = function(event) {
    try {
      var message = JSON.parse(event.data);
      handleCommand(message);
    } catch (err) {
      console.error('Failed to parse message:', err);
    }
  };
}

function scheduleReconnect() {
  var delay = Math.min(BASE_DELAY * Math.pow(2, reconnectAttempt), MAX_DELAY);
  var jitter = Math.floor(Math.random() * 1000);
  reconnectAttempt++;
  console.log('Reconnecting in ' + (delay + jitter) + 'ms (attempt ' + reconnectAttempt + ')');
  setTimeout(connect, delay + jitter);
}

/* Status display */
function updateStatus(state) {
  var el = document.getElementById('status');
  if (!el) return;
  var labels = { connected: 'Connected', disconnected: 'Disconnected', connecting: 'Connecting...' };
  el.textContent = labels[state] || state;
  el.className = 'status ' + state;
}

/* Command handler stub — Phase 3 implements real execution */
function handleCommand(message) {
  console.log('Received command:', message);
}
