import os

# Bind to all interfaces on Railway's PORT
bind = f"0.0.0.0:{os.environ.get('PORT', '5000')}"

# Single worker with threading - required for shared in-memory state (task status, log queues)
workers = 1
threads = 16
worker_class = "gthread"

# timeout=0 disables gunicorn worker kill timeout — safe with gthread and needed for SSE streams
# and long-running automation tasks
timeout = 0
keepalive = 5

# Allow large file uploads (50MB limit for Excel files)
limit_request_line = 8190
limit_request_fields = 200
