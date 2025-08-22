# Gunicorn configuration file for increased timeout limits
# This addresses the WORKER TIMEOUT issue in batch processing

# Server socket
bind = "0.0.0.0:5001"
backlog = 2048

# Worker processes
workers = 2
worker_class = "sync"
worker_connections = 1000
max_requests = 1000
max_requests_jitter = 50

# Timeout settings - Increased for batch processing
timeout = 300  # 5 minutes (default is 30 seconds)
keepalive = 2
graceful_timeout = 300  # 5 minutes for graceful shutdown

# Process naming
proc_name = "graph-project-api"

# Logging
accesslog = "-"
errorlog = "-"
loglevel = "info"
access_log_format = '%(h)s %(l)s %(u)s %(t)s "%(r)s" %(s)s %(b)s "%(f)s" "%(a)s" %(D)s'

# Security
limit_request_line = 4094
limit_request_fields = 100
limit_request_field_size = 8190

# Performance
preload_app = True
sendfile = True
reuse_port = True

# Memory management
max_requests_jitter = 50
worker_tmp_dir = "/dev/shm"

# Environment
raw_env = [
    "PYTHONDONTWRITEBYTECODE=1",
]

# Callbacks
def on_starting(server):
    server.log.info("🚀 Starting Graph Project API with increased timeout limits...")

def on_reload(server):
    server.log.info("🔄 Reloading Graph Project API...")

def worker_int(worker):
    worker.log.info("⚠️ Worker received INT or QUIT signal")

def pre_fork(server, worker):
    server.log.info("🔧 Worker spawned (pid: %s)", worker.pid)

def post_fork(server, worker):
    server.log.info("✅ Worker spawned (pid: %s)", worker.pid)

def post_worker_init(worker):
    worker.log.info("🎯 Worker initialized (pid: %s)", worker.pid)

def worker_abort(worker):
    worker.log.info("💥 Worker aborted (pid: %s)", worker.pid)
