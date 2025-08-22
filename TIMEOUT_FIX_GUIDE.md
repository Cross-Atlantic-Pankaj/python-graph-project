# API Timeout Fix for Batch Processing

## Problem
The batch processing API was timing out with `WORKER TIMEOUT` errors because the default Gunicorn timeout (30 seconds) was too short for chart generation operations.

## Solution
Created Gunicorn configuration with increased timeout limits:

### Key Changes:
- **Timeout**: Increased from 30s to 300s (5 minutes)
- **Graceful Timeout**: Increased to 300s for proper shutdown
- **Worker Processes**: Optimized for batch processing
- **Memory Management**: Added worker recycling to prevent memory leaks

## Files Created:

### 1. `backend/gunicorn.conf.py`
Complete Gunicorn configuration with:
- 5-minute timeout for batch processing
- Optimized worker settings
- Memory management
- Logging configuration

### 2. `backend/start_gunicorn.sh`
Startup script using the configuration file:
```bash
./backend/start_gunicorn.sh
```

### 3. `backend/start_with_timeout.sh`
Simple direct Gunicorn startup:
```bash
./backend/start_with_timeout.sh
```

## Usage on AWS Server:

### Option 1: Using Configuration File
```bash
cd /home/ubuntu/Python-Graph-Project/backend
./start_gunicorn.sh
```

### Option 2: Direct Command
```bash
cd /home/ubuntu/Python-Graph-Project/backend
gunicorn --bind 0.0.0.0:5001 --timeout 300 --graceful-timeout 300 --workers 2 app:create_app()
```

### Option 3: With PM2 (if using PM2)
```bash
pm2 start "gunicorn --bind 0.0.0.0:5001 --timeout 300 --graceful-timeout 300 --workers 2 app:create_app()" --name "graph-project-api"
```

## Timeout Settings Explained:

- **`timeout = 300`**: Worker timeout increased to 5 minutes
- **`graceful_timeout = 300`**: Graceful shutdown timeout
- **`workers = 2`**: Optimal for batch processing
- **`max_requests = 1000`**: Worker recycling to prevent memory leaks
- **`preload_app = True`**: Faster startup and better memory usage

## Benefits:
1. **No More Timeouts**: Batch processing can complete without worker timeouts
2. **Better Memory Management**: Worker recycling prevents memory accumulation
3. **Improved Performance**: Optimized settings for chart generation
4. **Stable Operation**: Proper graceful shutdown handling

## Monitoring:
The configuration includes detailed logging to monitor:
- Worker lifecycle events
- Memory usage
- Request processing times
- Error handling

## Important Notes:
- This fix addresses the timeout issue without changing any chart generation logic
- The 5-minute timeout should be sufficient for most batch processing operations
- If you need even longer timeouts, you can increase the values in `gunicorn.conf.py`
