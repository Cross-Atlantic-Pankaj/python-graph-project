# Server Startup Guide - No Cache Files

## Why `__pycache__` files are generated

Python automatically creates `__pycache__` directories and `.pyc` files to store compiled bytecode. This speeds up module imports but takes up disk space.

## How to prevent cache file generation

### Option 1: Use the startup script (Recommended)
```bash
cd backend
python start_server.py
```

### Option 2: Set environment variable
```bash
export PYTHONDONTWRITEBYTECODE=1
cd backend
python app.py
```

### Option 3: Use Python flag
```bash
cd backend
python -B app.py
```

## Cleanup existing cache files

### Quick cleanup
```bash
./cleanup_cache.sh
```

### Manual cleanup
```bash
# Remove all __pycache__ directories
find . -name "__pycache__" -type d -exec rm -rf {} +

# Remove all .pyc files
find . -name "*.pyc" -delete
```

## For AWS/Production deployment

Add to your deployment script:
```bash
export PYTHONDONTWRITEBYTECODE=1
```

Or add to your environment variables:
```
PYTHONDONTWRITEBYTECODE=1
```

## Benefits of disabling cache

✅ **Saves disk space** on your server
✅ **Cleaner codebase** without cache files
✅ **Faster deployments** (no cache to clean up)
✅ **Consistent behavior** across environments

## Note

Disabling bytecode cache may slightly increase startup time, but the space savings and cleaner deployment make it worthwhile for production servers.
