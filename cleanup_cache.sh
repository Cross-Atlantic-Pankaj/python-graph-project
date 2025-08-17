#!/bin/bash

echo "🧹 Cleaning up Python cache files..."

# Remove all __pycache__ directories
find . -name "__pycache__" -type d -exec rm -rf {} + 2>/dev/null || true

# Remove all .pyc files
find . -name "*.pyc" -delete 2>/dev/null || true

# Remove all .pyo files
find . -name "*.pyo" -delete 2>/dev/null || true

# Remove all .pyd files (compiled Python files)
find . -name "*.pyd" -delete 2>/dev/null || true

# Remove pytest cache
find . -name ".pytest_cache" -type d -exec rm -rf {} + 2>/dev/null || true

# Remove coverage files
find . -name ".coverage" -delete 2>/dev/null || true
find . -name "coverage.xml" -delete 2>/dev/null || true

echo "✅ Cache cleanup completed!"
echo ""
echo "To prevent future cache generation, use:"
echo "  export PYTHONDONTWRITEBYTECODE=1"
echo "  python start_server.py"


