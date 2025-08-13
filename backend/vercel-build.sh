#!/bin/bash

# Clean up any existing cache
rm -rf __pycache__
rm -rf .pytest_cache
rm -rf .coverage

# Remove test files
rm -f test_*.py

# Remove documentation
rm -f *.md

# Remove any large data files
find . -name "*.csv" -delete
find . -name "*.xlsx" -delete
find . -name "*.xls" -delete

echo "Build optimization completed" 