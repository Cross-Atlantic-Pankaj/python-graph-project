# Deployment Guide

## Issue: Vercel 250MB Size Limit

Your Flask application exceeds Vercel's 250MB serverless function size limit due to large Python libraries:
- `matplotlib` (~100MB)
- `pandas` (~50MB) 
- `plotly` (~30MB)
- Other dependencies

## Solutions:

### Option 1: Use Render.com (Recommended)
Render.com supports larger deployments and is better suited for Python applications with heavy dependencies.

1. Create account at render.com
2. Connect your GitHub repository
3. Deploy as a Web Service
4. Set environment variables (MONGO_URI, etc.)

### Option 2: Use Railway.app
Railway.app also supports larger Python applications.

### Option 3: Use Heroku
Heroku supports larger applications but requires credit card for verification.

### Option 4: Vercel with External API
Keep Vercel for the main app and use a separate service for chart generation.

## Current Setup:
- `requirements.txt` - Full dependencies for development
- `requirements-prod.txt` - Minimal dependencies for Vercel
- `routes/projects-simple.py` - Simplified version without chart libraries

## Next Steps:
1. Choose a deployment platform from above
2. Update environment variables
3. Deploy and test functionality 