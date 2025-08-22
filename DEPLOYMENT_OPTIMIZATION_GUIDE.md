# Deployment Optimization Guide

## Problem Analysis
Your AWS server was crashing after generating 20-25 reports due to high CPU utilization (99.8%). The main issues were:

### 1. **Matplotlib Memory Leaks (CRITICAL)**
- **Problem**: Matplotlib figures were not being properly closed, causing severe memory accumulation
- **Impact**: Each chart generation left figures in memory, leading to exponential memory growth
- **Solution**: Added proper `plt.close()` calls and garbage collection

### 2. **Inefficient Resource Management**
- **Problem**: No cleanup between report generations
- **Impact**: Memory and CPU resources accumulated over time
- **Solution**: Added systematic cleanup and resource monitoring

### 3. **High DPI Chart Generation**
- **Problem**: Charts generated at 200 DPI consumed excessive memory
- **Impact**: Each chart used more memory than necessary
- **Solution**: Reduced DPI to 150 (production) and 120 (optimized)

## Changes Made

### 1. **Memory Management Improvements**

#### In `backend/routes/projects.py`:
```python
# Added at the start of _generate_report function
import gc
plt.switch_backend('Agg')  # Non-interactive backend

# Added after each chart generation
plt.close(fig_mpl)
plt.close('all')  # Close all figures
gc.collect()  # Force garbage collection
```

#### In batch processing:
```python
# Added before each file processing
gc.collect()
plt.close('all')

# Added after each report
gc.collect()
plt.close('all')
```

### 2. **Configuration Optimization**

#### New file: `backend/config.py`
- **Matplotlib settings**: Reduced DPI, non-interactive backend
- **Memory thresholds**: Configurable cleanup intervals
- **Production optimizations**: Stricter resource limits

### 3. **Memory Monitoring**

#### New file: `backend/utils/memory_monitor.py`
- **Real-time monitoring**: Track memory and CPU usage
- **Automatic cleanup**: Force cleanup when thresholds exceeded
- **Resource logging**: Monitor performance during operations

### 4. **Dependencies Added**
- **psutil**: For system resource monitoring
- **Updated requirements.txt**: Added memory monitoring dependency

## Deployment Instructions

### 1. **Update Your AWS Server**

```bash
# SSH into your AWS instance
ssh -i your-key.pem ubuntu@your-aws-ip

# Navigate to your project directory
cd /path/to/your/project

# Pull the latest changes
git pull origin main

# Install new dependencies
pip install -r backend/requirements.txt

# Set environment variable for production
export FLASK_ENV=production

# Restart your application
sudo systemctl restart your-app-service
```

### 2. **Environment Variables**

Add these to your AWS environment:
```bash
export FLASK_ENV=production
export MONGO_URI=your-mongodb-connection-string
```

### 3. **System Configuration**

#### Increase system limits (if needed):
```bash
# Edit system limits
sudo nano /etc/security/limits.conf

# Add these lines:
* soft nofile 65536
* hard nofile 65536
* soft nproc 32768
* hard nproc 32768
```

#### Monitor system resources:
```bash
# Check memory usage
free -h

# Check CPU usage
top

# Monitor disk space
df -h
```

## Performance Monitoring

### 1. **Log Monitoring**
The application now logs resource usage:
```
📊 Starting report generation - Memory: 245.3MB (12.1%), CPU: 15.2%
🧹 Cleanup completed: 1250 objects collected, Memory: 180.2MB
```

### 2. **Memory Thresholds**
- **Warning**: 70% memory usage
- **Critical**: 80% memory usage
- **Auto-cleanup**: Triggered at 80% threshold

### 3. **Expected Performance**
- **Before**: Server crashes after 20-25 reports
- **After**: Should handle 100+ reports without issues
- **Memory usage**: Should remain stable around 200-400MB

## Troubleshooting

### 1. **If server still crashes:**
```bash
# Check application logs
sudo journalctl -u your-app-service -f

# Monitor memory in real-time
watch -n 1 'free -h && echo "---" && ps aux | grep python'
```

### 2. **If memory usage is still high:**
- Reduce `MATPLOTLIB_DPI` in `config.py`
- Increase `GARBAGE_COLLECTION_INTERVAL`
- Consider implementing report queuing

### 3. **If charts are low quality:**
- Increase `MATPLOTLIB_DPI` in `config.py`
- Balance between quality and memory usage

## Additional Recommendations

### 1. **AWS Instance Optimization**
- **Instance type**: Use t3.medium or larger for better performance
- **Storage**: Use EBS with sufficient IOPS
- **Auto-scaling**: Consider implementing auto-scaling for high load

### 2. **Application Architecture**
- **Queue system**: Implement Celery for background report generation
- **Caching**: Cache generated charts to avoid regeneration
- **Load balancing**: Use multiple instances for high availability

### 3. **Monitoring Setup**
- **CloudWatch**: Set up AWS CloudWatch for monitoring
- **Alerts**: Configure alerts for high CPU/memory usage
- **Logs**: Centralize logs for better debugging

## Expected Results

After implementing these changes:
- ✅ **Stable performance**: Server should handle 100+ reports
- ✅ **Reduced memory usage**: Memory should remain stable
- ✅ **Better monitoring**: Real-time resource tracking
- ✅ **Automatic cleanup**: Prevents memory leaks
- ✅ **Production ready**: Optimized for AWS deployment

## Support

If you encounter issues after deployment:
1. Check the application logs for error messages
2. Monitor system resources using the provided commands
3. Verify all environment variables are set correctly
4. Ensure all dependencies are installed

The optimizations should significantly improve your server's stability and performance for report generation.
