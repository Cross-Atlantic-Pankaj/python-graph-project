import psutil
import gc
import matplotlib.pyplot as plt
import logging
from datetime import datetime

class MemoryMonitor:
    def __init__(self, logger=None):
        self.logger = logger or logging.getLogger(__name__)
        self.process = psutil.Process()
        self.memory_threshold = 0.8  # 80% memory usage threshold
        
    def get_memory_usage(self):
        """Get current memory usage in MB"""
        memory_info = self.process.memory_info()
        return memory_info.rss / 1024 / 1024  # Convert to MB
    
    def get_memory_percentage(self):
        """Get memory usage as percentage of total system memory"""
        return self.process.memory_percent()
    
    def get_cpu_usage(self):
        """Get current CPU usage percentage"""
        return self.process.cpu_percent()
    
    def check_memory_threshold(self):
        """Check if memory usage is above threshold"""
        memory_percent = self.get_memory_percentage()
        if memory_percent > self.memory_threshold * 100:
            self.logger.warning(f"⚠️ High memory usage detected: {memory_percent:.1f}%")
            return True
        return False
    
    def force_cleanup(self):
        """Force cleanup of matplotlib and garbage collection"""
        try:
            # Close all matplotlib figures
            plt.close('all')
            
            # Force garbage collection
            collected = gc.collect()
            
            # Get memory usage after cleanup
            memory_after = self.get_memory_usage()
            
            self.logger.info(f"🧹 Cleanup completed: {collected} objects collected, Memory: {memory_after:.1f}MB")
            return memory_after
        except Exception as e:
            self.logger.error(f"❌ Error during cleanup: {e}")
            return None
    
    def log_resource_usage(self, operation="Unknown"):
        """Log current resource usage"""
        memory_mb = self.get_memory_usage()
        memory_percent = self.get_memory_percentage()
        cpu_percent = self.get_cpu_usage()
        
        self.logger.info(f"📊 {operation} - Memory: {memory_mb:.1f}MB ({memory_percent:.1f}%), CPU: {cpu_percent:.1f}%")
        
        # Warn if usage is high
        if memory_percent > 70:
            self.logger.warning(f"⚠️ High memory usage during {operation}: {memory_percent:.1f}%")
        if cpu_percent > 80:
            self.logger.warning(f"⚠️ High CPU usage during {operation}: {cpu_percent:.1f}%")
    
    def monitor_operation(self, operation_name):
        """Context manager to monitor resource usage during an operation"""
        class OperationMonitor:
            def __init__(self, monitor, operation_name):
                self.monitor = monitor
                self.operation_name = operation_name
                self.start_time = None
                
            def __enter__(self):
                self.start_time = datetime.now()
                self.monitor.log_resource_usage(f"Starting {self.operation_name}")
                return self.monitor
                
            def __exit__(self, exc_type, exc_val, exc_tb):
                duration = (datetime.now() - self.start_time).total_seconds()
                self.monitor.log_resource_usage(f"Completed {self.operation_name} ({duration:.1f}s)")
                
                # Force cleanup after operation
                if exc_type is None:  # Only cleanup if no exception occurred
                    self.monitor.force_cleanup()
        
        return OperationMonitor(self, operation_name)

# Global memory monitor instance
memory_monitor = MemoryMonitor()

def get_memory_monitor():
    """Get the global memory monitor instance"""
    return memory_monitor


