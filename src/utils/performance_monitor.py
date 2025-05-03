"""
Performance monitoring utility for the Word AI Agent

Tracks and logs performance metrics.
"""

import time
import psutil
from typing import Dict, Any, List, Optional
from dataclasses import dataclass
from datetime import datetime

from src.core.logging import get_logger

logger = get_logger(__name__)

@dataclass
class PerformanceMetric:
    """Performance metric data structure"""
    name: str
    value: float
    unit: str
    timestamp: datetime
    context: Dict[str, Any] = None

class PerformanceMonitor:
    """Monitors and logs performance metrics"""
    
    def __init__(self, enabled: bool = True):
        self.enabled = enabled
        self.metrics: List[PerformanceMetric] = []
        self.start_times: Dict[str, float] = {}
    
    def start_timer(self, name: str) -> None:
        """Start a timer for measuring operation duration"""
        if not self.enabled:
            return
        
        self.start_times[name] = time.time()
        logger.debug(f"Started timer for {name}")
    
    def stop_timer(self, name: str, context: Dict[str, Any] = None) -> Optional[float]:
        """Stop a timer and record the duration"""
        if not self.enabled or name not in self.start_times:
            return None
        
        duration = time.time() - self.start_times[name]
        self.record_metric(name, duration, "seconds", context)
        logger.debug(f"Stopped timer for {name}: {duration:.3f} seconds")
        
        return duration
    
    def record_metric(self, name: str, value: float, unit: str, context: Dict[str, Any] = None) -> None:
        """Record a performance metric"""
        if not self.enabled:
            return
        
        metric = PerformanceMetric(
            name=name,
            value=value,
            unit=unit,
            timestamp=datetime.now(),
            context=context
        )
        
        self.metrics.append(metric)
        logger.debug(f"Recorded metric {name}: {value} {unit}")
    
    def get_system_metrics(self) -> Dict[str, float]:
        """Get current system performance metrics"""
        if not self.enabled:
            return {}
        
        try:
            cpu_percent = psutil.cpu_percent(interval=0.1)
            memory = psutil.virtual_memory()
            disk = psutil.disk_usage('/')
            
            system_metrics = {
                "cpu_percent": cpu_percent,
                "memory_percent": memory.percent,
                "memory_used_mb": memory.used / (1024 * 1024),  # Convert to MB
                "disk_percent": disk.percent,
                "disk_used_gb": disk.used / (1024 * 1024 * 1024)  # Convert to GB
            }
            
            # Record system metrics
            for name, value in system_metrics.items():
                unit = "%" if name.endswith("percent") else ("MB" if name.endswith("mb") else "GB")
                self.record_metric(name, value, unit)
            
            return system_metrics
        except Exception as e:
            logger.error(f"Error getting system metrics: {str(e)}")
            return {}
    
    def get_metrics_summary(self) -> Dict[str, Any]:
        """Get a summary of all recorded metrics"""
        if not self.enabled or not self.metrics:
            return {}
        
        summary = {}
        for metric in self.metrics:
            if metric.name not in summary:
                summary[metric.name] = {
                    "count": 0,
                    "total": 0,
                    "min": float('inf'),
                    "max": float('-inf'),
                    "unit": metric.unit
                }
            
            summary[metric.name]["count"] += 1
            summary[metric.name]["total"] += metric.value
            summary[metric.name]["min"] = min(summary[metric.name]["min"], metric.value)
            summary[metric.name]["max"] = max(summary[metric.name]["max"], metric.value)
        
        # Calculate averages
        for name, data in summary.items():
            data["average"] = data["total"] / data["count"]
        
        return summary
    
    def clear_metrics(self) -> None:
        """Clear all recorded metrics"""
        self.metrics = []
        self.start_times = {}
        logger.debug("Cleared all performance metrics")
    
    def enable(self) -> None:
        """Enable performance monitoring"""
        self.enabled = True
        logger.info("Performance monitoring enabled")
    
    def disable(self) -> None:
        """Disable performance monitoring"""
        self.enabled = False
        logger.info("Performance monitoring disabled")
    
    def log_summary(self) -> None:
        """Log a summary of all recorded metrics"""
        if not self.enabled or not self.metrics:
            return
        
        summary = self.get_metrics_summary()
        logger.info("Performance Metrics Summary:")
        
        for name, data in summary.items():
            logger.info(f"  {name}: avg={data['average']:.3f} {data['unit']}, min={data['min']:.3f} {data['unit']}, max={data['max']:.3f} {data['unit']}, count={data['count']}")
