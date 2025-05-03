"""
Log Analyzer for Word Interface Explorer

Analyzes failure logs to identify patterns, recommend solutions,
and adapt behavior for more successful retries.
"""

import os
import glob
import json
import logging
import datetime
from pathlib import Path
import re
from typing import Dict, List, Any, Tuple, Optional

# Configure logging
logger = logging.getLogger(__name__)

# Path to failure logs
LOG_DIR = Path(r"C:\Users\djjme\OneDrive\Desktop\CC-Directory\Agent WordDoc\src\bootstrap_training\word_interface\i_failed_logs")
SCREENSHOT_DIR = LOG_DIR / "screenshots"

class LogAnalyzer:
    """Analyzes failure logs to identify patterns and recommend solutions"""
    
    def __init__(self):
        """Initialize the log analyzer"""
        self.common_errors = {
            # OCR-related errors
            "tesseract": {
                "pattern": re.compile(r"tesseract is not (installed|properly installed|in your PATH)", re.IGNORECASE),
                "solution": "Install Tesseract OCR or add it to PATH",
                "fix_function": self._fix_tesseract_issue
            },
            # Document-related errors
            "document_not_found": {
                "pattern": re.compile(r"(document|word) (not found|cannot be accessed)", re.IGNORECASE),
                "solution": "Ensure Microsoft Word is running and accessible",
                "fix_function": self._fix_word_access_issue
            },
            # COM errors
            "com_error": {
                "pattern": re.compile(r"(com_error|com error|pywintypes\.com_error)", re.IGNORECASE),
                "solution": "Reset COM connections to Microsoft Word",
                "fix_function": self._fix_com_issue
            },
            # UI element not found
            "element_not_found": {
                "pattern": re.compile(r"element ['\"](.*?)['\"] not found", re.IGNORECASE),
                "solution": "Recalibrate element positions",
                "fix_function": self._fix_element_position_issue
            },
            # Cursor errors
            "cursor_error": {
                "pattern": re.compile(r"(cursor|babyadam|robot_cursor)", re.IGNORECASE),
                "solution": "Fix cursor management issues",
                "fix_function": self._fix_cursor_issue
            },
            # Permission errors
            "permission_error": {
                "pattern": re.compile(r"(permission|access) (denied|error)", re.IGNORECASE),
                "solution": "Fix permission issues",
                "fix_function": self._fix_permission_issue
            }
        }
    
    def get_recent_logs(self, max_age_minutes: int = 5) -> List[Path]:
        """Get list of recent log files from the failure directory"""
        all_logs = sorted(LOG_DIR.glob("*.log"), key=os.path.getmtime, reverse=True)
        json_logs = sorted(LOG_DIR.glob("*.json"), key=os.path.getmtime, reverse=True)
        
        # Combine all logs and sort by modification time
        logs = sorted(all_logs + json_logs, key=os.path.getmtime, reverse=True)
        
        # Filter for recent logs
        if max_age_minutes > 0:
            cutoff_time = datetime.datetime.now() - datetime.timedelta(minutes=max_age_minutes)
            recent_logs = []
            for log in logs:
                mtime = datetime.datetime.fromtimestamp(os.path.getmtime(log))
                if mtime >= cutoff_time:
                    recent_logs.append(log)
            return recent_logs
        
        return logs[:10]  # Return at most 10 logs if not filtering by time
    
    def get_recent_screenshots(self, max_age_minutes: int = 5) -> List[Path]:
        """Get list of recent screenshot files"""
        if not SCREENSHOT_DIR.exists():
            return []
        
        screenshots = sorted(SCREENSHOT_DIR.glob("*.png"), key=os.path.getmtime, reverse=True)
        
        # Filter for recent screenshots
        if max_age_minutes > 0:
            cutoff_time = datetime.datetime.now() - datetime.timedelta(minutes=max_age_minutes)
            recent_screenshots = []
            for screenshot in screenshots:
                mtime = datetime.datetime.fromtimestamp(os.path.getmtime(screenshot))
                if mtime >= cutoff_time:
                    recent_screenshots.append(screenshot)
            return recent_screenshots
        
        return screenshots[:10]  # Return at most 10 screenshots if not filtering by time
    
    def analyze_logs(self, max_age_minutes: int = 5) -> Dict[str, Any]:
        """Analyze recent logs to identify issues and suggest fixes"""
        logs = self.get_recent_logs(max_age_minutes)
        screenshots = self.get_recent_screenshots(max_age_minutes)
        
        if not logs:
            return {
                "status": "no_logs",
                "message": "No recent failure logs found",
                "fixes": []
            }
        
        # Read log contents
        log_contents = []
        for log in logs:
            try:
                if log.suffix == '.json':
                    with open(log, 'r') as f:
                        content = json.load(f)
                        log_contents.append({
                            'path': str(log),
                            'content': json.dumps(content, indent=2),
                            'type': 'json'
                        })
                else:  # .log files
                    with open(log, 'r') as f:
                        content = f.read()
                        log_contents.append({
                            'path': str(log),
                            'content': content,
                            'type': 'text'
                        })
            except Exception as e:
                logger.error(f"Error reading log {log}: {e}")
        
        # Analyze log contents for patterns
        identified_issues = []
        fixes = []
        
        for log_entry in log_contents:
            content = log_entry['content']
            
            for error_type, error_info in self.common_errors.items():
                matches = error_info["pattern"].findall(content)
                if matches:
                    identified_issues.append({
                        "type": error_type,
                        "matches": matches,
                        "log": log_entry['path'],
                        "solution": error_info["solution"]
                    })
                    
                    # Add fix if not already present
                    if error_info["solution"] not in [f["solution"] for f in fixes]:
                        fixes.append({
                            "type": error_type,
                            "solution": error_info["solution"],
                            "fix_function": error_info["fix_function"],
                            "importance": self._calculate_importance(error_type, matches)
                        })
        
        # Sort fixes by importance
        fixes.sort(key=lambda x: x["importance"], reverse=True)
        
        return {
            "status": "analyzed",
            "log_count": len(logs),
            "screenshot_count": len(screenshots),
            "identified_issues": identified_issues,
            "fixes": fixes
        }
    
    def apply_fixes(self, analysis_result: Dict[str, Any]) -> Dict[str, Any]:
        """Apply recommended fixes based on log analysis"""
        if analysis_result["status"] != "analyzed" or not analysis_result["fixes"]:
            return {
                "status": "no_fixes_needed",
                "message": "No fixes to apply"
            }
        
        applied_fixes = []
        
        for fix in analysis_result["fixes"]:
            try:
                # Get the fix function
                fix_function = fix["fix_function"]
                
                # Apply the fix
                result = fix_function()
                
                applied_fixes.append({
                    "type": fix["type"],
                    "solution": fix["solution"],
                    "success": result["success"],
                    "message": result["message"]
                })
                
                if not result["success"]:
                    logger.warning(f"Fix for {fix['type']} failed: {result['message']}")
            except Exception as e:
                logger.error(f"Error applying fix for {fix['type']}: {e}")
                applied_fixes.append({
                    "type": fix["type"],
                    "solution": fix["solution"],
                    "success": False,
                    "message": f"Error: {str(e)}"
                })
        
        return {
            "status": "fixes_applied",
            "applied_fixes": applied_fixes
        }
    
    def _calculate_importance(self, error_type: str, matches: List[str]) -> int:
        """Calculate importance score for a particular error type"""
        # Base importance by error type
        importance_map = {
            "tesseract": 3,  # OCR issues are lower priority since they're optional
            "document_not_found": 10,  # Critical - can't proceed without document
            "com_error": 9,  # Critical COM issues
            "element_not_found": 8,  # Important UI issue
            "cursor_error": 6,  # Medium priority
            "permission_error": 7  # Important but not critical
        }
        
        base_importance = importance_map.get(error_type, 5)
        
        # Adjust importance based on number of matches
        match_count_factor = min(len(matches), 5)  # Cap at 5 to avoid over-weighting
        
        return base_importance + match_count_factor
    
    # Fix functions for different types of issues
    def _fix_tesseract_issue(self) -> Dict[str, Any]:
        """Fix Tesseract OCR issues"""
        try:
            # We can't really install Tesseract automatically, but we can make the system work without it
            # by modifying the document_awareness module to handle missing OCR gracefully
            return {
                "success": True,
                "message": "OCR functionality disabled for this retry",
                "settings": {"ocr_disabled": True}
            }
        except Exception as e:
            return {"success": False, "message": f"Failed to fix Tesseract issue: {e}"}
    
    def _fix_word_access_issue(self) -> Dict[str, Any]:
        """Fix Word access issues"""
        try:
            # Recommend a longer wait time before accessing Word
            return {
                "success": True, 
                "message": "Increased wait time for Word initialization",
                "settings": {"word_init_delay": 3.0}
            }
        except Exception as e:
            return {"success": False, "message": f"Failed to fix Word access issue: {e}"}
    
    def _fix_com_issue(self) -> Dict[str, Any]:
        """Fix COM connection issues"""
        try:
            # Force COM server reset
            import pythoncom
            pythoncom.CoInitialize()
            
            return {
                "success": True, 
                "message": "Reset COM connections",
                "settings": {"force_new_word_instance": True}
            }
        except Exception as e:
            return {"success": False, "message": f"Failed to fix COM issue: {e}"}
    
    def _fix_element_position_issue(self) -> Dict[str, Any]:
        """Fix element position issues"""
        try:
            # Force recalibration
            return {
                "success": True, 
                "message": "Recalibrating element positions",
                "settings": {"force_calibration": True}
            }
        except Exception as e:
            return {"success": False, "message": f"Failed to fix element position issue: {e}"}
    
    def _fix_cursor_issue(self) -> Dict[str, Any]:
        """Fix cursor management issues"""
        try:
            # Adjust cursor settings
            return {
                "success": True, 
                "message": "Adjusted cursor settings",
                "settings": {"cursor_size": "standard"}
            }
        except Exception as e:
            return {"success": False, "message": f"Failed to fix cursor issue: {e}"}
    
    def _fix_permission_issue(self) -> Dict[str, Any]:
        """Fix permission issues"""
        try:
            # Not much we can do automatically for permissions
            return {
                "success": False, 
                "message": "Permission issues need manual intervention",
                "settings": {}
            }
        except Exception as e:
            return {"success": False, "message": f"Failed to address permission issue: {e}"}
    
    def generate_retry_recommendations(self, analysis_result: Dict[str, Any]) -> Dict[str, Any]:
        """Generate recommendations for retry based on log analysis"""
        if analysis_result["status"] != "analyzed":
            return {
                "recommendation": "standard_retry",
                "message": "No specific issues identified. Performing standard retry.",
                "settings": {}
            }
        
        # Compile all setting changes from fixes
        settings = {}
        recommendations = []
        messages = []
        
        for fix in analysis_result.get("fixes", []):
            try:
                result = fix["fix_function"]()
                if result["success"]:
                    messages.append(result["message"])
                    recommendations.append(fix["solution"])
                    
                    # Add any settings from the fix
                    if "settings" in result:
                        settings.update(result["settings"])
            except Exception as e:
                logger.error(f"Error generating recommendation for {fix['type']}: {e}")
        
        if recommendations:
            return {
                "recommendation": "adaptive_retry",
                "message": "Identified issues and applied fixes: " + ", ".join(messages),
                "solutions": recommendations,
                "settings": settings
            }
        else:
            return {
                "recommendation": "standard_retry",
                "message": "No specific fixes available. Performing standard retry.",
                "settings": {}
            }

# Create a singleton instance
log_analyzer = LogAnalyzer()
