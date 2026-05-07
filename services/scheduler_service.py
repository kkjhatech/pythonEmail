import time
import threading
from typing import Callable, Optional
from datetime import datetime

from utils.logger import get_logger


class SchedulerService:
    """Service for scheduling periodic tasks."""
    
    def __init__(self, log_folder: str):
        self.logger = get_logger('SchedulerService', log_folder)
        self._running = False
        self._thread: Optional[threading.Thread] = None
        self._stop_event = threading.Event()
    
    def start(
        self,
        task: Callable,
        interval_minutes: int,
        run_immediately: bool = True
    ) -> bool:
        """
        Start the scheduler.
        
        Args:
            task: Function to execute periodically
            interval_minutes: Interval between executions
            run_immediately: Whether to run task immediately on start
        
        Returns:
            True if started successfully
        """
        if self._running:
            self.logger.warning("Scheduler is already running")
            return False
        
        self._running = True
        self._stop_event.clear()
        
        def run_loop():
            if run_immediately:
                self._execute_task(task)
            
            while not self._stop_event.is_set():
                self.logger.info(f"Next run in {interval_minutes} minutes")
                
                # Wait with ability to interrupt
                if self._stop_event.wait(timeout=interval_minutes * 60):
                    break
                
                if not self._stop_event.is_set():
                    self._execute_task(task)
        
        self._thread = threading.Thread(target=run_loop, daemon=True)
        self._thread.start()
        
        self.logger.info(f"Scheduler started with {interval_minutes} minute interval")
        return True
    
    def stop(self) -> bool:
        """Stop the scheduler."""
        if not self._running:
            self.logger.warning("Scheduler is not running")
            return False
        
        self._stop_event.set()
        self._running = False
        
        if self._thread and self._thread.is_alive():
            self._thread.join(timeout=5)
        
        self.logger.info("Scheduler stopped")
        return True
    
    def _execute_task(self, task: Callable):
        """Execute the scheduled task with error handling."""
        try:
            self.logger.info(f"Executing scheduled task at {datetime.now()}")
            task()
            self.logger.info("Scheduled task completed successfully")
        except Exception as e:
            self.logger.error(f"Scheduled task failed: {str(e)}")
    
    def is_running(self) -> bool:
        """Check if scheduler is currently running."""
        return self._running and self._thread and self._thread.is_alive()
    
    def run_once(self, task: Callable) -> bool:
        """Run task once immediately."""
        try:
            self.logger.info("Running task once")
            task()
            return True
        except Exception as e:
            self.logger.error(f"Task execution failed: {str(e)}")
            return False
