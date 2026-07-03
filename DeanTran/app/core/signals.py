from PySide6.QtCore import QObject, Signal

class SignalBus(QObject):
    """
    Central Event Bus for the application.
    UI connects to these signals. Logic emits them.
    Singleton-like usage: Create one instance in run.py or main_window.
    """
    # Log: level (INFO/WARN/ERROR), message
    log_signal = Signal(str, str)
    
    # Progress: value (0-100), total (optional description or max)
    progress_total_signal = Signal(int, int)   # cur, max
    progress_current_signal = Signal(int, int) # cur, max
    
    # Status: simple text update
    status_signal = Signal(str)

# Global instance for Phase 1 simplicity
signals = SignalBus()
