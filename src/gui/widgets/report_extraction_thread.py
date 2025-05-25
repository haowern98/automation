"""
Weekly Report Extraction Thread

Background thread for extracting weekly report data without blocking the UI.
"""
from PyQt5.QtCore import QThread, pyqtSignal


class WeeklyReportExtractionThread(QThread):
    """Thread for extracting weekly report data without blocking the UI"""
    
    # Signals for communicating with the main thread
    finished = pyqtSignal(bool, list, str)  # success, data, error_message
    progress = pyqtSignal(str)  # progress message
    
    def __init__(self, extractor, date_range_str):
        """
        Initialize the extraction thread
        
        Args:
            extractor: WeeklyReportExtractor instance
            date_range_str (str): Date range string to extract
        """
        super().__init__()
        self.extractor = extractor
        self.date_range_str = date_range_str
    
    def run(self):
        """Run the extraction in a separate thread"""
        try:
            self.progress.emit("Starting weekly report extraction...")
            success, data, error_msg = self.extractor.extract_data_for_date_range_gui(self.date_range_str)
            self.finished.emit(success, data, error_msg)
        except Exception as e:
            self.finished.emit(False, [], f"Unexpected error: {str(e)}")