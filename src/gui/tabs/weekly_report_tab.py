"""
Weekly Report Tab

Tab for viewing and exporting weekly reports from Excel files.
Modified to save as TXT instead of HTML.
"""
from PyQt5.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QLabel, QDateEdit, 
                             QPushButton, QGridLayout, QGroupBox, QProgressBar,
                             QSplitter, QLineEdit, QFileDialog, QMessageBox)
from PyQt5.QtCore import Qt, QDate
from PyQt5.QtGui import QFont

# Try to import QWebEngineView, fall back to QTextEdit if not available
try:
    from PyQt5.QtWebEngineWidgets import QWebEngineView
    WEB_ENGINE_AVAILABLE = True
except ImportError:
    from PyQt5.QtWidgets import QTextEdit
    WEB_ENGINE_AVAILABLE = False

# Import the extraction thread
from PyQt5.QtCore import Qt, QDate, QThread, pyqtSignal
from src.gui.widgets.report_extraction_thread import WeeklyReportExtractionThread

class CombinedWeeklyReportExtractionThread(QThread):
    """Thread for extracting combined MFA + GSN VS AD report data without blocking the UI"""
    
    # Signals for communicating with the main thread
    finished = pyqtSignal(bool, object, str)  # success, data (dict or list), error_message
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
        """Run the combined extraction in a separate thread"""
        try:
            self.progress.emit("Starting combined MFA + GSN VS AD report extraction...")
            success, data, error_msg = self.extractor.extract_combined_data_for_date_range_gui(self.date_range_str)
            self.finished.emit(success, data, error_msg)
        except Exception as e:
            self.finished.emit(False, {}, f"Unexpected error: {str(e)}")

class WeeklyReportTab(QWidget):
    """Weekly report viewer tab with date selector and HTML display"""
    
    def __init__(self, parent=None):
        """Initialize the weekly report tab"""
        super(WeeklyReportTab, self).__init__(parent)
        
        # Initialize the extractor
        from src.processors.weekly_report_extractor import WeeklyReportExtractor
        self.extractor = WeeklyReportExtractor()
        
        # Create main layout
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(20, 20, 20, 20)
        
        # Create header
        header_label = QLabel("Generate Report from Excel")
        header_font = QFont("Segoe UI", 14)
        header_font.setBold(True)
        header_label.setFont(header_font)
        
        description_label = QLabel("Extract and view weekly reports from the Excel file")
        description_label.setStyleSheet("color: #666; margin-bottom: 10px;")
        
        main_layout.addWidget(header_label)
        main_layout.addWidget(description_label)
        
        # Create date selection section
        date_section = self._create_date_selection_section()
        main_layout.addWidget(date_section)
        
        # Create status section
        self.status_label = QLabel("Ready to generate report")
        self.status_label.setStyleSheet("color: #666; font-size: 11px; padding: 5px;")
        main_layout.addWidget(self.status_label)
        
        # Create progress bar (hidden by default)
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        self.progress_bar.setRange(0, 0)  # Indeterminate progress
        main_layout.addWidget(self.progress_bar)
        
        # Create splitter for report display and controls
        splitter = QSplitter(Qt.Vertical)
        
        # Create report display section
        report_section = self._create_report_display_section()
        splitter.addWidget(report_section)
        
        # Create export section
        export_section = self._create_export_section()
        splitter.addWidget(export_section)
        
        # Set splitter proportions
        splitter.setSizes([400, 100])
        
        main_layout.addWidget(splitter)
        
        # Initialize UI state
        self._update_ui_state(False)
        
        # Initialize extraction thread
        self.extraction_thread = None
    
    def _create_date_selection_section(self):
        """Create the date selection section"""
        group_box = QGroupBox("Select Report Date Range")
        layout = QGridLayout(group_box)
        
        # Start date
        layout.addWidget(QLabel("Start Date:"), 0, 0)
        self.start_date_picker = QDateEdit()
        self.start_date_picker.setCalendarPopup(True)
        self.start_date_picker.setDate(QDate.currentDate().addDays(-7))  # Default to a week ago
        layout.addWidget(self.start_date_picker, 0, 1)
        
        # End date
        layout.addWidget(QLabel("End Date:"), 0, 2)
        self.end_date_picker = QDateEdit()
        self.end_date_picker.setCalendarPopup(True)
        self.end_date_picker.setDate(QDate.currentDate())  # Default to today
        layout.addWidget(self.end_date_picker, 0, 3)
        
        # Generate button
        self.generate_button = QPushButton("Generate Report")
        self.generate_button.setDefault(True)
        self.generate_button.clicked.connect(self._generate_report)
        layout.addWidget(self.generate_button, 0, 4)
        
        # Preview label
        layout.addWidget(QLabel("Preview:"), 1, 0)
        self.preview_label = QLineEdit()
        self.preview_label.setReadOnly(True)
        self.preview_label.setPlaceholderText("Date range preview will appear here")
        layout.addWidget(self.preview_label, 1, 1, 1, 4)
        
        # Status/error label
        self.date_status_label = QLabel("")
        self.date_status_label.setStyleSheet("color: red;")
        layout.addWidget(self.date_status_label, 2, 0, 1, 5)
        
        # Connect date change signals
        self.start_date_picker.dateChanged.connect(self._update_date_preview)
        self.end_date_picker.dateChanged.connect(self._update_date_preview)
        
        # Initial preview update
        self._update_date_preview()
        
        return group_box
    
    def _create_report_display_section(self):
        """Create the report display section"""
        group_box = QGroupBox("Report Content")
        layout = QVBoxLayout(group_box)
        
        # Choose display widget based on availability
        if WEB_ENGINE_AVAILABLE:
            self.report_display = QWebEngineView()
            self.report_display.setHtml("<div style='padding: 20px; text-align: center; color: #666;'>No report generated yet. Select a date range and click 'Generate Report' to view content.</div>")
        else:
            self.report_display = QTextEdit()
            self.report_display.setReadOnly(True)
            self.report_display.setHtml("<div style='padding: 20px; text-align: center; color: #666;'>No report generated yet. Select a date range and click 'Generate Report' to view content.</div>")
        
        layout.addWidget(self.report_display)
        
        return group_box
    
    def _create_export_section(self):
        """Create the export section"""
        group_box = QGroupBox("Export Options")
        layout = QHBoxLayout(group_box)
        
        # Export buttons - CHANGED: Save as TXT instead of HTML
        self.export_txt_button = QPushButton("Save as TXT")
        self.export_txt_button.clicked.connect(self._export_txt)
        
        self.open_browser_button = QPushButton("Open in Browser")
        self.open_browser_button.clicked.connect(self._open_in_browser)
        
        # Add buttons to layout
        layout.addWidget(self.export_txt_button)
        layout.addWidget(self.open_browser_button)
        layout.addStretch()  # Push buttons to the left
        
        return group_box
    
    def _update_date_preview(self):
        """Update the date range preview"""
        start_date = self.start_date_picker.date().toPyDate()
        end_date = self.end_date_picker.date().toPyDate()
        
        # Validate dates
        if end_date < start_date:
            self.date_status_label.setText("End date cannot be earlier than start date.")
            self.generate_button.setEnabled(False)
            return
        else:
            self.date_status_label.setText("")
            self.generate_button.setEnabled(True)
        
        # Format the date range
        if start_date.month == end_date.month and start_date.year == end_date.year:
            date_range_formatted = f"{start_date.day}-{end_date.day} {start_date.strftime('%B')} {start_date.year}"
        else:
            date_range_formatted = f"{start_date.day} {start_date.strftime('%B')} - {end_date.day} {end_date.strftime('%B')} {end_date.year}"
        
        # Update preview
        self.preview_label.setText(date_range_formatted)
        
        # Store the current date range
        self.current_date_range = date_range_formatted
    
    def _generate_report(self):
        """Generate the weekly report"""
        # Validate date range
        if hasattr(self, 'current_date_range'):
            date_range_str = self.current_date_range
        else:
            self.status_label.setText("Please select a valid date range")
            return
        
        # Update UI for loading state
        self._update_ui_state(True)
        self.status_label.setText(f"Generating combined MFA + GSN VS AD report for: {date_range_str}")
        
        # Start extraction in a separate thread using the combined method
        self.extraction_thread = CombinedWeeklyReportExtractionThread(self.extractor, date_range_str)
        self.extraction_thread.progress.connect(self._update_progress)
        self.extraction_thread.finished.connect(self._on_extraction_finished)
        self.extraction_thread.start()
    
    def _update_progress(self, message):
        """Update progress message"""
        self.status_label.setText(message)
    
    def _on_extraction_finished(self, success, data, error_message):
        """Handle extraction completion"""
        # Update UI state
        self._update_ui_state(False)
        
        if success and data:
            # Generate HTML and display
            try:
                # Check if this is combined data (dict) or regular data (list)
                if isinstance(data, dict):
                    # This is combined MFA + GSN VS AD data
                    html_content = self.extractor.generate_complete_html(data, self.current_date_range)
                    
                    # Count total rows for status message
                    mfa_count = len(data.get('mfa_data', []))
                    gsn_count = len(data.get('gsn_vs_ad_data', []))
                    total_rows = mfa_count + gsn_count
                    
                    # Create detailed status message
                    status_parts = []
                    if data.get('mfa_success', False):
                        status_parts.append(f"MFA: {mfa_count} rows")
                    if data.get('gsn_vs_ad_success', False):
                        status_parts.append(f"GSN VS AD: {gsn_count} rows")
                    
                    status_message = f"Report generated successfully ({', '.join(status_parts)})"
                    
                else:
                    # This is regular MFA-only data (backward compatibility)
                    html_content = self.extractor.generate_complete_html(data, self.current_date_range)
                    status_message = f"Report generated successfully ({len(data)} rows)"
                
                # Display the HTML content
                if WEB_ENGINE_AVAILABLE:
                    self.report_display.setHtml(html_content)
                else:
                    # For QTextEdit, we need to simplify the HTML a bit
                    simplified_html = self._simplify_html_for_text_edit(html_content)
                    self.report_display.setHtml(simplified_html)
                
                # Store the HTML for export
                self.current_html = html_content
                # Store the raw data for TXT export
                self.current_data = data
                
                # Update status
                self.status_label.setText(status_message)
                
                # Enable export buttons
                self.export_txt_button.setEnabled(True)
                self.open_browser_button.setEnabled(True)
                
            except Exception as e:
                self.status_label.setText(f"Error generating HTML: {str(e)}")
                QMessageBox.warning(self, "Generation Error", f"Failed to generate report HTML:\n{str(e)}")
        else:
            # Show error
            error_msg = error_message or "Unknown error occurred"
            self.status_label.setText(f"Failed to generate report: {error_msg}")
            QMessageBox.warning(self, "Extraction Error", f"Failed to extract report data:\n{error_msg}")
        
        # Clean up thread
        self.extraction_thread = None
    
    def _simplify_html_for_text_edit(self, html_content):
        """Simplify HTML content for QTextEdit display"""
        # QTextEdit has limited CSS support, so we need to simplify
        # This is a basic implementation - you could make it more sophisticated
        simplified = html_content.replace('class="pending"', 'style="background-color: #ffeb9c; color: #9c5700;"')
        simplified = simplified.replace('class="completed"', 'style="background-color: #c6efce; color: #006100;"')
        simplified = simplified.replace('class="section-header"', 'style="background-color: #ddebf7; font-weight: bold;"')
        return simplified
    
    def _update_ui_state(self, loading):
        """Update UI state based on loading status"""
        self.generate_button.setEnabled(not loading)
        self.start_date_picker.setEnabled(not loading)
        self.end_date_picker.setEnabled(not loading)
        self.progress_bar.setVisible(loading)
        
        if not loading:
            # Reset export buttons if not loading
            self.export_txt_button.setEnabled(hasattr(self, 'current_data'))
            self.open_browser_button.setEnabled(hasattr(self, 'current_html'))
    
    def _export_txt(self):
        if not hasattr(self, 'current_data'):
            QMessageBox.information(self, "No Report", "Please generate a report first.")
            return
        
        try:
            import os  # Import the os module
            
            # Hardcoded path to the desired folder
            hardcoded_folder = r"C:\Users\haowerwu\OneDrive - DPDHL\Documents\weeklyreportlog"
            
            # Create filename with timestamp to avoid conflicts
            timestamp = self._get_current_timestamp().replace(":", "-").replace(" ", "_")
            safe_date_range = self.current_date_range.replace(' ', '_').replace('-', '_')
            filename = f"weekly_report_{safe_date_range}_{timestamp}.txt"
            
            # Full file path in the hardcoded folder
            file_path = os.path.join(hardcoded_folder, filename)
            
            # Convert the data to plain text format
            txt_content = self._convert_data_to_txt(self.current_data, self.current_date_range)
            
            # Save the TXT file directly to the hardcoded folder
            with open(file_path, 'w', encoding='utf-8') as f:
                f.write(txt_content)
            
            QMessageBox.information(self, "Export Successful", f"Report saved to:\n{file_path}")
            
        except Exception as e:
            QMessageBox.critical(self, "Export Error", f"Error saving file:\n{str(e)}")
    
    def _convert_data_to_txt(self, data, date_range_str):
        """
        Convert the extracted data to HTML table format for TXT file.
        
        Args:
            data: The extracted data (dict for combined data, list for regular data)
            date_range_str: Date range string
            
        Returns:
            str: HTML table content with all formatting preserved
        """
        # Get the complete HTML table from the extractor
        if isinstance(data, dict):
            # This is combined data - generate the combined HTML table
            html_table = self.extractor.generate_combined_html_table(data)
        else:
            # This is regular MFA-only data
            html_table = self.extractor.generate_html_table(data)
        
        # Add starting paragraphs
        txt_content = []
        txt_content.append(f'<p class="editor-paragraph"><b>{date_range_str} Weekly Report</b></p><br>')
        txt_content.append('<p class="editor-paragraph"><b>MFA &amp; AD/EDS<br></b></p>')
        txt_content.append("")
        
        # Add the complete HTML table with all styling
        txt_content.append(html_table)
        
        return "\n".join(txt_content)
    
    def _get_current_timestamp(self):
        """Get current timestamp as string"""
        import datetime
        return datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    
    def _open_in_browser(self):
        """Open the current report in the default browser"""
        if not hasattr(self, 'current_html'):
            QMessageBox.information(self, "No Report", "Please generate a report first.")
            return
        
        try:
            import tempfile
            import webbrowser
            import os
            
            # Create a temporary HTML file
            with tempfile.NamedTemporaryFile(mode='w', suffix='.html', delete=False, encoding='utf-8') as f:
                f.write(self.current_html)
                temp_path = f.name
            
            # Open in browser
            webbrowser.open('file://' + os.path.abspath(temp_path))
            self.status_label.setText("Report opened in browser")
            
            # Note: We don't delete the temp file immediately since the browser needs time to load it
            # The OS will clean it up eventually
            
        except Exception as e:
            QMessageBox.critical(self, "Browser Error", f"Error opening in browser:\n{str(e)}")