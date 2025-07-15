import os
import pprint
import datetime
from templates import AssetTemplateMethods
from utils import AssetExtractorUtils

class AssetExtractor(AssetExtractorUtils, AssetTemplateMethods):
    
    def __init__(self, directory_path):
        AssetExtractorUtils.__init__(self, directory_path)
        AssetTemplateMethods.__init__(self)
        self.processed_files = []
        self.failed_files = []
        self.log_file = f"extraction_errors_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.log"
        self._init_log_file()
    
    def _init_log_file(self):
        """Initialize the log file with header."""
        with open(self.log_file, 'w') as f:
            f.write(f"Asset Extraction Error Log\n")
            f.write(f"Started: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
            f.write(f"Directory: {self.directory_path}\n")
            f.write("-" * 80 + "\n\n")
    
    def _log_error(self, filename, error, error_type="PROCESSING_ERROR"):
        """Log error details to file."""
        timestamp = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        with open(self.log_file, 'a') as f:
            f.write(f"[{timestamp}] {error_type}\n")
            f.write(f"File: {filename}\n")
            f.write(f"Error: {str(error)}\n")
            f.write(f"Error Type: {type(error).__name__}\n")
            f.write("-" * 40 + "\n\n")
    
    def process_file(self, filename):
        """Process a single file."""
        file_path = os.path.join(self.directory_path, filename)
        
        if "Old Customers" in file_path: 
            os.remove(file_path)
        
        try:
            print(f"Processing: {filename}")
        
            text = self.get_document_text(file_path)
            # Determine which extraction method to use
            method_name = self.get_extraction_method(text)
            
            if not method_name:
                error_msg = "No matching extraction method found"
                print(f"Error processing {filename}: {error_msg}")
                self._log_error(filename, error_msg, "NO_METHOD_FOUND")
                self.failed_files.append(filename)
                return None
            
            method = getattr(self, method_name)
            if method_name in ["fire_pumps", "alarm_system_devices"]: #pumps need text for non-table checkboxes, alarms need it for varying headers
                result = method(file_path, text)
            else: 
                result = method(file_path)
            
            print(f"Successfully extracted data from {filename} using {method_name}")
            self.processed_files.append(filename)
            
            # Delete file after successful processing
            try:
                os.remove(file_path)
                print(f"Deleted: {filename}")
            except OSError as e:
                self._log_error(filename, e, "FILE_DELETION_ERROR")
                print(f"Warning: Could not delete {filename}: {e}")
            
            return result
            
        except Exception as e:
            print(f"Error processing {filename}: {e}")
            self._log_error(filename, e)
            self.failed_files.append(filename)
            return None

    def process_all_files(self):
        """Process all .docx files in the directory."""
        files = self.get_docx_files()
        
        if not files:
            print("No .docx files found in the directory.")
            return
        
        total_files = len(files)
        print(f"Found {total_files} .docx file(s) to process")
        
        for i, filename in enumerate(files, 1):
            print(f"\n[{i}/{total_files}] ", end="")
            result = self.process_file(filename)
        
        # Log summary to file
        self._log_summary()
        
        print(f"\nProcessing Summary:")
        print(f"Successfully processed: {len(self.processed_files)} files")
        print(f"Failed to process: {len(self.failed_files)} files")
        print(f"Error log saved to: {self.log_file}")
        
        if self.failed_files:
            print(f"Failed files: {', '.join(self.failed_files)}")
    
    def _log_summary(self):
        """Log processing summary to file."""
        with open(self.log_file, 'a') as f:
            f.write(f"\nPROCESSING SUMMARY\n")
            f.write(f"Completed: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
            f.write(f"Total files found: {len(self.processed_files) + len(self.failed_files)}\n")
            f.write(f"Successfully processed: {len(self.processed_files)}\n")
            f.write(f"Failed to process: {len(self.failed_files)}\n")
            if self.failed_files:
                f.write(f"Failed files:\n")
                for failed_file in self.failed_files:
                    f.write(f"  - {failed_file}\n")
            f.write("-" * 80 + "\n")

    def get_stats(self):
        """Get processing statistics."""
        return {
            'total_files': len(self.processed_files) + len(self.failed_files),
            'processed_files': len(self.processed_files),
            'failed_files': len(self.failed_files),
            'success_rate': len(self.processed_files) / (len(self.processed_files) + len(self.failed_files)) * 100 if (self.processed_files or self.failed_files) else 0,
            'log_file': self.log_file
        }

if __name__ == "__main__": 
    directory_path = "/Users/austinrakowski/dev/random/firstresponse/frdocs"
    extractor = AssetExtractor(directory_path)
    extractor.process_all_files()