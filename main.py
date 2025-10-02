import os
import pprint
import datetime
from templates import AssetTemplateMethods
from utils import AssetExtractorUtils
import time
from threading import Lock, Semaphore
from concurrent.futures import ThreadPoolExecutor

class AssetExtractor(AssetExtractorUtils, AssetTemplateMethods):
    
    def __init__(self, directory_path):
        AssetExtractorUtils.__init__(self, directory_path)
        AssetTemplateMethods.__init__(self)
        self.processed_files = []
        self.failed_files = []
        self.log_file = f"extraction_errors_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.log"
        
        self.file_lock = Lock()
        self.workbook_lock = Lock()
        
        self.rate_limiter = RateLimiter(max_requests_per_minute=450)
        
        self._init_log_file()
    
    def _init_log_file(self):
        with open(self.log_file, 'w') as f:
            f.write(f"Asset Extraction Error Log\n")
            f.write(f"Started: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
            f.write(f"Directory: {self.directory_path}\n")
            f.write("-" * 80 + "\n\n")
    
    def _log_error(self, filename, error, error_type="PROCESSING_ERROR"):
        timestamp = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        with self.file_lock:
            with open(self.log_file, 'a') as f:
                f.write(f"[{timestamp}] {error_type}\n")
                f.write(f"File: {filename}\n")
                f.write(f"Error: {str(error)}\n")
                f.write(f"Error Type: {type(error).__name__}\n")
    
    def process_file(self, filename):
        """Process a single file."""
        file_path = os.path.join(self.directory_path, filename)
        
        if "Old Customers" in file_path: 
            os.remove(file_path)
            return None
        
        try:
            print(f"Processing: {filename}")
        
            text = self.get_document_text(file_path)
            # Determine which extraction method to use
            method_name = self.get_extraction_method(text)
            
            if not method_name:
                return None
            
            method = getattr(self, method_name)
            if method_name in ["fire_pumps", "alarm_system_devices"]: #pumps need text for non-table checkboxes, alarms need it for varying headers
                result = method(file_path, text)
            else: 
                result = method(file_path)
            
            print(f"Successfully extracted data from {filename} using {method_name}")
            with self.file_lock: 
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
        
        self._log_summary()
        
        print(f"\nProcessing Summary:")
        print(f"Successfully processed: {len(self.processed_files)} files")
        print(f"Failed to process: {len(self.failed_files)} files")
        print(f"Error log saved to: {self.log_file}")
        
        if self.failed_files:
            print(f"Failed files: {', '.join(self.failed_files)}")
    
    def _log_summary(self):
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
        return {
            'total_files': len(self.processed_files) + len(self.failed_files),
            'processed_files': len(self.processed_files),
            'failed_files': len(self.failed_files),
            'success_rate': len(self.processed_files) / (len(self.processed_files) + len(self.failed_files)) * 100 if (self.processed_files or self.failed_files) else 0,
            'log_file': self.log_file
        }
    

class RateLimiter:
    def __init__(self, max_requests_per_minute=500):
        self.max_requests = max_requests_per_minute
        self.requests = []
        self.lock = Lock()
    
    def wait_if_needed(self):
        with self.lock:
            now = time.time()
            # Remove requests older than 1 minute
            self.requests = [req_time for req_time in self.requests if now - req_time < 60]
            
            if len(self.requests) >= self.max_requests:
                # Wait until oldest request is 60 seconds old
                sleep_time = 60 - (now - self.requests[0]) + 0.1
                if sleep_time > 0:
                    print(f"Rate limit approaching. Waiting {sleep_time:.1f}seconds")
                    time.sleep(sleep_time)
                    now = time.time()
                    self.requests = [req_time for req_time in self.requests if now - req_time < 60]
            
            self.requests.append(now)

if __name__ == "__main__": 
    from wakepy import keep
    with keep.running():
        directory_path = "/Users/austinrakowski/dev/random/firstresponse/frdocs"
        extractor = AssetExtractor(directory_path)
        extractor.process_all_files()