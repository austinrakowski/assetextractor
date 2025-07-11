import os
import pprint
from templates import AssetTemplateMethods
from utils import AssetExtractorUtils

class AssetExtractor(AssetExtractorUtils, AssetTemplateMethods):
    
    def __init__(self, directory_path):
        super().__init__(directory_path)
        self.processed_files = []
        self.failed_files = []
    
    def process_file(self, filename):
        """Process a single file."""
        file_path = os.path.join(self.directory_path, filename)
        
        try:
            print(f"Processing: {filename}")

            text = self.get_document_text(file_path)
            # Determine which extraction method to use
            method_name = self.get_extraction_method(text)
            
            method = getattr(self, method_name)
            result = method(file_path)
            
            print(f"Successfully extracted data from {filename} using {method_name}")
            self.processed_files.append(filename)
            return result
            
        except Exception as e:
            print(f"Error processing {filename}: {e}")
            self.failed_files.append(filename)
            return None

    def process_all_pdfs(self):
        """Process all PDF files in the directory."""
        files = self.get_docx_files()
        
        if not files:
            print("No .docx files found in the directory.")
            return
        
        print(f"Found {len(files)} PDF file(s) to process")
        
        for filename in files:
            result = self.process_file(filename)
        
        # Print summary
        print(f"\nProcessing Summary:")
        print(f"Successfully processed: {len(self.processed_files)} files")
        print(f"Failed to process: {len(self.failed_files)} files")
        
        if self.failed_files:
            print(f"Failed files: {', '.join(self.failed_files)}")

    def debug_file(self, filename):
        """Debug a specific file to see its content and auto-detected method."""
        file_path = os.path.join(self.directory_path, filename)
        
        try:
            text = self.extract_text_from_pdf(file_path)
            detected_method = self.get_extraction_method(text)
            
            print(f"File: {filename}")
            print(f"Detected method: {detected_method}")
            print("-" * 50)
            print("Text")
            print(text)
            print("-" * 50)
            
        except Exception as e:
            print(f"Error debugging {filename}: {e}")

    def get_stats(self):
        """Get processing statistics."""
        return {
            'total_files': len(self.processed_files) + len(self.failed_files),
            'processed_files': len(self.processed_files),
            'failed_files': len(self.failed_files),
            'success_rate': len(self.processed_files) / (len(self.processed_files) + len(self.failed_files)) * 100 if (self.processed_files or self.failed_files) else 0
        }

# Example usage
if __name__ == "__main__":
    # Example of how to use the AssetExtractor
    directory_path = "/Users/austinrakowski/dev/random/firstresponse/frdocs"
    extractor = AssetExtractor(directory_path)
    # extractor.debug_file("/Users/austinrakowski/dev/random/firstresponse/frdocs/SS Report Jun 2025 (WET SYSTEM).pdf")
    extractor.process_all_pdfs()