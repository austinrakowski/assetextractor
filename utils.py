import os
import re
from openpyxl import Workbook
import pprint
from docx import Document
import inspect 

class AssetExtractorUtils:
    def __init__(self, directory_path):
        self.directory_path = directory_path
        self.workbook = None
        self.create_workbook()

    def create_workbook(self):
        
        self.workbook = Workbook()

        sheets_with_headers = {
            "Fixed Extinguishing Systems": [
                "address", "business_name", "asset_type", "variant", "last_recharge_date", "location_of_cylinders", 
                "manufacturer", "model_number", "serial_number", "size", 
            ],
            "Fire Hoses": [
                "address", "business_name", "location", "length", 
                "size", "nozzle", "status", "next_ht", "notes"
            ],
            "Fire Hydrants": [
                "address", "business_name", "hydrant_number", "location", "make", "model",
                "color", "shutoff_location", "type"
            ], 
            "Backflows" : [
                "name_of_premise", "address", "asset_type", "variant", "manufacturer", "model_number", 
                "serial_number", "size", "location"
            ], 
            "Extinguishers": [
                "address", "business_name", "location", "size", "brand", "serial_number", 
                "manufacture_date", "next_service_date", "comments"
            ], 
            "Fire Pumps": [
                "address", "business_name", "asset_type", "variant", "location", "system", 
                "water_supply_source", "pump_manufacturer", "controller_manufacturer", "controller_model" 
            ], 
            "Smoke Alarms" : [
                "address", "business_name", "device", "location", "remarks"
            ],  
            "Emergency Lights": [
                "address", "business_name", "type", "variant", "location", "battery_size",
                "battery #", "battery_date", "voltage / size", "comments"
            ], 
            "Special Suppression" : [
                "address", "business_name", "asset_type", "variant", "make", "model"
            ], 
            "Alarm Systems" : [
                'address', 'ref', 'business_name', 'manufacturere', 'model_number', 
                'fire_signal_receiving_centre', 'ulc_serial_number'
            ], 
            "Alarm System Devices": [
                'system', 'asset_type', 'variant', 'zone_circuit_number' 
            ], 
            "Sprinkler Systems": [
                'address', 'business_name', 'ref', 'asset_type', 'variant'
            ],
            "Sprinkler System Devices": [
                "system", 'asset_type', 'variant', 'size'
            ]
        }
       
        default_sheet = self.workbook.active
        self.workbook.remove(default_sheet)

        for sheet_name, headers in sheets_with_headers.items():
            sheet = self.workbook.create_sheet(title=sheet_name)
            sheet.append(headers)

        self.workbook.save("assets.xlsx")

    def update_workbook(self, sheet_name, data):
        
        sheet = self.workbook[sheet_name]

        for row in data:
            sheet.append(row)

        self.workbook.save("assets.xlsx")

    def get_document_text(self, file_path): 
        doc = Document(file_path)
        core_props = doc.core_properties
        full_text = []
    
        for paragraph in doc.paragraphs:
            full_text.append(paragraph.text)
    
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    full_text.append(cell.text)
        
        return '\n'.join(full_text)

    def get_document_tables(self, file_path): 
        doc = Document(file_path)

        return doc.tables
    
    def get_docx_files(self):
        """Get list of PDF files in directory."""
        try:
            return [f for f in os.listdir(self.directory_path) if f.lower().endswith('.docx')]
        except OSError as e:
            print(f"Error reading directory {self.directory_path}: {e}")
            return []

    def _headers_match(self, actual_headers, target_headers):
        """Check if actual headers match target headers (allowing for some flexibility)"""
        if len(actual_headers) < len(target_headers):
            return False
    
        key_matches = 0
        for i, target in enumerate(target_headers):
            if i < len(actual_headers):
                # Normalize text for comparison
                actual_normalized = actual_headers[i].replace('\n', ' ').replace('\r', '').strip()
                target_normalized = target.replace('\n', ' ').replace('\r', '').strip()
                
                if actual_normalized.lower() == target_normalized.lower():
                    key_matches += 1
                elif target in ["A", "B", "C", "D", "E", "F", "G"] and actual_normalized == target:
                    key_matches += 1
        
        return key_matches >= 7  
    
    def find_header_row(self, table, target_text):
        """
        Find the row index that contains the target header text.
        Returns (row_index, column_index) if found, otherwise (None, None)
        """
        
        for row_idx, row in enumerate(table.rows):
            for col_idx, cell in enumerate(row.cells):
                text = cell.text.strip()
                val += text
                if target_text in cell.text.strip():
                    return row_idx, col_idx
        
        return None, None

    def get_extraction_method(self, text):
        """Determine which extraction method to use based on its content."""
        text_lower = text.lower()
        
        method_keywords = {
            'fixed_extinguishing_systems': [
                'inspection, testing and maintenance report for fixed extinguishing systems',
                'location of system cylinders'
            ],
            'fire_hoses': [
                'fire hose test and inspection'
            ], 
            'fire_hydrants' : [
                'fire hydrant inspection & testing'
            ], 
            'backflows': [
                'location of backflow preventer'
            ], 
            'fire_pumps': [
                'fire pump annual performance tests', 
                'pump has a prv installed'
            ], 
            'smoke_alarms' : [
                'smoke alarm device record'
            ],  
            'emergency_lighting': [
                'unit emergency lighting test'
            ], 
            'emergency_lighting_extinguisher': [
                'unit emergency lighting /'
            ], 
            'extinguishers': [
                'extinguisher test & inspection'
            ], 
            'special_suppression': [
                'report for special fire suppression system', 
                'novec 1230'
            ], 
            'alarm_system_devices': [
                'nbc', 'provides single-stage operation'
            ], 
            'sprinkler_systems': [
                'is the fdc check valve free of leaks?'
            ]
        }
        
        for method_name, keywords in method_keywords.items():
            if any(keyword in text_lower for keyword in keywords):
                return method_name
        
        return None

    
