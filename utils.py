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
                "address", "business_name", "last_recharge_date", "location_of_cylinders", 
                "manufacturer", "model", "serial", "size", 
            ],
            "Fire Hoses": [
                "address", "business_name", "location", "length", 
                "size", "nozzle", "status", "next_ht", "notes"
            ],
            "Fire Hydrants": [
                "address", "business_name", "hydrant_number", "make", "model",
                "color", "shutoff_location", "type"
            ], 
            "Backflows" : [
                "address", "name_of_premise", "manufacturer", "model_number", 
                "serial_number", "type", "size", "location"
            ], 
            "Extinguishers": [
                "address", "business_name", "location", "size", "brand", "serial_number", 
                "manufacture_date", "next_service_date", "comments"
            ], 
            "Fire Pumps": [
                "address", "business_name", "location", "system", "water_supply_source",
                "pump_manufacturer", "pump_model", "controller_manufacturer", "controller_model", "type", "power"
            ], 
            "Smoke Alarms" : [
                "address", "business_name", "device", "location", "remarks"
            ], 
            "Indicator Valves": [
                "address", "business_name", "location"
            ], 
            "Emergency Lights": [
                "address", "business_name", "location", "unit_type", "battery_size",
                "battery #", "battery_date", "voltage / size", "comments"
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
            'indicator_valves': [
                "post indicator valve inspection"
            ], 
            'emergency_lighting': [
                'unit emergency lighting test'
            ], 
            'emergency_lighting_extinguisher': [
                'unit emergency lighting /'
            ], 
            'extinguishers': [
                'extinguisher test & inspection'
            ]
        }
        
        for method_name, keywords in method_keywords.items():
            if any(keyword in text_lower for keyword in keywords):
                return method_name
        
        return None
