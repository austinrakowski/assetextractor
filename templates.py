import re
import pprint


class AssetTemplateMethods:
    """Methods for extracting data from different types of assets"""

    def fixed_extinguishing_systems(self, file_path):
        """Extract data from Fixed Extinguishing Systems Template"""
        
        tables = self.get_document_tables(file_path)
        data = {}
        
        field_mappings = {
            "Business Name:": "Business_Name",
            "Address:": "Address", 
            "City:": "City",
            "Last Recharge Date:": "Last_Recharge_Date",
            "Location of System Cylinders:": "Location_of_System_Cylinders",
            "Manufacturer:": "Manufacturer",
            "Model #:": "Model_Number",
            "Serial #:": "Serial_Number",
            "Size:": "Size"
        }
        
        for table in tables:
            for row in table.rows:
                cells = [cell.text.strip() for cell in row.cells]
                
                for i, cell_text in enumerate(cells):
                    if cell_text in field_mappings and i + 1 < len(cells):
                        data[field_mappings[cell_text]] = cells[i + 1]
        
        self.update_workbook("Fixed Extinguishing Systems", [[
            f"{data.get('Address', '')} {data.get('City', '')}",
            data.get('Business_Name', ''),
            data.get('Last_Recharge_Date', ''), 
            data.get('Location_of_System_Cylinders', ''), 
            data.get('Manufacturer', ''), 
            data.get('Model_Number', ''), 
            data.get('Serial_Number', ''), 
            data.get('Size', '')
        ]])

    def fire_hoses(self, file_path):
        """Extract data from Fire Hose Test and Inspection Template"""
        
        tables = self.get_document_tables(file_path)
        extracted_data = {}
        hose_data = []
        
        field_mappings = {
            "Business Name:": "Business_Name",
            "Address:": "Address", 
            "City:": "City",
        }
        
        for table in tables:
            for row in table.rows:
                cells = [cell.text.strip() for cell in row.cells]
                
                # Extract header fields
                for i, cell_text in enumerate(cells):
                    if cell_text in field_mappings and i + 1 < len(cells):
                        extracted_data[field_mappings[cell_text]] = cells[i + 1]
                
                # Extract hose data rows
                if len(cells) >= 8 and cells[0].isdigit():
                    address_city = f"{extracted_data.get('Address', '')} {extracted_data.get('City', '')}"
                    business_name = extracted_data.get("Business_Name", '')
                    
                    hose_data.append([
                        address_city, business_name, cells[1], cells[2], 
                        cells[3], cells[4], cells[5], cells[6], 
                        cells[7] if len(cells) > 7 else ""
                    ])

        self.update_workbook("Fire Hoses", hose_data)

    def fire_hydrants(self, file_path):
        """Extract data from Fire Hydrant Inspection & Testing Template"""
        
        tables = self.get_document_tables(file_path)
        data = {}
        
        field_mappings = {
            "Business Name:": "Business_Name",
            "Address:": "Address", 
            "City:": "City",
            "HYDRANT #:": "Hydrant_Number",
            "Make:": "Make",
            "Model:": "Model",
            "Color:": "Color",
            "Hydrant Shut-Off Location:": "Shutoff_Location"
        }
        
        checkbox_mappings = {
            "PRIVATE": "Type",
            "PUBLIC": "Type"
        }
        
        for table in tables:
            for row in table.rows:
                cells = [cell.text.strip() for cell in row.cells]
                
                for i, cell_text in enumerate(cells):
                    if cell_text in field_mappings and i + 1 < len(cells):
                        data[field_mappings[cell_text]] = cells[i + 1]
                    
                    # Handle checkboxes
                    for checkbox_text, category in checkbox_mappings.items():
                        if checkbox_text in cell_text and "☒" in cell_text:
                            data[category] = checkbox_text
        
        self.update_workbook("Fire Hydrants", [[
            f"{data.get('Address', '')} {data.get('City', '')}",
            data.get('Business_Name', ''),
            data.get('Hydrant_Number', ''),
            data.get('Make', ''),
            data.get('Model', ''),
            data.get('Color', ''),
            data.get('Shutoff_Location', ''),
            data.get('Type', '')
        ]])

        return None
            
    def backflows(self, file_path):
        """Extract data from Backflow Prevention Assembly Test Report Template"""
        
        tables = self.get_document_tables(file_path)
        data = {}
        
        field_mappings = {
            "Name of Premise:": "Name_of_Premise",
            "Service Address:": "Service_Address", 
            "Postal Code:": "Postal_Code",
            "Location of Backflow Preventer:": "Location"
        }    
        
        for table in tables:
            for row_idx, row in enumerate(table.rows):
                cells = [cell.text.strip() for cell in row.cells]
                
                for i, cell_text in enumerate(cells):
                    if cell_text in field_mappings and i + 1 < len(cells):
                        data[field_mappings[cell_text]] = cells[i + 1]
                    
                    # When we find the Manufacturer label, get assembly values from previous row
                    if cell_text == "Manufacturer" and row_idx > 0:
                        prev_row_cells = [cell.text.strip() for cell in table.rows[row_idx - 1].cells]
        
                        current_row_labels = cells
                        for j, label in enumerate(current_row_labels):
                            if j < len(prev_row_cells):
                                if label == "Manufacturer":
                                    data["Manufacturer"] = prev_row_cells[j]
                                elif label == "Model #":
                                    data["Model_Number"] = prev_row_cells[j]
                                elif label == "Serial #":
                                    data["Serial_Number"] = prev_row_cells[j]
                                elif label == "Type":
                                    data["Type"] = prev_row_cells[j]
                                elif label == "Size":
                                    data["Size"] = prev_row_cells[j]
                        break
        
        self.update_workbook("Backflows", [[
            f"{data.get('Service_Address', '')} {data.get('Postal_Code', '')}",
            data.get('Name_of_Premise', ''),
            data.get('Manufacturer', ''),
            data.get('Model_Number', ''),
            data.get('Serial_Number', ''),
            data.get('Type', ''),
            data.get('Size', ''),
            data.get('Location_of_Backflow_Preventer', '')
        ]])
    
    def extinguishers(self, file_path):
        """Extract data from Fire Extinguisher Test & Inspection Template"""
        
        tables = self.get_document_tables(file_path)
        extracted_data = {}
        extinguisher_data = []
        
        field_mappings = {
            "Business Name:": "Business_Name",
            "Address:": "Address", 
            "City:": "City",
        }
        
        for table in tables:
            for row in table.rows:
                cells = [cell.text.strip() for cell in row.cells]
                
                # Extract header fields
                for i, cell_text in enumerate(cells):
                    if cell_text in field_mappings and i + 1 < len(cells):
                        extracted_data[field_mappings[cell_text]] = cells[i + 1]
                
            
                if (len(cells) >= 6 and 
                    cells[0] and 
                    cells[0] not in ["Location", "Mfg. Date", "Service Date", "Business Name:"] and
                    not any(header in cells[0] for header in ["Column Legend", "Dates", "Major Service"]) and
                    not cells[0].endswith(":")): 
                    
                    address_city = f"{extracted_data.get('Address', '')} {extracted_data.get('City', '')}"
                    business_name = extracted_data.get("Business_Name", '')
                    
                    extinguisher_data.append([
                        address_city, business_name, cells[0], cells[1], 
                        cells[2], cells[3], cells[4], cells[5],
                        cells[6] if len(cells) > 6 else ""
                    ])
        
        self.update_workbook("Extinguishers", extinguisher_data)
        return None
        
