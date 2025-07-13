import re
import pprint
import random
import shortuuid

class AssetTemplateMethods:
    """Methods for extracting data from different types of assets"""

    def __init__(self):

        # 0 for individual asset, 1 for list of assets 
        self.method_type = {
            "fixed_extinguishing_systems" : {"type" : 0, "devices": set()},
            "fire_hoses" : {"type": 1, "devices": set()},
            "fire_hydrants" : {"type": 0, "devices": set()}, 
            "backflows" : {"type": 0, "devices": set()}, 
            "extinguishers" : {"type": 1, "devices": set()}, 
            "fire_pumps" : {"type": 0, "devices": set()}, 
            "smoke_alarms" : {"type": 1, "devices": set()},
            "emergency_lighting": {"type": 1, "devices": set()},  
            "emergency_lighting_extinguisher": {"type" : 1, "devices": set()}, 
            "special_suppression": {"type" : 1, "devices": set()}, 
            "alarm_system_device": {"type" : 1, "devices": set()},
            "alarm_system": {"type" : 1, "devices": set()}, 
            "wet_systems": {"type" : 1, "devices": set()},
            "dry_systems": {"type" : 1, "devices": set()},

        }

        self.eml_mapping = {
            'SPU' : ('NFPA 101: Emergency Escape Lighting: System Type', 'Self-Powered Unit'), 
            'BP' : ('NFPA 101: Emergency Escape Lighting & Exit Signs', 'Battery Pack'), 
            'RH' : ('NFPA 101: Emergency Escape Lighting & Exit Signs', 'Remote Head'), 
            'EX' : ('NFPA 101: Emergency Escape Lighting: Exit Sign', 'Exit Sign'), 
            'COM' : ('NFPA 101: Emergency Escape Lighting & Exit Signs', 'Combo Unit')
        }

        self.alarm_mapping = {
        "M": ('NFPA 72: Fire Alarm System: Manual Pull Station (MPS)', 'Manual Pull Station'),
        "DS": ('NFPA 72: Fire Alarm System: Detector', 'Smoke Duct'),
        "B": ('NFPA 72: Sound and Intercom For Emergency Purposes: Audible/Visual', 'Bell'),
        "AD": ("NFPA 72: Fire Alarm System: Peripheral/Accessorie", 'Ancillary device'),
        "HT": ("NFPA 72: Fire Alarm System: Detector", "Heat Detector"),
        "RHT": ("NFPA 72: Fire Alarm System: Detector", "Heat Detector"),
        "S": ("NFPA 72: Fire Alarm System: Detector", "Smoke Detector"),
        "RI": ("NFPA 72: Smoke and Heat Alarm", "Remote Indicator Unit"),
        "SFD": ("NFPA 72: Smoke and Heat Alarm", "Supporting Field Device Monitor"),
        "FS": ("NFPA 72: Smoke and Heat Alarm", "Sprinkler Flow Switch"),
        "SS": ("NFPA 72: Smoke and Heat Alarm", "Sprinkler Supervisory Device"),
        "FM": ("NFPA 72: Smoke and Heat Alarm", "Isolation Module"),
        "H": ("NFPA 72: Sound and Intercom For Emergency Purposes: Audible/Visual", "Horn"),
        "V": ("NFPA 72: Sound and Intercom For Emergency Purposes: Audible/Visual", "Visual Warning Device"),
        "SP": ("NFPA 72: Sound and Intercom For Emergency Purposes: Audible/Visual", "Cone Type Speaker"),
        "HSP": ("NFPA 72: Sound and Intercom For Emergency Purposes: Audible/Visual", "Horn Type Speaker"),
        "ET": ("NFPA 72: Two-Way Emergency Communication System", "Emergency Telephone"),
        "EOL": ("NFPA 72: Fire Alarm System: Peripheral/Accessorie", "End of Line Resister (EOLR)"),
        "PZ": ("NFPA 72: Sound and Intercom For Emergency Purposes: Audible/Visual", "In-Suite Buzzer")
    }

        
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
            "Size:": "Size",
        }
        
        chemical_types = {
            "Wet Chemical": {
                "asset_type": "NFPA 17A: Special Hazard: Wet Chemical System (Cylinder)", 
                "variant": "Wet Chemical"
            },
            "Dry Chemical": {
                "asset_type": "NFPA 17: Special Hazard: Dry Chemical (Fire Extinguishing System)", 
                "variant": "Dry Chemical Extinguishing System"
            }
        }
        
        for table in tables:
            for row in table.rows:
                cells = [cell.text.strip() for cell in row.cells]
                
                for i, cell_text in enumerate(cells):
                    if cell_text in field_mappings and i + 1 < len(cells):
                        data[field_mappings[cell_text]] = cells[i + 1]
                    
                    # Handle checkboxes for chemical type
                    for checkbox_text in chemical_types.keys():
                        if checkbox_text in cell_text and "☒" in cell_text:
                            data['Chemical_Type'] = checkbox_text
                            data['asset_type'] = chemical_types[checkbox_text]['asset_type']
                            data['variant'] = chemical_types[checkbox_text]['variant']

        # Get values with defaults
        at = data.get('asset_type', 'Unknown')
        var = data.get('variant', 'Unknown')
        
        # Still capture asset if there is no serial number
        identifier = data.get("Serial_Number") or f'None - {random.randint(1000, 9999)}'
        if identifier not in self.method_type['fixed_extinguishing_systems']['devices']:
        
            self.update_workbook("Fixed Extinguishing Systems", [[
                f"{data.get('Address', '')} {data.get('City', '')}",
                data.get('Business_Name', ''),
                at, 
                var,
                data.get('Last_Recharge_Date', ''), 
                data.get('Location_of_System_Cylinders', ''), 
                data.get('Manufacturer', ''), 
                data.get('Model_Number', ''), 
                data.get('Serial_Number', ''), 
                data.get('Size', ''),
            ]])
        
            self.method_type['fixed_extinguishing_systems']['devices'].add(identifier)

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

                    identifier = f'{address_city} - {cells[1]}'
                    if identifier not in self.method_type['fire_hoses']['devices']:
                                    
                        hose_data.append([
                            address_city, business_name, cells[1], cells[2], 
                            cells[3], cells[4], cells[5], cells[6], 
                            cells[7] if len(cells) > 7 else ""
                        ])

                        self.method_type['fire_hoses']['devices'].add(identifier)
                                 
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
            "Hydrant Shut-Off Location:": "Shutoff_Location",
            "Date of Service:" : "Date", 
            "HYDRANT LOCATION:": "Hydrant_Location"
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
        
        identifier = f'{data.get('Business_Name', '')} - {data.get('Hydrant_Number')}'
        if identifier not in self.method_type['fire_hydrants']['devices']:

        
            self.update_workbook("Fire Hydrants", [[
                f"{data.get('Address', '')} {data.get('City', '')}",
                data.get('Business_Name', ''),
                data.get('Hydrant_Number', ''),
                data.get('Hydrant_Location', ''),
                data.get('Make', ''),
                data.get('Model', ''),
                data.get('Color', ''),
                data.get('Shutoff_Location', ''),
                data.get('Type', '')
            ]])

            self.method_type['fire_hydrants']['devices'].add(identifier)

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

        var_mapping = {
            'RP' : 'Reduced Pressure Zone', 
            'DCVA' : 'Double Check Valve Assembly', 
            'RPBA': 'Reduced Pressure Backflow Assembly'
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
        self.backflow_types.add(data.get('Type', ''))
        
        #still capture asset if there is no serial number
        identifier = data.get("Serial_Number") or f'None {random.randint(1000, 9999)}'
        if identifier not in self.method_type['backflows']['devices']:
        
            self.update_workbook("Backflows", [[
                data.get('Name_of_Premise', ''),
                f"{data.get('Service_Address', '')} {data.get('Postal_Code', '')}",
                "NFPA 25: Automatic Back-Flow Prevention", 
                var_mapping.get(data.get('type', '')),
                data.get('Manufacturer', ''),
                data.get('Model_Number', ''),
                data.get('Serial_Number', ''),
                data.get('Size', ''),
                data.get('Location_of_Backflow_Preventer', '')
            ]])

            self.method_type['backflows']['devices'].add(identifier)
        
        return None
        
    def extinguishers(self, file_path):
        """Extract data from Fire Extinguisher Test & Inspection Template"""
        
        tables = self.get_document_tables(file_path)
        extracted_data = {}
        extinguisher_data = []
        
        field_mappings = {
            "Business Name:": "Business_Name",
            "Address:": "Address", 
            "City:": "City",
            "Date of Service:": "Date"
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
                    identifier = cells[3] or f'None {random.randint(1000, 9999)}'
                    if identifier not in self.method_type['extinguishers']['devices']:
                    
                        extinguisher_data.append([
                            address_city, business_name, cells[0], cells[1], 
                            cells[2], cells[3], cells[4], cells[5],
                            cells[6] if len(cells) > 6 else ""
                        ])

                        self.method_type['extinguishers']['devices'].add(identifier)
                    
        self.update_workbook("Extinguishers", extinguisher_data)
        return None
    
    def fire_pumps(self, file_path, text):
        """Extract data from Fire Pump Annual Performance Tests Template"""
        
        tables = self.get_document_tables(file_path)
        data = {}
        
        field_mappings = {
            "Business Name:": "Business_Name",
            "Address:": "Address", 
            "City:": "City",
            "Location:": "Location",
            "System:": "System",
            "Water Supply Source:": "Water_Supply_Source",
            "Pump Manufacturer:": "Pump_Manufacturer",
            "Pump Model:": "Pump_Model",
            "Controller Manufacturer:": "Controller_Manufacturer",
            "Controller Model:": "Controller_Model", 
            "Date of Service:": "Date"
        }

        atv_mapping = { 
            'Electric' : 'NFPA 25: Fire Pump: Electric', 
            'Diesel' : 'NFPA 25: Fire Pump: Diesel', 
            'Steam' : 'NFPA 25: Fire Pump: Steam'
        }
        
        for table in tables:
            for row in table.rows:
                cells = [cell.text.strip() for cell in row.cells]
                
                for i, cell_text in enumerate(cells):
                    if cell_text in field_mappings and i + 1 < len(cells):
                        data[field_mappings[cell_text]] = cells[i + 1]
        
        pump_type = "Unknown"
        power_type = "Unknown"
        
        if "Centrifugal ☒" in text:
            pump_type = "Centrifugal"
        elif "Turbine ☒" in text:
            pump_type = "Turbine"
    
        if "Electric ☒" in text:
            power_type = "Electric"
        elif "Diesel ☒" in text:
            power_type = "Diesel"
        elif "Steam ☒" in text:
            power_type = "Steam"
        
        identifier = f'{data.get("Business_Name", "")} - {data.get("System", "")}'
        if identifier not in self.method_type['fire_pumps']["devices"]: 

            at = atv_mapping.get(power_type, 'Unknown')
            variant = f"{pump_type} Pump"
        
            self.update_workbook("Fire Pumps", [[
                f"{data.get('Address', '')} {data.get('City', '')}",
                data.get('Business_Name', ''),
                at, 
                variant,
                data.get('Location', ''),
                data.get('System', ''),
                data.get('Water_Supply_Source', ''),
                data.get('Pump_Manufacturer', ''),
                data.get('Pump_Model', ''),
                data.get('Controller_Manufacturer', ''),
                data.get('Controller_Model', ''),
            ]])

            self.method_type['fire_pumps']["devices"].add(identifier)

        return None

    def smoke_alarms(self, file_path):
        """Extract data from Inspection & Testing of Smoke Alarms Template"""
        
        tables = self.get_document_tables(file_path)
        extracted_data = {}
        alarm_data = []
        
        field_mappings = {
            "Business Name:": "Business_Name",
            "Address:": "Address", 
            "City:": "City",
            "Date of Service" : "Date"
        }
        
        for table in tables:
            for row in table.rows:
                cells = [cell.text.strip() for cell in row.cells]
                
                # Extract header fields
                for i, cell_text in enumerate(cells):
                    if cell_text in field_mappings and i + 1 < len(cells):
                        extracted_data[field_mappings[cell_text]] = cells[i + 1]
                
                # Extract alarm device data rows (skip header rows)
                if (len(cells) >= 5 and 
                    cells[0] and 
                    cells[0] not in ["Device", "Business Name:", "A.", "B.", "C."] and
                    not any(header in cells[0] for header in ["Page", "Correctly", "Requires"]) and
                    not cells[0].endswith(":")):
                    
                    address = f"{extracted_data.get('Address', '')} {extracted_data.get('City', '')}"
                    business_name = extracted_data.get("Business_Name", '')
                    identifier = f'{address} - {cells[0]} - {cells[1]}'

                    if identifier not in self.method_type['smoke_alarms']['devices']:
                    
                        alarm_data.append([
                            address,
                            business_name,
                            cells[0],  # Device
                            cells[1],  # Location
                            cells[5] if len(cells) > 4 else ""  # Remarks
                        ])

                        self.method_type['smoke_alarms']['devices'].add(identifier)

        self.update_workbook("Smoke Alarms", alarm_data)
        return None

    
    
    def emergency_lighting(self, file_path):
        """Extract data from Unit Emergency Lighting Test & Inspection Template"""
        
        tables = self.get_document_tables(file_path)
        extracted_data = {}
        lighting_data = []
        
        field_mappings = {
            "Business Name:": "Business_Name",
            "Address:": "Address", 
            "City:": "City",
            "Date of Service:": "Date"
        }

        for table in tables:
            for row in table.rows:
                cells = [cell.text.strip() for cell in row.cells]
                
                for i, cell_text in enumerate(cells):
                    if cell_text in field_mappings and i + 1 < len(cells):
                        extracted_data[field_mappings[cell_text]] = cells[i + 1]
                
                if (len(cells) >= 7 and 
                    cells[0] and 
                    cells[0] not in ["Unit Location", "Business Name:", "SPU", "BP", "RH", "EX", "COM"] and
                    not any(header in cells[0] for header in ["Monthly", "Annual", "UNIT TYPES", "Yes", "No"]) and
                    not cells[0].endswith(":")):
                    
                    business_name = extracted_data.get("Business_Name", '')
                    address = extracted_data.get("Address", '')
                    city = extracted_data.get("City", '')
                    identifier = f'{address} - {cells[0]} - {cells[1]}'
                    asset_type = ''
                    variant = ''

                    if at := self.eml_mapping.get(cells[1], ''): 
                        asset_type = at[0]
                        variant = at[1]

                    if identifier not in self.method_type["emergency_lighting"]["devices"]:

                        lighting_data.append([
                            f'{address} {city}',
                            business_name,
                            asset_type if asset_type else (cells[1] or ''), 
                            variant,
                            cells[0],  # Unit Location
                            cells[8] if len(cells) > 7 else "",  # Battery Size
                            cells[9] if len(cells) > 8 else "",  # Battery #
                            cells[10] if len(cells) > 9 else "",  # Battery Date
                            cells[11] if len(cells) > 10 else "",  # Voltage/Size
                            cells[12] if len(cells) > 11 else ""   # Comments
                        ])

                        self.method_type[method_name]["devices"].add(identifier)
        
        self.update_workbook("Emergency Lights", lighting_data)

    def emergency_lighting_extinguisher(self, file_path):
        """Extract data from Unit Emergency Lighting / Extinguisher Test & Inspection Template"""
        
        tables = self.get_document_tables(file_path)
        extracted_data = {}
        lighting_data = []
        extinguisher_data = []
        
        field_mappings = {
            "Business Name:": "Business_Name",
            "Address:": "Address", 
            "City:": "City",
            "Date of Service:": "Date"
        }
        
        if len(tables) > 0:
            for row in tables[0].rows:
                cells = [cell.text.strip() for cell in row.cells]
                for i, cell_text in enumerate(cells):
                    if cell_text in field_mappings and i + 1 < len(cells):
                        extracted_data[field_mappings[cell_text]] = cells[i + 1]
        
        # Extract emergency lighting data
        if len(tables) > 4:
            for row_idx, row in enumerate(tables[4].rows):
                if row_idx >= 2:  
                    cells = [cell.text.strip() for cell in row.cells]
                    if cells[0]: 
                        business_name = extracted_data.get("Business_Name", '')
                        address = extracted_data.get("Address", '')
                        city = extracted_data.get("City", '')
                        identifier = f'{address} - {cells[0]} - {cells[1]}'
                        asset_type = ''
                        variant = ''
                        if at := self.eml_mapping.get(cells[1], ''): 
                            asset_type = at[0]
                            variant = at[1]

                        if identifier not in self.method_type['emergency_lighting']['devices']:
                        
                            lighting_data.append([
                                f"{address} {city}",
                                business_name,
                                asset_type if asset_type else (cells[1] or ''), 
                                variant,
                                cells[0],  # Unit Location
                                cells[8] if len(cells) > 8 else "",  # Battery Size
                                cells[9] if len(cells) > 9 else "",  # Battery #
                                cells[10] if len(cells) > 10 else "", # Battery Date
                                cells[11] if len(cells) > 11 else "", # Voltage/Size
                                cells[12] if len(cells) > 12 else ""  # Comments
                            ])

                            self.method_type['emergency_lighting']['devices'].add(identifier)
        
        # Extract fire extinguisher data
        if len(tables) > 6:
            for row_idx, row in enumerate(tables[6].rows):
                if row_idx >= 1:
                    cells = [cell.text.strip() for cell in row.cells]
                    if cells[0]: 
                        business_name = extracted_data.get("Business_Name", '')
                        address = extracted_data.get("Address", '')
                        city = extracted_data.get("City", '')

                        identifier = cells[3] or f'None {random.randint(1000, 9999)}'
                        if identifier not in self.method_type["extinguishers"]["devices"]: 
                        
                            extinguisher_data.append([
                                f"{address} {city}",
                                business_name,
                                cells[0],  # Location
                                cells[1] if len(cells) > 1 else "",  # Size/Type
                                cells[2] if len(cells) > 2 else "",  # Brand
                                cells[3] if len(cells) > 3 else "",  # Serial #
                                cells[4] if len(cells) > 4 else "",  # Mfg. Date
                                cells[5] if len(cells) > 5 else "",  # Next Service Date
                                cells[6] if len(cells) > 6 else ""   # Comments
                            ])

                            self.method_type["extinguishers"]["devices"].add(identifier)
        
        if lighting_data:
            self.update_workbook("Emergency Lights", lighting_data)
        if extinguisher_data:
            self.update_workbook("Extinguishers", extinguisher_data)
    
    def special_suppression(self, file_path):
        """Extract data from Special Fire Suppression System Template"""
        
        tables = self.get_document_tables(file_path)
        extracted_data = {}
        suppression_data = []
        
        field_mappings = {
            "Business Name:": "Business_Name",
            "Address:": "Address", 
            "City:": "City",
        }
        
        system_types = [
            "FM-200", "Halon 1301", "Dry Chemical", "Carbon Dioxide", 
            "Argonite", "Novec 1230", "Foam", "Watermist", "Inergen"
        ]

        #asset type and variant mapping
        atv_mapping = {
            'FM-200' : ('NFPA 12A: Special Hazard: Gaseous (Cylinder)', 'HFC-227ea (FM-200, R-227)'), 
            'Halon 1301' : ('NFPA 12A: Special Hazard: Halon (Fire Extinguishing System)', 'Halon 1301'), 
            'Dry Chemical' : ('NFPA 17: Special Hazard: Dry Chemical (Fire Extinguishing System)', 'Dry Chemical Extinguishing System'), 
            'Carbon Dioxide' : ('NFPA 12: Standard on Carbon Dioxide Extinguishing Systems', 'Carbon Dioxide Extinguishing System'), 
            'Argonite' : ('NFPA 12A: Special Hazard: Gaseous (Cylinder)', 'IG-55 (Argonite)'), 
            'Novec 1230': ('NFPA 2001: Special Hazard: Clean-Agent (Fire Extinguishing System)', 'FK-5-1-12 (Novec 1230)'), 
            'Foam': ('NFPA 25: Special Hazard: Foam (Fire Extinguishing System)', 'AFFF (Aqueous Film Forming Foam)'), 
            'Watermist': ('NFPA 750: Standard on Water Mist Fire Protection', 'Watermist'), 
            'Inergen': ('NFPA 12A: Special Hazard: Gaseous (Cylinder)', 'Inergen')
        }
        
        for table in tables:
            for row in table.rows:
                cells = [cell.text.strip() for cell in row.cells]
                
                for i, cell_text in enumerate(cells):
                    if cell_text in field_mappings and i + 1 < len(cells):
                        extracted_data[field_mappings[cell_text]] = cells[i + 1]
                

                for cell in cells:
                    for system_type in system_types:
                        if system_type in cell:

                            row_data = cells
                            make_value = ""
                            model_value = ""
                            
                            for j, cell_content in enumerate(row_data):
                                if cell_content == "Make:" and j + 1 < len(row_data):
                                    make_value = row_data[j + 1]
                                elif cell_content == "Model:" and j + 1 < len(row_data):
                                    model_value = row_data[j + 1]
                            
                            if make_value or model_value:
                                # Get asset type and variant from mapping
                                asset_type, variant = atv_mapping.get(system_type, ('Unknown', 'Unknown'))

                                business_name = extracted_data.get("Business_Name", '')
                                address = extracted_data.get("Address", '')
                                city = extracted_data.get("City", '')
                                identifier = f"{business_name} - {make_value + ' - ' if make_value else ''}{model_value if model_value else ''}"
                                
                                if identifier not in self.method_type['special_suppression']['devices']: 
                                
                                    suppression_data.append([
                                        f"{address} {city}",
                                        business_name,
                                        asset_type,
                                        variant,
                                        make_value,
                                        model_value
                                    ])

                                    self.method_type['special_suppression']['devices'].add(identifier)
        
        if suppression_data:
            self.update_workbook("Special Suppression", suppression_data)
    

    def alarm_system_devices(self, file_path, text):
        
        tables = self.get_document_tables(file_path)
        if not tables:
            return
        
        target_headers = [
            ["Device", "Location", "A", "B", "C", "D", "E", "F", "Remarks"],
            ["Device", "Annunciation Label / Device Location", "A", "B", "C", "D", "E", "F", "G", "Remarks"]
        ]

        target_table = None
        device_data = []
                        
        # Process the first table (which contains the header information)
        if tables and len(tables) > 0:
            if 'nbc' in text.lower(): 
                uid = self.alarm_system_new(tables[0])
            else: 
                uid = self.alarm_system_old_fav(tables[0])

        if uid: # system was successfully created  

            # find device data table
            for table in tables:
                if not table.rows:
                    continue
                    
                if not target_table and len(table.rows) > 0:
                    header_row = [cell.text.strip() for cell in table.rows[0].cells]
                    for target_header in target_headers:
                        if self._headers_match(header_row, target_header):
                            target_table = table
                            break
            
            if target_table:
                for row_idx, row in enumerate(target_table.rows):
                    if row_idx == 0: 
                        continue
                        
                    cells = [cell.text.strip() for cell in row.cells]
                    
                    if not cells or not cells[0]:
                        continue
                    
                    at = self.alarm_mapping.get(cells[0], ['', ''])[0]
                    var = self.alarm_mapping.get(cells[0], ['', ''])[1]
                    
                    device_entry = {
                        'system' : uid,
                        'asset_type': at,
                        'varient': var,
                        "zcn": cells[6] if len(cells) > 6 else "", 
                        
                    }
                    
                    device_data.append([
                        device_entry["system"],
                        device_entry["asset_type"],
                        device_entry["varient"], 
                        device_entry["zcn"]
                    ])
            

            if device_data:

                self.update_workbook("Alarm System Devices", device_data)

    def alarm_system_new(self, table):

            system_uid = shortuuid.ShortUUID().random(length=10)
            data = {}
            
            if len(table.rows) >= 11:
    
                if len(table.rows) > 9:
                    row_9_cells = [cell.text.strip() for cell in table.rows[9].cells]
                    print(row_9_cells[14])
                    row_10_cells = [cell.text.strip() for cell in table.rows[10].cells]
                    
                    if len(row_9_cells) > 5:
                        data["manufacturer"] = row_9_cells[6]
                    
                    if len(row_9_cells) > 9:
                            data["model_number"] = row_9_cells[9]
                    
                    if len(row_9_cells) > 13:
                            data["ulc_serial_number"] = row_9_cells[14]
                
                if len(table.rows) > 10:
                    row_11_cells = [cell.text.strip() for cell in table.rows[10].cells]
                    
                    if len(row_11_cells) > 2 and "Business/Building Name:" in row_11_cells[1]:
                        if row_11_cells[2]:
                           data["business_name"] = row_11_cells[2]

                    if len(row_11_cells) > 4 and "Address:" in row_11_cells[3]:
                        if len(row_11_cells) > 4 and row_11_cells[4]:
                            data["address"] = row_11_cells[4]
                    
                    if len(row_11_cells) > 14 and "City:" in row_11_cells[11]:
                        if len(row_11_cells) > 14 and row_11_cells[14]:
                            data["city"] = row_11_cells[14]
                

                if len(table.rows) > 11:
                    row_12_cells = [cell.text.strip() for cell in table.rows[11].cells]
                    
                    for i, cell_text in enumerate(row_12_cells):
                        if "Fire Signal Receiving Centre" in cell_text and i + 1 < len(row_12_cells):
                            if row_12_cells[i + 1]:
                                data["fire_signal_receiving_centre"] = row_12_cells[i + 1]
                
                identifier = data.get('address', '')
                if identifier not in self.method_type['alarm_system']['devices']: 

                    self.update_workbook("Alarm Systems", [[
                        f"{data.get('address', '')} {data.get('city', '')}",
                        system_uid,
                        data.get('business_name', ''),
                        data.get('manufacturer'), 
                        data.get('model_number'), 
                        data.get('fire_signal_receiving_centre', ''),
                        data.get('ulc_serial_number', '')
                    ]])

                    self.method_type['alarm_system']['devices'].add(identifier)

                return system_uid

    def alarm_system_old_fav(self, table):

        system_uid = shortuuid.ShortUUID().random(length=10)
        data = {}
        
        field_mappings = {
            "Business Name:": "Business_Name",
            "Address:": "Address", 
            "City:": "City",
            "Fire Signal Receiving Centre:": "fire_signal_receiving_centre", 
            "Manufacturer:" : "manufacturer",
            "System Manufacturer:" : "manufacturer",
            "Model #:" : "model_number" 
        }
        
        for row in table.rows:
            cells = [cell.text.strip() for cell in row.cells]
            
            for i, cell_text in enumerate(cells):
                if cell_text in field_mappings and i + 1 < len(cells):
                    data[field_mappings[cell_text]] = cells[i + 1]
        
        identifier = data.get('Address', '')
        if identifier not in self.method_type['alarm_system']['devices']:

            self.update_workbook("Alarm Systems", [[
                        f"{data.get('Address', '')} {data.get('City', '')}",
                        system_uid,
                        data.get('Business_Name', ''),
                        data.get('manufacturer'), 
                        data.get('model_number'), 
                        data.get('fire_signal_receiving_centre', ''),
                    ]])

            self.method_type['alarm_system']['devices'].add(identifier)
            
        return system_uid

    def sprinkler_systems(self, file_path): 
        
        tables = self.get_document_tables(file_path)
        data = {}

        field_mappings = {
            "Business Name:": "Business_Name",
            "Address:": "Address", 
            "City:": "City", 
            "REPORT OF INSPECTION/TEST FOR WET SYSTEM:" : "Wet_System",
            "REPORT OF INSPECTION/TEST FOR DRY SYSTEM:" : "Dry_System"
        }

        for table in tables: 
            for row in table.rows:
                cells = [cell.text.strip() for cell in row.cells]
                
                for i, cell_text in enumerate(cells):
                    if cell_text in field_mappings and i + 1 < len(cells):
                        data[field_mappings[cell_text]] = cells[i + 1]

        if data.get('Wet_System', ''): 
            wet_uid = shortuuid.ShortUUID().random(length=10)
            self.wet_system_devices(tables, wet_uid)

            identifier = f'{data.get('Address', '')} - {data.get('REPORT OF INSPECTION/TEST FOR WET SYSTEM:', '')}'
            if identifier not in self.method_type['wet_systems']['devices']: 
                self.update_workbook("Sprinkler Systems", [[
                        f"{data.get('Address', '')} {data.get('City', '')}",
                        data.get('Business_Name', ''),
                        system_uid,
                        data.get('REPORT OF INSPECTION/TEST FOR WET SYSTEM:', ''),
                        "NFPA 25: Fire Sprinkler System: Wet Pipe (General)", 
                        "General System"
                    ]])

        if data.get('Dry_System', ''): 
            dry_uid = shortuuid.ShortUUID().random(length=10)
            self.dry_system_devices(tables, dry_uid)

            identifier = f'{data.get('Address', '')} - {data.get('REPORT OF INSPECTION/TEST FOR DRY SYSTEM:', '')}'
            if identifier not in self.method_type['dry_systems']['devices']: 
                self.update_workbook("Sprinkler Systems", [[
                        f"{data.get('Address', '')} {data.get('City', '')}",
                        data.get('Business_Name', ''),
                        system_uid,
                        data.get('REPORT OF INSPECTION/TEST FOR DRY SYSTEM:', ''),
                        "NFPA 25: Fire Sprinkler System: Dry Pipe Fire (General)", 
                        "General System"
                    ]])
        return None

    def wet_system_devices(self, tables, uid): 

        # Initialize all table variables
        table_vars = {
            "main_drain_table": None,
            "hose_valves_table": None,
            "test_connection_table": None,
            "valves_table": None,
            "drain_valves_table": None
        }

        target_headers = {
            "main_drain_table": "Initial Static \n(psi)",
            "hose_valves_table": "Hose Valve",  
            "test_connection_table": "Smooth \n Orifice",  
            "valves_table": "Valve Type",
            "drain_valves_table": "Drain Valve",  
        }

        for table in tables:
            if not table.rows:
                continue
            
            # Check each target header to see if it exists anywhere in this table
            for table_name, header_text in target_headers.items():
                row_idx, col_idx = self.find_header_row(table, header_text)
                
                if row_idx is not None:
                    table_vars[table_name] = table
                    print(f"Found {table_name} at row {row_idx}, column {col_idx}")
                    break  # Stop checking other headers once we find a match

        # Extract the assigned tables
        main_drain_table = table_vars["main_drain_table"]
        hose_valves_table = table_vars["hose_valves_table"]
        test_connection_table = table_vars["test_connection_table"]
        valves_table = table_vars["valves_table"]
        drain_valves_table = table_vars["drain_valves_table"]

        # Process each table if it was found
        for table_name, table in table_vars.items():
            if table:
                print(f"Processing {table_name}...")








    

    

    

        

        
            
                
