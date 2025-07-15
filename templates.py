import random
import shortuuid
from docx.oxml.ns import qn
from prompts import Prompts

class AssetTemplateMethods:
    """Methods for extracting data from different types of assets"""

    def __init__(self):

    
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
            'COM' : ('NFPA 101: Emergency Escape Lighting & Exit Signs', 'Combo Unit'), 
            'INV' : ('NFPA 101: Emergency Escape Lighting & Exit Signs', 'Inverter'), 
            'BP' : ('NFPA 101: Emergency Escape Lighting & Exit Signs', 'Battery Pack'), 
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
        "PZ": ("NFPA 72: Sound and Intercom For Emergency Purposes: Audible/Visual", "In-Suite Buzzer"), 
        "AV": ("NFPA 72: Sound and Intercom For Emergency Purposes: Audible/Visual", "A/V Device"),
        "SA": ("NFPA 72: Smoke and Heat Alarm", "Smoke Alarm"),
        "O": ("NFPA 72: Smoke and Heat Alarm", "Override"), 
        
    }
        
    def fixed_extinguishing_systems(self, file_path):
        """Extract data from Fixed Extinguishing Systems Template"""

        response = self.api_call(file_path, page=0, prompt=Prompts.fixed_extinguishing_systems)
        data = response.split(',')

        chem_type = data[0]
        last_recharge = data[1]
        address = data[2]
        business_name = data[3]
        location_of_cylinders = data[4]
        power = data[5]
        manufacturer = data[6]
        model_num = data[7]
        serial = data[8]

        if "wet" in chem_type.lower(): 
            ct = 1
        else: 
            ct = 2
        
        chemical_types = {
            1: {
                "asset_type": "NFPA 17A: Special Hazard: Wet Chemical System (Cylinder)", 
                "variant": "Wet Chemical"
            },
            2: {
                "asset_type": "NFPA 17: Special Hazard: Dry Chemical (Fire Extinguishing System)", 
                "variant": "Dry Chemical Extinguishing System"
            }
        }
       
        at = chemical_types[ct]['asset_type']
        var = chemical_types[ct]['variant']
        
        identifier = f'{address} - {model_num} - {serial}'
        if identifier not in self.method_type['fixed_extinguishing_systems']['devices']:
        
            self.update_workbook("Fixed Extinguishing Systems", [[
                address,
                business_name,
                at, 
                var,
                last_recharge, 
                location_of_cylinders, 
                manufacturer, 
                model_num, 
                serial, 
                power,
            ]])
        
            self.method_type['fixed_extinguishing_systems']['devices'].add(identifier)

    def fire_hydrants(self, file_path):
        """Extract data from Fire Hydrant Inspection & Testing Template"""

        response = self.api_call(file_path, page=0, prompt=Prompts.hydrants)
        data = response.split(',')

        hydrant_type = data[0]
        hydrant_number = data[1]
        address = data[2]
        make = data[4]
        model = data[5]
        color = data[6]
        shut_off_location = data[7]
        location = data[8]
        business_name = data[9]
        
        if f'{address} - {hydrant_number}' not in self.method_type['fire_hydrants']['devices']:
            self.update_workbook("Fire Hydrants", [[
                address,
                business_name,
                hydrant_number,
                location,
                make,
                model,
                color,
                shut_off_location,
                hydrant_type
            ]])
            self.method_type['fire_hydrants']['devices'].add(f'{address} - {hydrant_number}')

        return None
            
    def backflows(self, file_path):
        """Extract data from Backflow Prevention Assembly Test Report Template"""

        response = self.api_call(file_path, page=0, prompt=Prompts.backflow_page_1)
        data = response.split(',')

        business_name = data[0]
        address = data[1]
        manufacturer = data[2]
        model = data[3]
        serial = data[4]
        type = data[5].upper()
        size = data[6]
        location = self.api_call(file_path, page=1, prompt=Prompts.backflow_page_2)

    
        var_mapping = {
            'RP' : 'Reduced Pressure Zone', 
            'DCVA' : 'Double Check Valve Assembly', 
            'RPBA': 'Reduced Pressure Backflow Assembly'
        }
        
        if f'{business_name} - {serial}' not in self.method_type['backflows']['devices']:
        
            self.update_workbook("Backflows", [[
                business_name,
                address,
                "NFPA 25: Automatic Back-Flow Prevention", 
                var_mapping.get(type, ''),
                manufacturer,
                model,
                serial,
                size,
                location
            ]])

            self.method_type['backflows']['devices'].add(f'{business_name} - {serial}')
        
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

                    if at := self.eml_mapping.get(cells[1].upper(), ''): 
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

                        self.method_type["emergency_lighting"]["devices"].add(identifier)
        
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
                    
                    if len(set(cells)) > 2:

                        if "X" in cells[0].upper(): 
                            cells[0] = cells[0][:1]
                    
                        at = self.alarm_mapping.get(cells[0].upper(), ['', ''])[0]
                        var = self.alarm_mapping.get(cells[0].upper(), ['', ''])[1]

                        if not at: 
                            at = cells[0]
                        
                        
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
                    
                    else: 
                        continue
                

                if device_data:

                    self.update_workbook("Alarm System Devices", device_data)

    def alarm_system_new(self, table):

        system_uid = shortuuid.ShortUUID().random(length=10)
        data = {}
        
        if len(table.rows) >= 11:

            if len(table.rows) > 9:
                row_9_cells = [cell.text.strip() for cell in table.rows[9].cells]                
                row_10_cells = [cell.text.strip() for cell in table.rows[10].cells]
                
                # Add length checks before accessing cells
                if len(row_9_cells) > 6:
                    data["manufacturer"] = row_9_cells[6]
                
                if len(row_9_cells) > 9:
                    data["model_number"] = row_9_cells[9]
                
                if len(row_9_cells) > 14:
                    data["ulc_serial_number"] = row_9_cells[14]
            
            if len(table.rows) > 10:
                row_11_cells = [cell.text.strip() for cell in table.rows[10].cells]
                
                if len(row_11_cells) > 2 and "Business/Building Name:" in row_11_cells[1]:
                    if len(row_11_cells) > 2 and row_11_cells[2]:
                        data["business_name"] = row_11_cells[2]

                if len(row_11_cells) > 4 and "Address:" in row_11_cells[3]:
                    if row_11_cells[4]:
                        data["address"] = row_11_cells[4]
                
                if len(row_11_cells) > 14 and "City:" in row_11_cells[11]:
                    if row_11_cells[14]:
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
            
            identifier = f'{data.get('Address', '')} - {data.get('REPORT OF INSPECTION/TEST FOR WET SYSTEM:', '')}'
            if identifier not in self.method_type['wet_systems']['devices']: 
                self.update_workbook("Sprinkler Systems", [[
                        f"{data.get('Address', '')} {data.get('City', '')}",
                        data.get('Business_Name', ''),
                        data.get('Wet_System', ''),
                        
                    ]])

        if data.get('Dry_System', ''): 
            
            identifier = f'{data.get('Address', '')} - {data.get('REPORT OF INSPECTION/TEST FOR DRY SYSTEM:', '')}'
            if identifier not in self.method_type['dry_systems']['devices']: 
                self.update_workbook("Sprinkler Systems", [[
                        f"{data.get('Address', '')} {data.get('City', '')}",
                        data.get('Business_Name', ''),
                        data.get('Dry_System', ''),
            
                    ]])
       
        return None


