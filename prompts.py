class Prompts:

    hydrants = """
    EXTRACT EXACTLY 10 VALUES FROM THIS DOCUMENT IN THIS EXACT ORDER:

    1. Private or Public (check which checkbox is selected - if neither, use 'blank')
    2. Hydrant # (the hydrant number - if missing, use 'blank')
    3. Address (combine street address + city into single BC address - if missing, use 'blank')
    4. Date of service (convert to unix timestamp - if missing, use 'blank')
    5. Make (if missing, use 'blank')
    6. Model (if missing, use 'blank')
    7. Color (if missing, use 'blank')
    8. Hydrant Shut-Off Location (if missing, use 'blank')
    9. Hydrant Location (if missing, use 'blank')
    10. Business Name (if missing, use 'blank')

    CRITICAL FORMATTING RULES:
    - Output ONLY the comma-separated values with NO spaces after commas
    - Remove ALL commas from extracted values before outputting
    - Replace any missing/empty fields with exactly 'blank'
    - Must return exactly 10 values separated by commas
    - No explanatory text, headers, or additional content

    Example format: Public,H-001,123 Main St Vancouver BC,1672531200,Acme,Model-X,Red,Street,Corner,ABC Company
    """

    fixed_extinguishing_systems = """
    EXTRACT EXACTLY 9 VALUES FROM THIS DOCUMENT IN THIS EXACT ORDER:

    1. Wet Chemical or Dry Chemical (check which checkbox is selected - if both checked, use 'Wet Chemical' - if neither, use 'blank')
    2. Last Recharge Date (if missing, use 'blank')
    3. Address (combine street address + city into single BC address - if missing, use 'blank')
    4. Business Name (if missing, use 'blank')
    5. Location of System Cylinders (if missing, use 'blank')
    6. Electric or Gas (check which checkbox is selected - if neither, use 'blank')
    7. Manufacturer (if missing, use 'blank')
    8. Model Number (if missing, use 'blank')
    9. Serial Number (if missing, use 'blank')

    CRITICAL FORMATTING RULES:
    - Output ONLY the comma-separated values with NO spaces after commas
    - Remove ALL commas from extracted values before outputting
    - Replace any missing/empty fields with exactly 'blank'
    - Must return exactly 9 values separated by commas
    - No explanatory text, headers, or additional content

    Example format: Wet Chemical,2024-01-15,123 Main St Vancouver BC,ABC Restaurant,Kitchen,Gas,Ansul,R-102,SN123456
    """

    backflow_page_1 = """
    EXTRACT EXACTLY 7 VALUES FROM THIS DOCUMENT IN THIS EXACT ORDER:

    1. Name of Premise (if missing, use 'blank')
    2. Service Address (if missing, use 'blank')
    3. Manufacturer (if missing, use 'blank')
    4. Model # (if missing, use 'blank')
    5. Serial # (if missing, use 'blank')
    6. Type (if missing, use 'blank')
    7. Size (if missing, use 'blank')

    CRITICAL FORMATTING RULES:
    - Output ONLY the comma-separated values with NO spaces after commas
    - Remove ALL commas from extracted values before outputting
    - Replace any missing/empty fields with exactly 'blank'
    - Must return exactly 7 values separated by commas
    - No explanatory text, headers, or additional content

    Example format: ABC Building,123 Main St,Watts,909,12345,RPZ,2 inch
    """

    backflow_page_2 = """
    EXTRACT THE LOCATION OF THE BACKFLOW PREVENTER FROM THIS DOCUMENT.

    CRITICAL FORMATTING RULES:
    - Output ONLY the location text with no additional words
    - If the location field is blank or missing, output exactly 'unknown'
    - Remove any commas from the location before outputting
    - No explanatory text or additional content

    Example outputs: 
    - Basement utility room
    - unknown
    - Front yard near meter
    """