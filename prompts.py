class Prompts: 

    hydrants = """
    Please observe this document and extract the following from it:

    - Private / Public (depending on which checkbox is selected)
    - Hydrant #
    - Address (combine address and city for me to make a proper british columbia address)
    - Date of service (converted to unix timestamp)
    - Make
    - Model
    - Color
    - Hydrant Shut-Off Location
    - Hydrant Location
    - Business Name

    please return your response as nothing but a comma seperated list with no spaces between. It is very imporant that you return nothing but that.

    If there are any commas in any of the values that you extract, remove them from the string before giving me your response. 

    If there are any blanks in the fields i requested, please add them as 'blank' to your response. 
    
    """

    fixed_extinguishing_systems = """

    Please observe this document and extract the following from it:

    - Wet Chemical / Dry Chemical (depending on which checkbox is selected)
    - Last Recharge Date
    - Address (combine address and city for me to make a proper british columbia address)
    - Business Name
    - Location of System Cylinders
    - Electric / Gas (depending on which checkbox is selected)
    - Manufacturer
    - Model Number
    - Serial Number

    please return your response as nothing but a comma seperated list with no spaces between each item (the items themselves can have spaces). It is very imporant that you return nothing but that and in that exact order.

    If there are any commas in any of the values that you extract, remove them from the string before giving me your response. 

    If there are any blanks in the fields I requested, please add them as 'blank' to your response. 

    If for any reason both Wet and Dry chemical are selected, default to Wet Chemical

    """

    backflow_page_1 = """

    Please observe this document and extract the following from it:

    -Name of Premise
    -Service Address
    -Manufacturer 
    -Model #
    -Serial #
    -Type 
    -Size

    please return your response as nothing but a comma seperated list with no spaces between each item (the items themselves can have spaces). It is very imporant that you return nothing but that and in that exact order.

    If there are any commas in any of the values that you extract, remove them from the string before giving me your response. 

    If there are any blanks in the fields I requested, please add them as 'blank' to your response.

    """

    backflow_page_2 = """

    please look at this document and tell me where the backflow preventer is located. It is very important that your response contains only the location of the backflow preventer
    and nothing else. If that field happens to be blank, only respond with 'unknown'


    """

    backflow_page_1 = """

    Please observe this document and extract the following from it:

    -Name of Premise
    -Service Address
    -Manufacturer 
    -Model #
    -Serial #
    -Type 
    -Size

    please return your response as nothing but a comma seperated list with no spaces between each item (the items themselves can have spaces). It is very imporant that you return nothing but that and in that exact order.

    If there are any commas in any of the values that you extract, remove them from the string before giving me your response. 

    If there are any blanks in the fields I requested, please add them as 'blank' to your response.

    """


