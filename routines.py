import csv 

routines = {}
mapping = {
    'EXT' : 'NFPA 10: Fire Extinguishers', 
    'EML' : 'NFPA 101: Emergency Escape Light & Exit Signs', 
    'SS' : 'Special Suppression', 
    'FA' : 'NFPA 72: Fire Detection: (Fire Alarm Control Panel (FACP)', 
    'SMK (smoke alarm only)': 'NFPA 72: Smoke and Heat Alarms', 
    '5yr/3yr' : ''
    }

with open('routines.csv', 'r', encoding='utf-8') as file: 
    reader = csv.reader(file)
    rows = [row for row in reader]

    for i, row in enumerate(rows): 

        if i == 0: 
            continue 

        client = row[0]
        address = f'{row[1]}, {row[2]}'
