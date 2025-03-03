import os
from pyrevit import script, forms

try:
    folder_name = "c:\\Temp"
    filepath = os.path.join(folder_name, 'Ribbon_Sleeve.txt')

    if not os.path.exists(folder_name):
        os.makedirs(folder_name)
    if not os.path.exists(filepath):
        with open(filepath, 'w') as f:
            f.write('1')

    with open(filepath, 'r') as f:
        SleeveLength = f.read()
        SleeveLength = round(float(SleeveLength) * 12, 3)

    SleeveLength = forms.ask_for_string(default= str(SleeveLength), prompt='Default Sleeve Length (Inches)', title='Sleeve Configuration')

    # Convert dialog input into variable
    SleeveLength = str(float(SleeveLength) / 12)

    with open(filepath, 'w') as f:
        f.write (SleeveLength)
except:
    pass