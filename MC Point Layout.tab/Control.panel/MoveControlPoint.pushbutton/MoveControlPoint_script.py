# -*- coding: utf-8 -*-
from pyrevit import revit, DB, UI, forms
from Autodesk.Revit.DB import Transaction, XYZ, ElementId
from rpw.ui.forms import FlexForm, Label, TextBox, Button
import os

# Get the current Revit document and UI document
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

# Define the file path for saving distances
folder_name = "c:\\Temp"
filepath = os.path.join(folder_name, 'Ribbon_MoveGeneric.txt')

# Custom selection filter for generic model elements
class GenericModelFilter(UI.Selection.ISelectionFilter):
    def AllowElement(self, element):
        # Allow only elements in the Generic Models category
        return element.Category and element.Category.Id.IntegerValue == int(DB.BuiltInCategory.OST_GenericModel)

    def AllowReference(self, reference, point):
        return False  # Not allowing references (e.g., edges, faces)

# Function to prompt user to window select generic model elements
def pick_elements():
    try:
        # Window selection with filter
        filter = GenericModelFilter()
        selections = uidoc.Selection.PickObjects(
            UI.Selection.ObjectType.Element,
            filter,
            "Window select generic model elements to move (drag from left to right)."
        )
        if selections:
            selected_elements = [doc.GetElement(sel.ElementId) for sel in selections]
            return selected_elements
        else:
            forms.alert("No generic model elements selected. Script will exit.", exitscript=True)

    except:
        forms.alert("Selection cancelled or failed. Script will exit.", exitscript=True)

# Function to load saved distances from file
def load_saved_distances():
    # Create folder and file if they don't exist
    if not os.path.exists(folder_name):
        os.makedirs(folder_name)
    if not os.path.exists(filepath):
        with open(filepath, 'w') as the_file:
            line1 = "0\n"  # Default X distance
            line2 = "0\n"  # Default Y distance
            the_file.writelines([line1, line2])

    # Read text file for stored values
    with open(filepath, 'r') as file:
        lines = [line.rstrip() for line in file.readlines()]

    # Ensure file has at least 2 lines, reset to defaults if not
    if len(lines) < 2:
        with open(filepath, 'w') as the_file:
            line1 = "0\n"
            line2 = "0\n"
            the_file.writelines([line1, line2])
        return "0", "0"

    return lines[0], lines[1]

# Function to save distances to file
def save_distances(x_dist, y_dist):
    with open(filepath, 'w') as file:
        file.write("{}\n{}\n".format(x_dist, y_dist))

# Main script execution
def move_elements():
    # Step 1: Prompt user to window select generic model elements
    elements = pick_elements()
    if not elements:
        return

    # Step 2: Load saved distances
    saved_x_dist, saved_y_dist = load_saved_distances()

    # Step 3: Create and show the FlexForm dialog for X and Y distances with saved values
    components = [
        Label("X Distance (ft) (positive = right, negative = left):"),
        TextBox("x_dist", default=saved_x_dist, placeholder="Enter X distance (positive = right, negative = left)"),
        Label("Y Distance (ft) (positive = up, negative = down):"),
        TextBox("y_dist", default=saved_y_dist, placeholder="Enter Y distance (positive = up, negative = down)"),
        Button("OK")
    ]
    form = FlexForm("Move Generic Model Elements", components)
    result = form.show()

    # Step 4: Exit quietly if dialog is closed (result is False or None)
    if not result:  # Dialog closed without clicking "OK"
        return

    # Step 5: Retrieve and validate the input values
    values = form.values
    try:
        x_dist = float(values["x_dist"])
        y_dist = float(values["y_dist"])
    except ValueError:
        forms.alert("Please enter valid numeric values for distances. Script will exit.", exitscript=True)
        return

    # Step 6: Save the new distances
    save_distances(x_dist, y_dist)

    # Step 7: Move each selected element individually by the specified distances
    with revit.Transaction("Move Generic Model Elements"):
        # Create a translation vector (Revit uses feet internally)
        translation = XYZ(x_dist, y_dist, 0)
        for element in elements:
            # Move each element individually from its current location
            DB.ElementTransformUtils.MoveElement(doc, element.Id, translation)

    # Optional: Uncomment if you want a confirmation message
    # forms.alert("Elements moved successfully!\nX: {} ft, Y: {} ft".format(x_dist, y_dist))

# Execute the script
if __name__ == "__main__":
    move_elements()