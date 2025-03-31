# -*- coding: utf-8 -*-
from pyrevit import revit, DB, forms, script
import clr
# Add Windows Forms references
clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')
from System.Windows.Forms import (Form, Label, ListBox, Button, FormBorderStyle, 
                                  FormStartPosition, SelectionMode, MessageBox, 
                                  MessageBoxButtons, DialogResult)
from System import Array
from System.Drawing import Point, Size
import os

# Add reference to Excel COM interop
clr.AddReference("Microsoft.Office.Interop.Excel")
from Microsoft.Office.Interop.Excel import ApplicationClass

# Global variable to store the last exported directory
last_export_dir = os.path.expanduser("~") + "\\Desktop"  # Default to Desktop

# Function to export selected title block parameters to Excel
def export_titleblocks_to_excel(title_blocks, filepath, selected_params):
    global last_export_dir
    excel = ApplicationClass()
    excel.Visible = False
    wb = excel.Workbooks.Add()
    ws = wb.Worksheets[1]
    ws.Name = "Title Blocks"

    # Headers
    headers = ["ElementId", "Sheet Number"] + selected_params

    for col, header in enumerate(headers, 1):
        ws.Cells[1, col].Value2 = header

    row = 2
    for tb in title_blocks:
        ws.Cells[row, 1].Value2 = str(tb.Id.IntegerValue)
        sheet_number_param = tb.LookupParameter("Sheet Number")
        ws.Cells[row, 2].Value2 = sheet_number_param.AsString() if sheet_number_param else "N/A"

        col = 3
        for param_name in selected_params:
            param = tb.LookupParameter(param_name)
            value = ""
            if param:
                if param.StorageType == DB.StorageType.String:
                    value = param.AsString() if param.AsString() is not None else ""
                elif param.StorageType == DB.StorageType.Double:
                    value = param.AsDouble()
                elif param.StorageType == DB.StorageType.Integer:
                    value = param.AsInteger()
            ws.Cells[row, col].Value2 = value
            col += 1
        row += 1

    # Auto-fit all columns based on content
    total_columns = len(headers)
    ws.Range[ws.Cells[1, 1], ws.Cells[row - 1, total_columns]].Columns.AutoFit()

    wb.SaveAs(filepath)
    wb.Close()
    excel.Quit()
    last_export_dir = os.path.dirname(filepath)

# Function to import changes from Excel for title blocks
def import_changes_from_excel(filepath):
    if not os.path.exists(filepath):
        forms.alert("Excel file not found at:\n{}".format(filepath))
        return
    
    excel = ApplicationClass()
    excel.Visible = False
    wb = excel.Workbooks.Open(filepath)
    ws = wb.Worksheets[1]
    
    # Get headers
    headers = []
    col = 1
    while ws.Cells[1, col].Value2:
        headers.append(ws.Cells[1, col].Value2)
        col += 1
    header_map = {h: i+1 for i, h in enumerate(headers)}  # 1-based indexing for Excel

    # Import data
    row_count = ws.UsedRange.Rows.Count
    if row_count <= 1:
        wb.Close()
        excel.Quit()
        forms.alert("No data found in Excel file.")
        return
    
    updated_count = 0
    with revit.Transaction("Update Title Blocks from Excel"):
        for row in range(2, row_count + 1):
            elem_id = ws.Cells[row, header_map["ElementId"]].Value2
            if elem_id:
                try:
                    elem_id = int(float(elem_id))  # Excel might return as float
                    element = revit.doc.GetElement(DB.ElementId(elem_id))
                    if element:
                        for header, col_idx in header_map.items():
                            if header != "ElementId":  # Skip ElementId column
                                param = element.LookupParameter(header)
                                if param and not param.IsReadOnly:
                                    value = ws.Cells[row, col_idx].Value2
                                    if value is not None:
                                        try:
                                            if param.StorageType == DB.StorageType.String:
                                                param.Set(str(value))
                                            elif param.StorageType == DB.StorageType.Double:
                                                param.Set(float(value) if isinstance(value, (int, float)) else 0.0)
                                            elif param.StorageType == DB.StorageType.Integer:
                                                param.Set(int(float(value)) if isinstance(value, (int, float)) else 0)
                                            updated_count += 1
                                        except Exception as e:
                                            print("Failed to set param {} for ElementId {}: {}".format(header, elem_id, e))
                except Exception as e:
                    print("Failed to process ElementId {}: {}".format(elem_id, e))
                    continue
    
    # Cleanup
    wb.Close()
    excel.Quit()
    forms.alert("Import complete")

# Dialog for selecting title blocks
class TitleBlockSelectionDialog(Form):
    def __init__(self, title_blocks):
        self.Text = "Select Title Blocks"
        self.Width = 400
        self.Height = 400
        self.FormBorderStyle = FormBorderStyle.FixedDialog
        self.StartPosition = FormStartPosition.CenterScreen

        # Label
        self.label = Label()
        self.label.Text = "Select Title Blocks to Export:"
        self.label.Location = Point(20, 10)
        self.label.Size = Size(350, 20)

        # ListBox for title blocks
        self.tb_listbox = ListBox()
        self.tb_listbox.Location = Point(20, 40)
        self.tb_listbox.Size = Size(350, 260)
        self.tb_listbox.SelectionMode = SelectionMode.MultiExtended
        self.title_blocks = title_blocks
        display_items = []
        for tb in title_blocks:
            family_name = tb.Symbol.Family.Name if tb.Symbol and tb.Symbol.Family else "N/A"
            display_items.append(family_name)
        self.tb_listbox.Items.AddRange(Array[object](display_items))

        # OK Button
        self.ok_button = Button()
        self.ok_button.Text = "OK"
        self.ok_button.Location = Point(150, 320)
        self.ok_button.Size = Size(100, 30)
        self.ok_button.Click += self.ok_clicked

        # Add Controls
        self.Controls.Add(self.label)
        self.Controls.Add(self.tb_listbox)
        self.Controls.Add(self.ok_button)

        self.selected_title_blocks = []
        self.selected_family_name = None

    def ok_clicked(self, sender, args):
        selected_indices = list(self.tb_listbox.SelectedIndices)
        self.selected_title_blocks = [self.title_blocks[i] for i in selected_indices]
        if self.selected_title_blocks:
            first_tb = self.selected_title_blocks[0]
            self.selected_family_name = first_tb.Symbol.Family.Name if first_tb.Symbol and first_tb.Symbol.Family else "TitleBlockExport"
        self.Close()

# WinForms dialog for selecting and ordering parameters
class ParameterSelectionDialog(Form):
    def __init__(self, param_names):
        self.Text = "Select and Order Parameters"
        self.Width = 500
        self.Height = 400
        self.FormBorderStyle = FormBorderStyle.FixedDialog
        self.StartPosition = FormStartPosition.CenterScreen

        # Labels
        self.label_available = Label()
        self.label_available.Text = "Available Parameters"
        self.label_available.Location = Point(20, 10)
        self.label_available.Size = Size(180, 20)

        self.label_selected = Label()
        self.label_selected.Text = "Selected Parameters (Order Matters)"
        self.label_selected.Location = Point(270, 10)
        self.label_selected.Size = Size(200, 20)

        # ListBox - Available parameters
        self.available_listbox = ListBox()
        self.available_listbox.Location = Point(20, 40)
        self.available_listbox.Size = Size(180, 260)
        self.available_listbox.SelectionMode = SelectionMode.MultiExtended
        self.available_listbox.Items.AddRange(Array[object](param_names))

        # ListBox - Selected parameters
        self.selected_listbox = ListBox()
        self.selected_listbox.Location = Point(270, 40)
        self.selected_listbox.Size = Size(180, 260)
        self.selected_listbox.SelectionMode = SelectionMode.MultiExtended

        # Buttons to move items
        self.add_button = Button()
        self.add_button.Text = ">>"
        self.add_button.Location = Point(210, 80)
        self.add_button.Size = Size(50, 30)
        self.add_button.Click += self.add_selected

        self.remove_button = Button()
        self.remove_button.Text = "<<"
        self.remove_button.Location = Point(210, 120)
        self.remove_button.Size = Size(50, 30)
        self.remove_button.Click += self.remove_selected

        # Buttons to reorder
        self.up_button = Button()
        self.up_button.Text = "Up"
        self.up_button.Location = Point(210, 180)
        self.up_button.Size = Size(50, 30)
        self.up_button.Click += self.move_up

        self.down_button = Button()
        self.down_button.Text = "Down"
        self.down_button.Location = Point(210, 220)
        self.down_button.Size = Size(50, 30)
        self.down_button.Click += self.move_down

        # OK Button
        self.ok_button = Button()
        self.ok_button.Text = "OK"
        self.ok_button.Location = Point(200, 320)
        self.ok_button.Size = Size(100, 30)
        self.ok_button.Click += self.ok_clicked

        # Add Controls
        self.Controls.Add(self.label_available)
        self.Controls.Add(self.label_selected)
        self.Controls.Add(self.available_listbox)
        self.Controls.Add(self.selected_listbox)
        self.Controls.Add(self.add_button)
        self.Controls.Add(self.remove_button)
        self.Controls.Add(self.up_button)
        self.Controls.Add(self.down_button)
        self.Controls.Add(self.ok_button)

        self.selected_params = []

    def add_selected(self, sender, args):
        selected_items = list(self.available_listbox.SelectedItems)
        for item in selected_items:
            if item not in self.selected_listbox.Items:
                self.selected_listbox.Items.Add(item)

    def remove_selected(self, sender, args):
        selected_items = list(self.selected_listbox.SelectedItems)
        for item in selected_items:
            self.selected_listbox.Items.Remove(item)

    def move_up(self, sender, args):
        selected_indices = list(self.selected_listbox.SelectedIndices)
        if not selected_indices:
            return
        for index in selected_indices:
            if index > 0:
                item = self.selected_listbox.Items[index]
                self.selected_listbox.Items.RemoveAt(index)
                self.selected_listbox.Items.Insert(index - 1, item)
                self.selected_listbox.SetSelected(index - 1, True)

    def move_down(self, sender, args):
        selected_indices = list(reversed(self.selected_listbox.SelectedIndices))
        if not selected_indices:
            return
        for index in selected_indices:
            if index < self.selected_listbox.Items.Count - 1:
                item = self.selected_listbox.Items[index]
                self.selected_listbox.Items.RemoveAt(index)
                self.selected_listbox.Items.Insert(index + 1, item)
                self.selected_listbox.SetSelected(index + 1, True)

    def ok_clicked(self, sender, args):
        self.selected_params = [self.selected_listbox.Items[i] for i in range(self.selected_listbox.Items.Count)]
        self.Close()

# Main execution
def main():
    # Initial dialog to choose between export and import
    options = ["Export Title Blocks", "Import Title Blocks"]
    choice = forms.alert("Select an action:", options=options, title="Title Block Manager")
    
    if choice == "Export Title Blocks":
        # Collect all title blocks
        title_blocks = list(DB.FilteredElementCollector(revit.doc)
            .OfCategory(DB.BuiltInCategory.OST_TitleBlocks)
            .WhereElementIsNotElementType()
            .ToElements())

        if not title_blocks:
            forms.alert("No title blocks found in the project.")
            return

        # Show title block selection dialog
        tb_dialog = TitleBlockSelectionDialog(title_blocks)
        tb_dialog.ShowDialog()

        if not tb_dialog.selected_title_blocks:
            forms.alert("No title blocks selected for export.")
            return

        # Collect available parameters from selected title blocks
        param_names = sorted(set(param.Definition.Name for tb in tb_dialog.selected_title_blocks for param in tb.Parameters))

        # Show parameter selection dialog
        param_dialog = ParameterSelectionDialog(param_names)
        param_dialog.ShowDialog()

        if not param_dialog.selected_params:
            forms.alert("No parameters selected for export.")
            return

        # Use the selected family name as the default file name
        default_file_name = tb_dialog.selected_family_name if tb_dialog.selected_family_name else "TitleBlockExport"
        filepath = forms.save_file(file_ext="xlsx", default_name=default_file_name, init_dir=last_export_dir)
        if filepath:
            export_titleblocks_to_excel(tb_dialog.selected_title_blocks, filepath, param_dialog.selected_params)
            
            # Show dialog to ask if user wants to open the file
            result = MessageBox.Show("Export complete. Would you like to open the file?", 
                                   "Open File", 
                                   MessageBoxButtons.OKCancel)
            if result == DialogResult.OK:
                os.startfile(filepath)
    
    elif choice == "Import Title Blocks":
        # Prompt user to select Excel file for import
        filepath = forms.pick_file(file_ext="xlsx", init_dir=last_export_dir, title="Select Excel File to Import")
        if filepath:
            import_changes_from_excel(filepath)

if __name__ == "__main__":
    main()