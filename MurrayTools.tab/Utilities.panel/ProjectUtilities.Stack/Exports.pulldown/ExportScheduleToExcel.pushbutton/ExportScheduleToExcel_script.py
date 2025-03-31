# -*- coding: utf-8 -*-
from pyrevit import revit, DB, forms, script
import clr
# Add Windows Forms references
clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')
from System.Windows.Forms import (Form, Label, ListBox, Button, DialogResult, TextBox,
                                FormBorderStyle, FormStartPosition, SelectionMode, Control, AnchorStyles)
from System import Array
from System.Drawing import Point, Size, Color
import os
import System.Diagnostics  # For Process.Start to open the file

# Add reference to Excel COM interop
clr.AddReference("Microsoft.Office.Interop.Excel")
from Microsoft.Office.Interop.Excel import ApplicationClass

# Global variable to store the last exported directory
last_export_dir = os.path.expanduser("~") + "\\Desktop"  # Default to Desktop

# Simple dialog for opening the exported file
class OpenFileDialog(Form):
    def __init__(self, filepath):
        self.filepath = filepath
        self.Text = "Export Successful"
        self.Width = 300
        self.Height = 150
        self.FormBorderStyle = FormBorderStyle.FixedDialog
        self.MaximizeBox = False
        self.MinimizeBox = False
        self.StartPosition = FormStartPosition.CenterScreen

        # Success message
        self.label = Label()
        self.label.Text = "Schedule exported successfully.\nWould you like to open the file?"
        self.label.Location = Point(10, 10)
        self.label.Size = Size(260, 40)
        self.label.TextAlign = System.Drawing.ContentAlignment.MiddleCenter

        # Open File button
        self.open_button = Button()
        self.open_button.Text = "Open File"
        self.open_button.Location = Point(100, 80)
        self.open_button.Size = Size(100, 25)
        self.open_button.Click += self.open_file_clicked

        # Add controls to form
        self.Controls.Add(self.label)
        self.Controls.Add(self.open_button)

    def open_file_clicked(self, sender, args):
        try:
            System.Diagnostics.Process.Start(self.filepath)  # Open the file with default application
        except Exception as e:
            forms.alert("Failed to open file:\n{}".format(str(e)))
        self.Close()

# Function to export schedule to Excel
def export_schedule_to_excel(schedule, filepath):
    global last_export_dir
    excel = ApplicationClass()
    excel.Visible = False
    wb = None  # Initialize workbook as None
    
    try:
        wb = excel.Workbooks.Add()
        ws = wb.Worksheets[1]
        ws.Name = schedule.Name[:31]  # Excel sheet names have a 31-char limit

        # Collect schedule data and element parameters
        schedule_view = schedule
        elements = DB.FilteredElementCollector(revit.doc, schedule_view.Id).ToElements()
        schedule_def = schedule_view.Definition
        fields = [schedule_def.GetField(fid) for fid in schedule_def.GetFieldOrder()]

        # Write headers (ElementId + parameter names)
        headers = ["ElementId"] + [field.GetName() for field in fields if field.ParameterId != DB.ElementId.InvalidElementId]
        for col, header in enumerate(headers, 1):
            ws.Cells[1, col].Value2 = header

        # Write data (ElementId + parameter values)
        row = 2
        for element in elements:
            ws.Cells[row, 1].Value2 = str(element.Id.IntegerValue)
            col = 2
            for field in fields:
                if field.ParameterId != DB.ElementId.InvalidElementId:
                    param = None
                    field_name = field.GetName()
                    param = element.LookupParameter(field_name)

                    # Try BuiltInParameter if LookupParameter doesn't work
                    if not param:
                        try:
                            bip = DB.BuiltInParameter(field.ParameterId.IntegerValue)
                            param = element.get_Parameter(bip)
                        except:
                            pass

                    # Handle parameter values
                    if param:
                        value = None
                        if param.StorageType == DB.StorageType.String:
                            value = param.AsString()
                        elif param.StorageType == DB.StorageType.Double:
                            value = param.AsDouble()
                        elif param.StorageType == DB.StorageType.Integer:
                            value = param.AsInteger()
                        
                        # Special handling for Level Parameter (Reference Level)
                        if param.Definition.Name.lower() in ["level", "reference level", "schedule level"]:
                            level_id = param.AsElementId()
                            if level_id and level_id.IntegerValue > 0:
                                level = revit.doc.GetElement(level_id)
                                value = level.Name if level else "Unknown Level"

                        ws.Cells[row, col].Value2 = value if value is not None else ""
                    col += 1
            row += 1

        # Save and cleanup
        wb.SaveAs(filepath)
        wb.Close()
        last_export_dir = os.path.dirname(filepath)  # Update last export directory
        
        # Show dialog to open the file
        open_dialog = OpenFileDialog(filepath)
        open_dialog.ShowDialog()

    except Exception as e:
        # Specifically check for the COMException with HRESULT 0x800A03EC (file access error)
        if clr.GetClrType(type(e)).Name == "COMException" and "0x800A03EC" in str(e):
            forms.alert("Cannot save '{}'\nPlease close the file in Excel and try again.".format(os.path.basename(filepath)))
        else:
            forms.alert("Export failed:\n{}".format(str(e).split('\n')[0]))  # Show only first line of other errors
        # Cleanup in case of error
        if wb is not None:
            wb.Close(SaveChanges=False)
            
    finally:
        excel.Quit()

# Function to import changes from Excel
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

    # Collect all levels in the project for quick lookup
    level_map = {lvl.Name: lvl for lvl in DB.FilteredElementCollector(revit.doc).OfClass(DB.Level)}

    # Import data
    row_count = ws.UsedRange.Rows.Count
    if row_count <= 1:
        wb.Close()
        excel.Quit()
        forms.alert("No data found in Excel file.")
        return
    
    updated_count = 0
    with revit.Transaction("Update Elements from Excel"):
        for row in range(2, row_count + 1):
            elem_id = ws.Cells[row, header_map["ElementId"]].Value2
            if elem_id:
                try:
                    elem_id = int(float(elem_id))  # Excel might return as float
                    element = revit.doc.GetElement(DB.ElementId(elem_id))
                    if element:
                        for header, col_idx in header_map.items():
                            if header != "ElementId":
                                param = element.LookupParameter(header)

                                # Special handling for Reference Level
                                if param and param.Definition.Name.lower() in ["level", "reference level", "schedule level"]:
                                    level_name = ws.Cells[row, col_idx].Value2
                                    if level_name in level_map:
                                        level_id = level_map[level_name].Id
                                        param.Set(level_id)
                                        updated_count += 1
                                    else:
                                        print "Level '{}' not found for Element {}".format(level_name, elem_id)

                                # Standard parameter handling
                                elif param and not param.IsReadOnly:
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
                                            print "Failed to set param {} for ElementId {}: {}".format(header, elem_id, e)
                except Exception as e:
                    print "Failed to process ElementId {}: {}".format(elem_id, e)
                    continue
    
    # Cleanup
    wb.Close()
    excel.Quit()
    # forms.alert("Import complete. Updated {} parameters.".format(updated_count))



# Custom WinForms dialog with ListBox and buttons
from System.Windows.Forms import TextBox

class ScheduleDialog(Form):
    def __init__(self, schedules):
        self.schedules = schedules  # Store schedules for filtering

        self.Text = "Schedule Import/Export"
        self.Width = 400
        self.Height = 300
        self.FormBorderStyle = FormBorderStyle.FixedDialog
        self.MaximizeBox = False
        self.MinimizeBox = False
        self.StartPosition = FormStartPosition.CenterScreen

        # Search label
        self.search_label = Label()
        self.search_label.Text = "Search:"
        self.search_label.Location = Point(10, 5)
        self.search_label.Size = Size(50, 20)

        # Search box
        self.search_box = TextBox()
        self.search_box.Location = Point(70, 5)
        self.search_box.Size = Size(300, 20)
        self.search_box.TextChanged += self.search_changed

        # ListBox for schedules
        self.schedule_listbox = ListBox()
        self.schedule_listbox.Location = Point(10, 30)
        self.schedule_listbox.Size = Size(360, 180)
        self.schedule_listbox.SelectionMode = SelectionMode.One

        # Populate list with schedule names
        for schedule in self.schedules:
            self.schedule_listbox.Items.Add(schedule.Name)

        # Buttons
        self.export_button = Button()
        self.export_button.Text = "Export to Excel"
        self.export_button.Location = Point(150, 230)
        self.export_button.Size = Size(100, 25)
        self.export_button.Click += self.export_clicked

        self.import_button = Button()
        self.import_button.Text = "Import from Excel"
        self.import_button.Location = Point(260, 230)
        self.import_button.Size = Size(120, 25)
        self.import_button.Click += self.import_clicked

        # Add controls to form
        self.Controls.Add(self.search_label)
        self.Controls.Add(self.search_box)
        self.Controls.Add(self.schedule_listbox)
        self.Controls.Add(self.export_button)
        self.Controls.Add(self.import_button)

        self.selected_schedule = None
        self.action = None

    def search_changed(self, sender, args):
        search_term = self.search_box.Text.lower()
        self.schedule_listbox.Items.Clear()
        
        # Filter schedules by search term
        filtered_schedules = [s.Name for s in self.schedules if search_term in s.Name.lower()]
        
        # Update listbox with filtered results
        self.schedule_listbox.Items.AddRange(Array[object](filtered_schedules))

    def export_clicked(self, sender, args):
        if self.schedule_listbox.SelectedIndex != -1:
            self.selected_schedule = self.schedule_listbox.SelectedItem
            self.action = "Export"
            self.Close()
        else:
            forms.alert("Please select a schedule to export.")

    def import_clicked(self, sender, args):
        self.action = "Import"
        self.Close()


# Main execution
def main():
    global last_export_dir
    # Get available schedules
    schedules = [s for s in DB.FilteredElementCollector(revit.doc).OfClass(DB.ViewSchedule).ToElements() if not s.IsTemplate]
    if not schedules:
        forms.alert("No schedules found in the project.")
        return
    
    # Show custom dialog with ListBox and buttons
    dialog = ScheduleDialog(schedules)
    dialog.ShowDialog()
    
    if not dialog.action:
        return
    
    if dialog.action == "Export":
        # Get the selected schedule object based on name
        selected_schedule_name = dialog.selected_schedule
        selected_schedule = next(s for s in schedules if s.Name == selected_schedule_name)
        # Use Save As dialog for export location
        default_name = "{}.xlsx".format(selected_schedule.Name)
        filepath = forms.save_file(
            file_ext="xlsx",
            default_name=default_name,
            init_dir=last_export_dir,
            title="Save Schedule Export As",
            unc_paths=False
        )
        if filepath:
            export_schedule_to_excel(selected_schedule, filepath)
    elif dialog.action == "Import":
        # Use Open dialog with last export dir as default
        filepath = forms.pick_file(
            file_ext="xlsx",
            init_dir=last_export_dir,
            title="Select Excel File to Import"
        )
        if filepath:
            import_changes_from_excel(filepath)

if __name__ == "__main__":
    main()