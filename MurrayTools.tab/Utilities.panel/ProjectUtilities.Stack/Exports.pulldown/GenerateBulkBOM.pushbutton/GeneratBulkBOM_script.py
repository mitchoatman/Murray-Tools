# -*- coding: utf-8 -*-

from pyrevit import forms, revit, DB, script
import clr
import os.path as op
import csv
import tempfile
import os

# Add Windows Forms references
clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')
from System.Windows.Forms import Form, Label, Button, DialogResult, FormBorderStyle, FormStartPosition
from System.Drawing import Point, Size, Color
import System.Diagnostics  # For Process.Start to open the file

# Add reference to Excel COM interop and COM release
clr.AddReference("Microsoft.Office.Interop.Excel")
clr.AddReference("System.Runtime.InteropServices")
from Microsoft.Office.Interop.Excel import ApplicationClass
from System.Runtime.InteropServices import Marshal

logger = script.get_logger()

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
def export_schedule_to_excel(schedule, source_excel_path, output_filepath):
    excel = ApplicationClass()
    excel.Visible = False
    wb = None
    ws = None
    
    try:
        # Create a temporary CSV file
        with tempfile.NamedTemporaryFile(delete=False, suffix='.csv') as temp_csv:
            temp_csv_path = temp_csv.name

        # Configure export options for visible fields only, no headers
        vseop = DB.ViewScheduleExportOptions()
        vseop.ColumnHeaders = DB.ExportColumnHeaders.None
        vseop.FieldDelimiter = ','
        vseop.Title = False
        vseop.HeadersFootersBlanks = True

        # Export

        # Export schedule to temporary CSV (only visible fields, no ElementId, no headers)
        schedule.Export(op.dirname(temp_csv_path), op.basename(temp_csv_path), vseop)
        revit.files.correct_text_encoding(temp_csv_path)

        # Open the source MR-BOM.xlsx file
        wb = excel.Workbooks.Open(source_excel_path)
        ws = wb.Worksheets[1]
        ws.Name = schedule.Name[:31]  # Excel sheet names have a 31-char limit

        # Read CSV and write to Excel starting at row 17
        with open(temp_csv_path, 'r') as csv_file:
            csv_reader = csv.reader(csv_file)
            row = 17
            for csv_row in csv_reader:
                for col, value in enumerate(csv_row, 1):
                    # Format column C (size column) as text
                    if col == 3:
                        ws.Cells[row, col].NumberFormat = '@'  # Set format to text
                    # Skip writing '0' or '0"' in column C (col 3)
                    if col == 3 and value in ['0', '0"']:
                        ws.Cells[row, col].Value2 = ''
                    else:
                        ws.Cells[row, col].Value2 = value
                row += 1

        # Sort data in columns A to G, starting at row 17, by column B (descending)
        last_row = ws.Cells(ws.Rows.Count, 1).End(-4162).Row  # Find last row (xlUp)
        if last_row >= 17:  # Ensure there is data to sort
            range_to_sort = ws.Range(ws.Cells(17, 1), ws.Cells(last_row, 7))  # A17:G<last_row>
            range_to_sort.Sort(
                Key1=ws.Cells(17, 2),  # Sort by column B
                Order1=2,  # Descending (2 = xlDescending)
                Orientation=1  # Top to bottom (1 = xlTopToBottom)
            )

        # Find and replace '0"' with '' in column B, starting at row 17 (strict match)
        if last_row >= 17:  # Ensure there is data to process
            column_b_range = ws.Range(ws.Cells(17, 2), ws.Cells(last_row, 2))  # B17:B<last_row>
            column_b_range.Replace(
                What='0"',
                Replacement='',
                LookAt=2,  # xlWhole (2 = match entire cell contents)
                SearchOrder=1,  # xlByRows (1 = search by rows)
                MatchCase=False
            )

        # Suppress Excel overwrite prompt and save as new file
        excel.DisplayAlerts = False
        wb.SaveAs(output_filepath)
        excel.DisplayAlerts = True
        wb.Close()
        wb = None
        Marshal.ReleaseComObject(ws)
        ws = None

        # Delete temporary CSV
        os.remove(temp_csv_path)

        # Show dialog to open the file
        open_dialog = OpenFileDialog(output_filepath)
        open_dialog.ShowDialog()

    except Exception as e:
        # Specifically check for the COMException with HRESULT 0x800A03EC (file access error)
        if clr.GetClrType(type(e)).Name == "COMException" and "0x800A03EC" in str(e):
            forms.alert("Cannot save '{}'\nPlease close the file in Excel and try again.".format(op.basename(output_filepath)))
        else:
            forms.alert("Export failed:\n{}".format(str(e).split('\n')[0]))  # Show only first line of errors
        # Cleanup in case of error
        if wb is not None:
            wb.Close(SaveChanges=False)
            Marshal.ReleaseComObject(wb)
            wb = None
        if ws is not None:
            Marshal.ReleaseComObject(ws)
            ws = None
            
    finally:
        if excel is not None:
            excel.Quit()
            Marshal.ReleaseComObject(excel)
            excel = None

# Main execution
def main():
    # Get the script's directory
    script_dir, _ = os.path.split(__file__)
    
    # Construct path to MR-BOM.xlsx in the script's folder
    source_excel_path = os.path.join(script_dir, "MR-BOM.xlsx")
    
    # Check if MR-BOM.xlsx exists in the script's folder
    if not os.path.exists(source_excel_path):
        forms.alert("MR-BOM.xlsx not found in script's folder:\n{}".format(script_dir))
        return

    # Prompt user to select schedules
    schedules_to_export = forms.select_schedules(title="Select Schedules to Export")
    if not schedules_to_export:
        forms.alert("No schedules selected. Export cancelled.")
        return

    # Process each selected schedule
    for sched in schedules_to_export:
        # Prompt user for output file name, defaulting to Desktop
        default_name = "{}.xlsx".format(revit.query.get_name(sched))
        desktop_path = os.path.expandvars("%USERPROFILE%\\Desktop")
        output_filepath = forms.save_file(
            file_ext="xlsx",
            default_name=default_name,
            init_dir=desktop_path,
            title="Save Schedule Export As",
            unc_paths=False
        )
        if output_filepath:
            export_schedule_to_excel(sched, source_excel_path, output_filepath)

if __name__ == "__main__":
    main()