# -*- coding: utf-8 -*-

from pyrevit import forms, revit, DB, script
import clr
import os.path as op
import csv
import os
from datetime import datetime  # Import datetime for today's date
import codecs  # Added for Python 2 encoding support

doc = __revit__.ActiveUIDocument.Document
app = doc.Application
RevitVersion = app.VersionNumber
RevitINT = float(RevitVersion)

if RevitINT < 2025:
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
            self.open_button.Location = Point(30, 80)  # Adjusted for centering
            self.open_button.Size = Size(100, 25)
            self.open_button.Click += self.open_file_clicked

            # Continue button (was Nah Dawg)
            self.continue_button = Button()
            self.continue_button.Text = "Continue"
            self.continue_button.Location = Point(140, 80)  # Adjusted for centering
            self.continue_button.Size = Size(100, 25)
            self.continue_button.Click += self.continue_clicked

            # Add controls to form
            self.Controls.Add(self.label)
            self.Controls.Add(self.open_button)
            self.Controls.Add(self.continue_button)

        def open_file_clicked(self, sender, args):
            try:
                System.Diagnostics.Process.Start(self.filepath)  # Open the file with default application
            except Exception, e:
                forms.alert("Failed to open file:\n%s" % str(e))
            self.Close()

        def continue_clicked(self, sender, args):
            self.Close()  # Close dialog and continue to next schedule or do nothing

    # Function to export schedule to Excel
    def export_schedule_to_excel(schedule, source_excel_path, output_filepath, doc):
        excel = ApplicationClass()
        excel.Visible = False
        wb = None
        ws = None
        
        try:
            # Ensure C:\temp exists
            temp_dir = r'C:\temp'
            if not os.path.exists(temp_dir):
                os.makedirs(temp_dir)

            # Create a temporary CSV file in C:\temp
            temp_csv_path = os.path.join(temp_dir, "temp_schedule_%s.csv" % datetime.now().strftime('%Y%m%d_%H%M%S'))

            # Configure export options for visible fields only, no headers
            vseop = DB.ViewScheduleExportOptions()
            vseop.ColumnHeaders = DB.ExportColumnHeaders.None
            vseop.FieldDelimiter = ','
            vseop.Title = False
            vseop.HeadersFootersBlanks = True

            # Export schedule to temporary CSV (only visible fields, no ElementId, no headers)
            schedule.Export(op.dirname(temp_csv_path), op.basename(temp_csv_path), vseop)
            revit.files.correct_text_encoding(temp_csv_path)

            # Open the source MR-BOM.xlsx file
            wb = excel.Workbooks.Open(source_excel_path)
            ws = wb.Worksheets[1]
            ws.Name = schedule.Name[:31]  # Excel sheet names have a 31-char limit

            # Get project information and today's date
            project_info = doc.ProjectInformation
            project_name = project_info.Name
            project_number = project_info.Number
            today_date = datetime.today().strftime('%Y-%m-%d')  # Format as YYYY-MM-DD

            # Write project name, number
            ws.Cells[7, 6].Value2 = today_date
            ws.Cells[8, 6].Value2 = project_name
            ws.Cells[9, 6].Value2 = project_number

            # Format columns B and C as text for rows 17 to 1000
            ws.Range(ws.Cells(17, 2), ws.Cells(1000, 2)).NumberFormat = '@'  # Column B (length)
            ws.Range(ws.Cells(17, 3), ws.Cells(1000, 3)).NumberFormat = '@'  # Column C (size)

            # Read CSV and write to Excel starting at row 17
            with codecs.open(temp_csv_path, 'r', encoding='utf-8-sig') as csv_file:
                csv_reader = csv.reader(csv_file)
                row = 17
                for csv_row in csv_reader:
                    for col, value in enumerate(csv_row, 1):
                        # Skip writing '0' - 0"' in column C (col 3)
                        if col == 3 and value == "0' - 0\"":
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

            # Find and replace '0' - 0"' with '' in column B, starting at row 17 (strict match)
            if last_row >= 17:  # Ensure there is data to process
                for row in range(17, last_row + 1):
                    cell = ws.Cells(row, 2)
                    cell_value = cell.Value() if cell.Value() else ""
                    if cell_value == "0' - 0\"":
                        cell.Value = ''

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

        except Exception, e:
            # Specifically check for the COMException with HRESULT 0x800A03EC (file access error)
            if clr.GetClrType(type(e)).Name == "COMException" and "0x800A03EC" in str(e):
                forms.alert("Cannot save '%s'\nPlease close the file in Excel and try again." % op.basename(output_filepath))
            else:
                forms.alert("Export failed:\n%s" % str(e).split('\n')[0])  # Show only first line of errors
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
        # Ensure C:\temp exists
        temp_dir = r'C:\temp'
        if not os.path.exists(temp_dir):
            os.makedirs(temp_dir)

        # Check for stored path in C:\temp\Ribbon_GenerateBOM.txt
        path_file = os.path.join(temp_dir, "Ribbon_GenerateBOM.txt")
        desktop_path = os.path.expandvars("%USERPROFILE%\\Desktop")
        default_dir = desktop_path  # Default to Desktop

        if os.path.exists(path_file):
            try:
                with open(path_file, 'r') as f:
                    stored_path = f.read().strip()
                    if os.path.isdir(stored_path):
                        default_dir = stored_path
            except Exception, e:
                print("Error reading path from %s: %s" % (path_file, str(e)))
                default_dir = desktop_path

        # Get the script's directory
        script_dir, _ = os.path.split(__file__)
        
        # Construct path to MR-BOM.xlsx in the script's folder
        source_excel_path = os.path.join(script_dir, "MR-BOM.xlsx")
        
        # Check if MR-BOM.xlsx exists in the script's folder
        if not os.path.exists(source_excel_path):
            forms.alert("MR-BOM.xlsx not found in script's folder:\n%s" % script_dir)
            return

        # Prompt user to select schedules
        schedules_to_export = forms.select_schedules(title="Select Schedules for BOM")
        if not schedules_to_export:
            forms.alert("No schedules selected. Export cancelled.")
            return

        # Get the active Revit document
        doc = revit.doc

        # Process each selected schedule
        for sched in schedules_to_export:
            # Prompt user for output file name, defaulting to stored path or Desktop
            default_name = "%s.xlsx" % revit.query.get_name(sched)
            output_filepath = forms.save_file(
                file_ext="xlsx",
                default_name=default_name,
                init_dir=default_dir,
                title="BOM Save Location",
                unc_paths=False
            )
            if output_filepath:
                # Get the selected folder path
                selected_dir = op.dirname(output_filepath)
                # Save the selected folder to Ribbon_GenerateBOM.txt (always)
                try:
                    with open(path_file, 'w') as f:
                        f.write(selected_dir)
                except Exception, e:
                    print("Error writing path to %s: %s" % (path_file, str(e)))

                export_schedule_to_excel(sched, source_excel_path, output_filepath, doc)

    if __name__ == "__main__":
        main()

else:
    from pyrevit import forms
    from pyrevit import coreutils
    from pyrevit import revit, DB
    from pyrevit import script

    # Select schedules
    schedules_to_export = forms.select_schedules()

    if schedules_to_export:
        # Set up export options
        vseop = DB.ViewScheduleExportOptions()
        vseop.ColumnHeaders = coreutils.get_enum_value(DB.ExportColumnHeaders, "OneRow")
        vseop.FieldDelimiter = ','
        vseop.Title = False
        vseop.HeadersFootersBlanks = True

        # Define export directory
        export_dir = r'C:\temp'
        if not os.path.exists(export_dir):
            os.makedirs(export_dir)

        # Export each schedule
        for sched in schedules_to_export:
            # Generate clean filename
            fname = coreutils.cleanup_filename(revit.query.get_name(sched)) + '.csv'
            export_path = op.join(export_dir, fname)
            
            # Export schedule
            try:
                sched.Export(export_dir, fname, vseop)
                # Correct text encoding if needed
                revit.files.correct_text_encoding(export_path)
            except Exception, e:
                print("Error exporting schedule %s: %s" % (revit.query.get_name(sched), str(e)))

    path, filename = os.path.split(__file__)
    NewFilename = '\MR-BOM.xlsm'

    # Open the BOM template
    bom_template = path + NewFilename
    if op.exists(bom_template):
        os.startfile(bom_template)
    else:
        print("BOM template file not found at: %s" % bom_template)