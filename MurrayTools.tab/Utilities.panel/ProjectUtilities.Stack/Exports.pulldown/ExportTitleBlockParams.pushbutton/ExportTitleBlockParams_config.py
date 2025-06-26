# -*- coding: utf-8 -*-
from pyrevit import revit, DB, forms
import clr
clr.AddReference('System.Windows.Forms')
from System.Windows.Forms import MessageBox, MessageBoxButtons, DialogResult
import os

clr.AddReference("Microsoft.Office.Interop.Excel")
from Microsoft.Office.Interop.Excel import ApplicationClass

last_export_dir = os.path.expanduser("~") + "\\Desktop"

def import_changes_from_excel(filepath):
    if not os.path.exists(filepath):
        forms.alert("Excel file not found at:\n{}".format(filepath))
        return
    
    excel = ApplicationClass()
    excel.Visible = False
    wb = excel.Workbooks.Open(filepath)
    ws = wb.Worksheets[1]
    
    headers = []
    col = 1
    while ws.Cells[1, col].Value2:
        headers.append(ws.Cells[1, col].Value2)
        col += 1
    header_map = {h: i+1 for i, h in enumerate(headers)}

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
                    elem_id = int(float(elem_id))
                    element = revit.doc.GetElement(DB.ElementId(elem_id))
                    if element:
                        for header, col_idx in header_map.items():
                            if header != "ElementId":
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
                                        except Exception, e:
                                            print "Failed to set param {} for ElementId {}: {}".format(header, elem_id, e)
                except Exception, e:
                    print "Failed to process ElementId {}: {}".format(elem_id, e)
                    continue
    
    wb.Close()
    excel.Quit()
    forms.alert("Import complete")

def main():
    filepath = forms.pick_file(file_ext="xlsx", init_dir=last_export_dir, title="Select Excel File to Import")
    if filepath:
        import_changes_from_excel(filepath)

if __name__ == "__main__":
    main()