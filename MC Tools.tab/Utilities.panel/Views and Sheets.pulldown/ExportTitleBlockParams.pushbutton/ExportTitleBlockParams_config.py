# -*- coding: utf-8 -*-
import os
import clr

from pyrevit import revit, DB, forms
from pyrevit.compat import get_elementid_from_value_func
from pyrevit.interop import xl

clr.AddReference('RevitAPIUI')
from Autodesk.Revit.UI import TaskDialog

doc = revit.doc
get_elementid_from_value = get_elementid_from_value_func()


def show_message(msg, title="Title Block Import"):
    TaskDialog.Show(title, msg)


def parse_element_id(value):
    if value is None or value == "":
        return None
    try:
        return int(float(value))
    except Exception:
        try:
            return int(str(value).strip())
        except Exception:
            return None


def row_to_dict(headers, row):
    data = {}
    for i, header in enumerate(headers):
        key = str(header).strip()
        value = row[i] if i < len(row) else None
        data[key] = value
    return data


def get_candidate_elements(tb):
    candidates = [tb]

    try:
        if hasattr(tb, "Symbol") and tb.Symbol:
            candidates.append(tb.Symbol)
    except Exception:
        pass

    try:
        if tb.OwnerViewId and tb.OwnerViewId != DB.ElementId.InvalidElementId:
            owner = doc.GetElement(tb.OwnerViewId)
            if owner:
                candidates.append(owner)   # usually ViewSheet
    except Exception:
        pass

    return candidates


def find_param(tb, param_name):
    for obj in get_candidate_elements(tb):
        try:
            p = obj.LookupParameter(param_name)
            if p:
                return obj, p
        except Exception:
            pass
    return None, None


def set_parameter_value(param, value):
    if not param or param.IsReadOnly or value is None or value == "":
        return False, "read-only or blank"

    try:
        st = param.StorageType

        if st == DB.StorageType.String:
            param.Set(str(value))
            return True, None

        elif st == DB.StorageType.Integer:
            if isinstance(value, bool):
                param.Set(1 if value else 0)
            else:
                param.Set(int(float(value)))
            return True, None

        elif st == DB.StorageType.Double:
            try:
                param.SetValueString(str(value))
                return True, None
            except Exception:
                pass

            try:
                param.Set(float(value))
                return True, None
            except Exception:
                return False, "double conversion failed"

        elif st == DB.StorageType.ElementId:
            try:
                param.Set(get_elementid_from_value(int(float(value))))
                return True, None
            except Exception:
                return False, "elementid conversion failed"

        return False, "unsupported storage type"

    except Exception as ex:
        return False, str(ex)


def import_changes_from_excel(filepath):
    if not os.path.exists(filepath):
        show_message("Excel file not found:\n{}".format(filepath))
        return

    try:
        workbook_data = xl.load(filepath)
    except Exception as ex:
        show_message("Failed to read Excel file:\n{}".format(str(ex)))
        return

    if not workbook_data:
        show_message("No data found in Excel file.")
        return

    updated_params = 0
    updated_elements = set()
    skipped = []
    rows_found = 0

    with revit.Transaction("Update Title Blocks from Excel"):
        for sheet_name, sheet_data in workbook_data.items():
            headers = sheet_data.get("headers", [])
            rows = sheet_data.get("rows", [])

            if not headers or not rows:
                continue

            for row_index, row in enumerate(rows, start=2):
                if isinstance(row, dict):
                    row_data = {str(k).strip(): v for k, v in row.items()}
                elif isinstance(row, list):
                    row_data = row_to_dict(headers, row)
                else:
                    skipped.append("Sheet '{}', row {}: unsupported row type".format(sheet_name, row_index))
                    continue

                elem_raw = row_data.get("ElementId")
                elem_id_val = parse_element_id(elem_raw)
                if elem_id_val is None:
                    skipped.append("Sheet '{}', row {}: invalid ElementId".format(sheet_name, row_index))
                    continue

                tb = doc.GetElement(get_elementid_from_value(elem_id_val))
                if not tb:
                    skipped.append("Sheet '{}', row {}: element {} not found".format(sheet_name, row_index, elem_id_val))
                    continue

                rows_found += 1
                row_changed = False

                for header, value in row_data.items():
                    if header == "ElementId":
                        continue

                    owner_obj, param = find_param(tb, header)
                    if not param:
                        skipped.append("Element {} / '{}': parameter not found".format(elem_id_val, header))
                        continue

                    ok, reason = set_parameter_value(param, value)
                    if ok:
                        updated_params += 1
                        row_changed = True
                    else:
                        skipped.append("Element {} / '{}': {}".format(elem_id_val, header, reason))

                if row_changed:
                    updated_elements.add(elem_id_val)

    if skipped:
        print("\n".join(skipped[:500]))

    show_message("Imported Changes Successful")


def main():
    filepath = forms.pick_excel_file(title="Select Excel File to Import")
    if filepath:
        import_changes_from_excel(filepath)


if __name__ == "__main__":
    main()