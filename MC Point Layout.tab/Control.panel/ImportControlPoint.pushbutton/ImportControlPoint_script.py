# -*- coding: utf-8 -*-

import clr
import os
import csv
import sys

clr.AddReference('System')
clr.AddReference('System.Drawing')
clr.AddReference('System.Windows.Forms')

from System.Drawing import Point, Size
from System.Windows.Forms import (
    Form, Label, TextBox, Button, GroupBox, RadioButton,
    OpenFileDialog, DialogResult, FormStartPosition,
    FormBorderStyle, MessageBox, MessageBoxButtons, MessageBoxIcon
)

from Autodesk.Revit import DB
from Autodesk.Revit.DB import (
    FilteredElementCollector,
    Family,
    FamilySymbol,
    Transaction,
    Level,
    XYZ,
    Structure
)
from Autodesk.Revit.UI import TaskDialog


doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

FAMILY_NAME = "Control Point"
DEFAULT_TYPE_NAME = "CP"

SETTINGS_FOLDER = r"C:\Temp"
SETTINGS_FILE = os.path.join(SETTINGS_FOLDER, "ImportControlPoints.txt")

SCRIPT_DIR = os.path.dirname(__file__)
CONTROL_POINT_FAMILY_PATH = os.path.join(SCRIPT_DIR, "Control Point.rfa")


# ------------------------------------------------------------
# SETTINGS
# ------------------------------------------------------------
def load_settings():
    settings = {
        "csv_path": "",
        "type_name": DEFAULT_TYPE_NAME,
        "coordinate_system": "SHARED"
    }

    try:
        if os.path.exists(SETTINGS_FILE):
            with open(SETTINGS_FILE, "r") as f:
                for line in f.readlines():
                    line = line.strip()
                    if not line or "=" not in line:
                        continue
                    key, value = line.split("=", 1)
                    key = key.strip()
                    value = value.strip()

                    if key == "csv_path":
                        settings["csv_path"] = value
                    elif key == "type_name":
                        settings["type_name"] = value
                    elif key == "coordinate_system":
                        settings["coordinate_system"] = value
    except:
        pass

    return settings


def save_settings(settings):
    try:
        if not os.path.exists(SETTINGS_FOLDER):
            os.makedirs(SETTINGS_FOLDER)

        with open(SETTINGS_FILE, "w") as f:
            f.writelines([
                "csv_path={}\n".format(settings.get("csv_path", "")),
                "type_name={}\n".format(settings.get("type_name", DEFAULT_TYPE_NAME)),
                "coordinate_system={}\n".format(settings.get("coordinate_system", "SHARED"))
            ])
    except:
        pass


# ------------------------------------------------------------
# HELPERS
# ------------------------------------------------------------
def normalize_header(value):
    if value is None:
        return ""
    try:
        s = str(value).strip()
    except:
        s = ""
    s = s.replace('\xef\xbb\xbf', '').replace(u'\ufeff', '')
    s = s.upper()
    return "".join([c for c in s if c.isalnum()])


def to_text(value):
    if value is None:
        return ""
    try:
        return str(value).strip()
    except:
        return ""


def to_float(value):
    if value is None:
        return None
    try:
        s = str(value).strip()
        s = s.replace(",", "")
        if not s:
            return None
        return float(s)
    except:
        return None


def set_param_string(element, param_name, value):
    try:
        p = element.LookupParameter(param_name)
        if p and (not p.IsReadOnly) and p.StorageType == DB.StorageType.String:
            p.Set(value if value else "")
            return True
    except:
        pass
    return False


def get_levels():
    try:
        lvls = list(FilteredElementCollector(doc).OfClass(Level))
        lvls.sort(key=lambda x: x.Elevation)
        return lvls
    except:
        return []


def get_nearest_level(z_val):
    levels = get_levels()
    if not levels:
        return None

    nearest = None
    min_diff = None
    for lvl in levels:
        diff = abs(lvl.Elevation - z_val)
        if min_diff is None or diff < min_diff:
            min_diff = diff
            nearest = lvl
    return nearest


def get_first_symbol_id(family):
    try:
        ids = family.GetFamilySymbolIds()
        for sid in ids:
            return sid
    except:
        pass
    return None


# ------------------------------------------------------------
# FAMILY LOADING
# ------------------------------------------------------------
class FamilyLoaderOptionsHandler(DB.IFamilyLoadOptions):
    def OnFamilyFound(self, familyInUse, overwriteParameterValues):
        overwriteParameterValues.Value = False
        return True

    def OnSharedFamilyFound(self, sharedFamily, familyInUse, source, overwriteParameterValues):
        source.Value = DB.FamilySource.Family
        overwriteParameterValues.Value = False
        return True


def find_family(family_name):
    for fam in FilteredElementCollector(doc).OfClass(Family):
        try:
            if fam.Name == family_name:
                return fam
        except:
            pass
    return None


def find_symbol(family_name, type_name):
    for sym in FilteredElementCollector(doc).OfClass(FamilySymbol):
        try:
            if sym.Family and sym.Family.Name == family_name:
                name_param = sym.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM)
                sym_name = name_param.AsString() if name_param else ""
                if sym_name == type_name:
                    return sym
        except:
            pass
    return None


def ensure_family_loaded():
    fam = find_family(FAMILY_NAME)
    if fam:
        return True

    if not os.path.exists(CONTROL_POINT_FAMILY_PATH):
        TaskDialog.Show("Missing Family", "Family file not found:\n{}".format(CONTROL_POINT_FAMILY_PATH))
        return False

    t = Transaction(doc, "Load Control Point Family")
    started = False
    try:
        t.Start()
        started = True
        handler = FamilyLoaderOptionsHandler()
        doc.LoadFamily(CONTROL_POINT_FAMILY_PATH, handler)
        t.Commit()
        return True
    except Exception as ex:
        if started:
            try:
                t.RollBack()
            except:
                pass
        TaskDialog.Show("Load Family Failed", str(ex))
        return False


def get_or_create_symbol(type_name):
    symbol = find_symbol(FAMILY_NAME, type_name)
    if symbol:
        return symbol

    fam = find_family(FAMILY_NAME)
    if not fam:
        return None

    base_symbol_id = get_first_symbol_id(fam)
    if not base_symbol_id:
        return None

    t = Transaction(doc, "Create Control Point Type")
    started = False
    try:
        t.Start()
        started = True
        base_symbol = doc.GetElement(base_symbol_id)
        new_symbol = base_symbol.Duplicate(type_name)
        t.Commit()
        return new_symbol
    except Exception as ex:
        if started:
            try:
                t.RollBack()
            except:
                pass
        TaskDialog.Show("Create Type Failed", "Could not create type '{}'\n\n{}".format(type_name, str(ex)))
        return None


# ------------------------------------------------------------
# CSV HEADER DETECTION
# ------------------------------------------------------------
X_HEADERS = set(["X", "E", "EAST", "EASTING"])
Y_HEADERS = set(["Y", "N", "NORTH", "NORTHING"])
Z_HEADERS = set(["Z", "H", "EL", "ELEV", "ELEVATION", "HEIGHT"])

NUMBER_HEADERS = set([
    "POINTNUMBER", "POINTNO", "POINTNUM", "POINTID",
    "NUMBER", "PTNO", "PTNUM", "TSPOINTNUMBER",
    "POINTNU"
])

DESCRIPTION_HEADERS = set([
    "DESCRIPTION", "DESC", "POINTDESCRIPTION",
    "POINTDESC", "TSPOINTDESCRIPTION",
    "DESCRIPTI"
])


def detect_columns(header_row):
    result = {
        "x": None,
        "y": None,
        "z": None,
        "number": None,
        "description": None
    }

    for idx, header in enumerate(header_row):
        h = normalize_header(header)

        if result["x"] is None and h in X_HEADERS:
            result["x"] = idx
            continue
        if result["y"] is None and h in Y_HEADERS:
            result["y"] = idx
            continue
        if result["z"] is None and h in Z_HEADERS:
            result["z"] = idx
            continue
        if result["number"] is None and h in NUMBER_HEADERS:
            result["number"] = idx
            continue
        if result["description"] is None and h in DESCRIPTION_HEADERS:
            result["description"] = idx
            continue

    return result


def get_row_value(row, idx):
    if idx is None:
        return ""
    if idx < 0 or idx >= len(row):
        return ""
    return row[idx]


def read_csv_rows(file_path):
    rows = []
    skipped = []

    with open(file_path, 'rb') as f:
        sample = f.read(4096)
        f.seek(0)

        try:
            dialect = csv.Sniffer().sniff(sample, delimiters=",;\t|")
        except:
            dialect = csv.excel

        reader = csv.reader(f, dialect)

        try:
            header_row = next(reader)
        except:
            raise Exception("CSV file is empty.")

        if not header_row:
            raise Exception("CSV header row is empty.")

        cols = detect_columns(header_row)

        if cols["x"] is None or cols["y"] is None or cols["z"] is None:
            raise Exception(
                "Could not detect X/Y/Z columns.\n\n"
                "Accepted X headers: X, E, EASTING\n"
                "Accepted Y headers: Y, N, NORTHING\n"
                "Accepted Z headers: Z, H, ELEVATION"
            )

        line_no = 1
        for row in reader:
            line_no += 1

            if not row:
                continue

            x_raw = get_row_value(row, cols["x"])
            y_raw = get_row_value(row, cols["y"])
            z_raw = get_row_value(row, cols["z"])

            if not str(x_raw).strip() and not str(y_raw).strip() and not str(z_raw).strip():
                continue

            x_val = to_float(x_raw)
            y_val = to_float(y_raw)
            z_val = to_float(z_raw)

            if x_val is None or y_val is None or z_val is None:
                skipped.append("Row {} skipped: invalid coordinate(s).".format(line_no))
                continue

            point_number = to_text(get_row_value(row, cols["number"]))
            description = to_text(get_row_value(row, cols["description"]))

            rows.append({
                "line_no": line_no,
                "x": x_val,
                "y": y_val,
                "z": z_val,
                "point_number": point_number,
                "description": description
            })

    return rows, skipped


# ------------------------------------------------------------
# COORDINATE CONVERSION
# ------------------------------------------------------------
def get_shared_basis():
    proj_loc = doc.ActiveProjectLocation

    p0 = proj_loc.GetProjectPosition(XYZ(0, 0, 0))
    px = proj_loc.GetProjectPosition(XYZ(1, 0, 0))
    py = proj_loc.GetProjectPosition(XYZ(0, 1, 0))

    basis = {
        "origin_e": p0.EastWest,
        "origin_n": p0.NorthSouth,
        "origin_z": p0.Elevation,

        "vx_e": px.EastWest - p0.EastWest,
        "vx_n": px.NorthSouth - p0.NorthSouth,

        "vy_e": py.EastWest - p0.EastWest,
        "vy_n": py.NorthSouth - p0.NorthSouth
    }

    det = (basis["vx_e"] * basis["vy_n"]) - (basis["vy_e"] * basis["vx_n"])
    basis["det"] = det
    return basis


def shared_to_internal(east, north, elev, basis):
    dx = east - basis["origin_e"]
    dy = north - basis["origin_n"]

    det = basis["det"]
    if abs(det) < 1e-12:
        raise Exception("Shared coordinate transform is not invertible.")

    x = ((dx * basis["vy_n"]) - (basis["vy_e"] * dy)) / det
    y = ((basis["vx_e"] * dy) - (dx * basis["vx_n"])) / det
    z = elev - basis["origin_z"]

    return XYZ(x, y, z)


def csv_row_to_xyz(row, coordinate_system, basis):
    if coordinate_system == "INTERNAL":
        return XYZ(row["x"], row["y"], row["z"])
    return shared_to_internal(row["x"], row["y"], row["z"], basis)


# ------------------------------------------------------------
# PLACEMENT
# ------------------------------------------------------------
def place_instance(symbol, point_xyz):
    try:
        return doc.Create.NewFamilyInstance(
            point_xyz,
            symbol,
            Structure.StructuralType.NonStructural
        )
    except:
        pass

    lvl = get_nearest_level(point_xyz.Z)
    if lvl:
        return doc.Create.NewFamilyInstance(
            point_xyz,
            symbol,
            lvl,
            Structure.StructuralType.NonStructural
        )

    raise Exception("Could not place family instance at point.")


# ------------------------------------------------------------
# UI
# ------------------------------------------------------------
class ImportControlPointsForm(Form):
    def __init__(self, settings):
        Form.__init__(self)

        self.Text = "Import Control Points from CSV"
        self.ClientSize = Size(480, 230)
        self.StartPosition = FormStartPosition.CenterScreen
        self.FormBorderStyle = FormBorderStyle.FixedDialog
        self.MaximizeBox = False
        self.MinimizeBox = False
        self.TopMost = True

        self.csv_path = settings.get("csv_path", "")
        self.type_name = settings.get("type_name", DEFAULT_TYPE_NAME)
        self.coordinate_system = settings.get("coordinate_system", "SHARED")
        self.result_ok = False

        self.init_ui()

    def init_ui(self):
        lbl_type = Label()
        lbl_type.Text = "Control Point Type:"
        lbl_type.Location = Point(15, 20)
        lbl_type.Size = Size(120, 20)
        self.Controls.Add(lbl_type)

        self.txt_type = TextBox()
        self.txt_type.Location = Point(145, 18)
        self.txt_type.Size = Size(180, 22)
        self.txt_type.Text = self.type_name
        self.txt_type.TextChanged += self.on_type_changed
        self.Controls.Add(self.txt_type)

        grp_coord = GroupBox()
        grp_coord.Text = "CSV Coordinate System"
        grp_coord.Location = Point(15, 55)
        grp_coord.Size = Size(310, 75)
        self.Controls.Add(grp_coord)

        self.rb_internal = RadioButton()
        self.rb_internal.Text = "Project Internal"
        self.rb_internal.Location = Point(15, 30)
        self.rb_internal.Size = Size(120, 20)
        self.rb_internal.Checked = (self.coordinate_system == "INTERNAL")
        grp_coord.Controls.Add(self.rb_internal)

        self.rb_shared = RadioButton()
        self.rb_shared.Text = "Shared"
        self.rb_shared.Location = Point(160, 30)
        self.rb_shared.Size = Size(100, 20)
        self.rb_shared.Checked = (self.coordinate_system == "SHARED")
        grp_coord.Controls.Add(self.rb_shared)

        left_margin = 20
        right_margin = 20
        label_w = 70
        gap = 8
        browse_w = 85

        lbl_file = Label()
        lbl_file.Text = "CSV File:"
        lbl_file.Location = Point(left_margin, 145)
        lbl_file.Size = Size(label_w, 20)
        self.Controls.Add(lbl_file)

        file_x = left_margin + label_w + gap
        browse_x = self.ClientSize.Width - right_margin - browse_w
        txt_w = browse_x - gap - file_x

        self.txt_file = TextBox()
        self.txt_file.Location = Point(file_x, 142)
        self.txt_file.Size = Size(txt_w, 22)
        self.txt_file.Text = self.csv_path
        self.Controls.Add(self.txt_file)

        btn_browse = Button()
        btn_browse.Text = "Browse..."
        btn_browse.Location = Point(browse_x, 140)
        btn_browse.Size = Size(browse_w, 26)
        btn_browse.Click += self.on_browse
        self.Controls.Add(btn_browse)

        button_y = 180
        gap = 10

        import_w = 90
        cancel_w = 90
        template_w = 120
        button_h = 28

        total_w = import_w + cancel_w + template_w + (gap * 2)
        start_x = (self.ClientSize.Width - total_w) // 2

        btn_import = Button()
        btn_import.Text = "Import"
        btn_import.Location = Point(start_x, button_y)
        btn_import.Size = Size(import_w, button_h)
        btn_import.Click += self.on_import
        self.Controls.Add(btn_import)

        btn_cancel = Button()
        btn_cancel.Text = "Cancel"
        btn_cancel.Location = Point(start_x + import_w + gap, button_y)
        btn_cancel.Size = Size(cancel_w, button_h)
        btn_cancel.Click += self.on_cancel
        self.Controls.Add(btn_cancel)

        btn_template = Button()
        btn_template.Text = "Make Template"
        btn_template.Location = Point(start_x + import_w + gap + cancel_w + gap, button_y)
        btn_template.Size = Size(template_w, button_h)
        btn_template.Click += self.on_make_template
        self.Controls.Add(btn_template)

    def on_type_changed(self, sender, args):
        try:
            current_text = self.txt_type.Text
            upper_text = current_text.upper()
            if current_text != upper_text:
                caret = self.txt_type.SelectionStart
                self.txt_type.Text = upper_text
                self.txt_type.SelectionStart = caret
        except:
            pass

    def on_browse(self, sender, args):
        dlg = OpenFileDialog()
        dlg.Title = "Select CSV File"
        dlg.Filter = "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*"
        dlg.Multiselect = False

        try:
            current_dir = os.path.dirname(self.txt_file.Text)
            if current_dir and os.path.exists(current_dir):
                dlg.InitialDirectory = current_dir
        except:
            pass

        if dlg.ShowDialog() == DialogResult.OK:
            self.txt_file.Text = dlg.FileName

    def on_import(self, sender, args):
        csv_path = self.txt_file.Text.strip()
        type_name = self.txt_type.Text.strip().upper()

        if not csv_path:
            TaskDialog.Show("Missing File", "Please select a CSV file.")
            return

        if not os.path.exists(csv_path):
            TaskDialog.Show("File Not Found", csv_path)
            return

        if not type_name:
            TaskDialog.Show("Missing Type", "Please enter a control point type.")
            return

        self.csv_path = csv_path
        self.type_name = type_name
        self.coordinate_system = "INTERNAL" if self.rb_internal.Checked else "SHARED"
        self.result_ok = True
        self.DialogResult = DialogResult.OK
        self.Close()

        if not os.path.exists(csv_path):
            TaskDialog.Show("File Not Found", csv_path)
            return

        if not type_name:
            TaskDialog.Show("Missing Type", "Please enter a control point type.")
            return

        self.csv_path = csv_path
        self.type_name = type_name
        self.coordinate_system = "INTERNAL" if self.rb_internal.Checked else "SHARED"
        self.result_ok = True
        self.Close()

    def on_cancel(self, sender, args):
        self.result_ok = False
        self.DialogResult = DialogResult.Cancel
        self.Close()

    def on_make_template(self, sender, args):
        try:
            template_path = self.txt_file.Text.strip()

            if not template_path:
                MessageBox.Show(
                    self,
                    "Enter or browse to a CSV path first.",
                    "Missing File",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning
                )
                return

            folder = os.path.dirname(template_path)
            if folder and not os.path.exists(folder):
                os.makedirs(folder)

            with open(template_path, "wb") as f:
                writer = csv.writer(f)
                writer.writerow(["POINT NUMBER", "Y", "X", "Z", "DESCRIPTION"])

            MessageBox.Show(
                self,
                "Template created at:\n{}".format(template_path),
                "Template Created",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information
            )

        except Exception as ex:
            MessageBox.Show(
                self,
                str(ex),
                "Template Failed",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error
            )

# ------------------------------------------------------------
# MAIN
# ------------------------------------------------------------
def main():
    settings = load_settings()

    form = ImportControlPointsForm(settings)
    form.ShowDialog()

    if not form.result_ok:
        return

    csv_path = form.csv_path
    type_name = form.type_name
    coordinate_system = form.coordinate_system

    save_settings({
        "csv_path": csv_path,
        "type_name": type_name,
        "coordinate_system": coordinate_system
    })

    try:
        csv_rows, skipped_rows = read_csv_rows(csv_path)
    except Exception as ex:
        TaskDialog.Show("CSV Read Failed", str(ex))
        return

    if not csv_rows:
        msg = "No valid point rows found in CSV."
        if skipped_rows:
            msg += "\n\nFirst issues:\n" + "\n".join(skipped_rows[:10])
        TaskDialog.Show("Import Stopped", msg)
        return

    if not ensure_family_loaded():
        return

    symbol = get_or_create_symbol(type_name)
    if not symbol:
        TaskDialog.Show("Type Error", "Could not find or create type '{}'.".format(type_name))
        return

    basis = None
    if coordinate_system == "SHARED":
        try:
            basis = get_shared_basis()
        except Exception as ex:
            TaskDialog.Show("Coordinate Error", "Could not evaluate shared coordinates.\n\n{}".format(str(ex)))
            return

    placed_count = 0
    failed_count = 0
    number_param_fail = 0
    desc_param_fail = 0
    failed_details = []

    t = Transaction(doc, "Import Control Points from CSV")
    started = False
    try:
        t.Start()
        started = True

        if not symbol.IsActive:
            symbol.Activate()
            doc.Regenerate()

        for row in csv_rows:
            try:
                xyz = csv_row_to_xyz(row, coordinate_system, basis)
                inst = place_instance(symbol, xyz)

                if not set_param_string(inst, "TS_Point_Number", row["point_number"]):
                    number_param_fail += 1

                if not set_param_string(inst, "TS_Point_Description", row["description"]):
                    desc_param_fail += 1

                placed_count += 1

            except Exception as row_ex:
                failed_count += 1
                if len(failed_details) < 10:
                    failed_details.append("Row {} failed: {}".format(row["line_no"], str(row_ex)))

        t.Commit()

    except Exception as ex:
        if started:
            try:
                t.RollBack()
            except:
                pass
        TaskDialog.Show("Import Failed", str(ex))
        return

    msg = []
    msg.append("Import complete.")
    msg.append("")
    msg.append("Placed: {}".format(placed_count))
    msg.append("Failed: {}".format(failed_count))
    msg.append("CSV rows skipped before placement: {}".format(len(skipped_rows)))
    msg.append("")
    msg.append("Coordinate system used: {}".format(coordinate_system))
    msg.append("Family: {}".format(FAMILY_NAME))
    msg.append("Type: {}".format(type_name))
    msg.append("File: {}".format(csv_path))

    if number_param_fail > 0 or desc_param_fail > 0:
        msg.append("")
        msg.append("Parameter warnings:")
        msg.append("TS_Point_Number not written on {} instance(s).".format(number_param_fail))
        msg.append("TS_Point_Description not written on {} instance(s).".format(desc_param_fail))

    if skipped_rows:
        msg.append("")
        msg.append("First skipped CSV issues:")
        msg.extend(skipped_rows[:10])

    if failed_details:
        msg.append("")
        msg.append("First placement failures:")
        msg.extend(failed_details[:10])

    TaskDialog.Show("CSV Import Result", "\n".join(msg))


main()