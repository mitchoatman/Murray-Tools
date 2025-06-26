import clr
clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')
import System
from System.Windows.Forms import Form, Label, TextBox, Button, DialogResult, FormStartPosition, FormBorderStyle, MessageBox, ListBox
from System.Drawing import Point, Size
from Autodesk.Revit.DB import Transaction, FilteredElementCollector, BuiltInCategory, FabricationPart
from Parameters.Get_Set_Params import get_parameter_value_by_name_AsString, set_parameter_by_name
import re

uidoc = __revit__.ActiveUIDocument
doc = uidoc.Document

def natural_key(s):
    return [int(text) if text.isdigit() else text.lower() for text in re.split('([0-9]+)', s)]

def get_fabrication_parts(view_id):
    categories = [BuiltInCategory.OST_FabricationPipework,
                  BuiltInCategory.OST_FabricationDuctwork,
                  BuiltInCategory.OST_FabricationContainment]
    parts = []
    for cat in categories:
        collector = FilteredElementCollector(doc, view_id).OfCategory(cat).OfClass(FabricationPart)
        parts.extend(list(collector))
    return parts

def get_assemblies():
    assemblies = set()
    for elem in get_fabrication_parts(doc.ActiveView.Id):
        param_asm = elem.LookupParameter("STRATUS Assembly")
        param_pkg = elem.LookupParameter("STRATUS Package")
        if param_asm and param_asm.HasValue:
            asm_val = param_asm.AsString()
            if asm_val:
                assemblies.add("Assembly: {}".format(asm_val))
        if param_pkg and param_pkg.HasValue:
            pkg_val = param_pkg.AsString()
            if pkg_val:
                assemblies.add("Package: {}".format(pkg_val))
    return sorted(assemblies, key=natural_key)

class RenameForm(Form):
    def __init__(self, assemblies):
        self.Text = "Rename STRATUS Assembly or Package"
        self.scale_factor = self.get_dpi_scale()
        self.all_assemblies = assemblies
        self.InitializeComponents(assemblies)

    def get_dpi_scale(self):
        screen = System.Windows.Forms.Screen.PrimaryScreen
        graphics = self.CreateGraphics()
        dpi_x = graphics.DpiX
        graphics.Dispose()
        return dpi_x / 96.0

    def scale_value(self, value):
        return int(value * self.scale_factor)

    def update_assemblies_list(self, filter_text=""):
        self.listbox.Items.Clear()
        filtered = [a for a in self.all_assemblies if filter_text in a]
        for item in filtered:
            self.listbox.Items.Add(item)

    def on_search_text_changed(self, sender, event):
        self.update_assemblies_list(self.find_textbox.Text)

    def on_select_all_click(self, sender, event):
        for i in range(self.listbox.Items.Count):
            self.listbox.SetSelected(i, True)

    def on_listbox_double_click(self, sender, event):
        if self.listbox.SelectedItem:
            item = self.listbox.SelectedItem
            if item.startswith("Assembly: "):
                self.find_textbox.Text = item.replace("Assembly: ", "")
            elif item.startswith("Package: "):
                self.find_textbox.Text = item.replace("Package: ", "")


    def on_rename_click(self, sender, event):
        find = self.find_textbox.Text
        replace = self.replace_textbox2.Text
        selected = [self.listbox.Items[i] for i in self.listbox.SelectedIndices]

        if not find or not replace:
            MessageBox.Show("Please enter both Find and Replace values.", "Warning")
            return
        if not selected:
            MessageBox.Show("No assemblies selected.", "Warning")
            return

        elements_to_process = []
        for elem in get_fabrication_parts(doc.ActiveView.Id):
            asm_val = get_parameter_value_by_name_AsString(elem, "STRATUS Assembly")
            pkg_val = get_parameter_value_by_name_AsString(elem, "STRATUS Package")
            for sel in selected:
                if sel.startswith("Assembly: ") and sel[10:] == asm_val:
                    elements_to_process.append((elem, "STRATUS Assembly"))
                elif sel.startswith("Package: ") and sel[9:] == pkg_val:
                    elements_to_process.append((elem, "STRATUS Package"))

        if not elements_to_process:
            MessageBox.Show("No fabrication parts matched your selection.", "Warning")
            return

        t = Transaction(doc, "Rename STRATUS Parameters")
        t.Start()
        try:
            for elem, pname in elements_to_process:
                current = get_parameter_value_by_name_AsString(elem, pname)
                if current and not elem.LookupParameter(pname).IsReadOnly:
                    new_val = current.replace(find, replace)
                    set_parameter_by_name(elem, pname, new_val)
            t.Commit()
        except Exception as e:
            MessageBox.Show("Error during transaction:\n{}".format(str(e)), "Error")
            if t.HasStarted():
                t.RollBack()
            return

        # Refresh assemblies with updated names and keep prefix
        self.all_assemblies = get_assemblies()
        self.update_assemblies_list(self.find_textbox.Text)

    def add_textbox_with_clear(self, y, handler=None):
        tb = TextBox()
        tb.Location = Point(self.scale_value(20), y)
        tb.Size = Size(self.scale_value(210), self.scale_value(20))
        if handler:
            tb.TextChanged += handler
        self.Controls.Add(tb)

        clear_btn = Button()
        clear_btn.Text = "X"
        clear_btn.Size = Size(self.scale_value(20), self.scale_value(20))
        clear_btn.Location = Point(self.scale_value(235), y)
        clear_btn.Click += lambda sender, args: tb.Clear()
        self.Controls.Add(clear_btn)

        return tb, y + self.scale_value(25)

    def add_label(self, text, y):
        lbl = Label()
        lbl.Text = text
        lbl.Location = Point(self.scale_value(20), y)
        lbl.Size = Size(self.scale_value(260), self.scale_value(20))
        self.Controls.Add(lbl)
        return y + self.scale_value(21)

    def add_textbox(self, y):
        tb = TextBox()
        tb.Location = Point(self.scale_value(20), y)
        tb.Size = Size(self.scale_value(240), self.scale_value(20))
        self.Controls.Add(tb)
        return tb, y + self.scale_value(25)

    def InitializeComponents(self, assemblies):
        self.FormBorderStyle = FormBorderStyle.FixedSingle
        self.MaximizeBox = False
        self.MinimizeBox = False
        self.StartPosition = FormStartPosition.CenterScreen

        self.Width = self.scale_value(300)
        item_height = self.scale_value(20)
        listbox_height = int(item_height * min(15, max(7, len(assemblies)))) + self.scale_value(10)
        header_height = self.scale_value(160)
        button_height = self.scale_value(30)
        bottom_margin = self.scale_value(20)
        self.Height = listbox_height + header_height + button_height + bottom_margin

        y = self.scale_value(10)
        y = self.add_label("Find (Case Sensitive):", y)
        self.find_textbox, y = self.add_textbox_with_clear(y, self.on_search_text_changed)

        y = self.add_label("Replace:", y)
        self.replace_textbox2, y = self.add_textbox(y)

        y = self.add_label("Select STRATUS Assemblies to Rename:", y)
        self.listbox = ListBox()
        self.listbox.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended
        self.listbox.Location = Point(self.scale_value(20), y)
        self.listbox.Size = Size(self.scale_value(240), listbox_height)
        self.listbox.DoubleClick += self.on_listbox_double_click
        for a in assemblies:
            self.listbox.Items.Add(a)
        self.Controls.Add(self.listbox)
        y += listbox_height + self.scale_value(10)

        btn_y = y
        btn_w = self.scale_value(75)

        self.select_all_button = Button()
        self.select_all_button.Text = "Select All"
        self.select_all_button.Size = Size(btn_w, button_height)
        self.select_all_button.Location = Point(self.scale_value(20), btn_y)
        self.select_all_button.Click += self.on_select_all_click
        self.Controls.Add(self.select_all_button)

        self.rename_button = Button()
        self.rename_button.Text = "Rename"
        self.rename_button.Size = Size(btn_w, button_height)
        self.rename_button.Location = Point(self.scale_value(105), btn_y)
        self.rename_button.Click += self.on_rename_click
        self.Controls.Add(self.rename_button)

        self.cancel_button = Button()
        self.cancel_button.Text = "Close"
        self.cancel_button.Size = Size(btn_w, button_height)
        self.cancel_button.Location = Point(self.scale_value(190), btn_y)
        self.cancel_button.DialogResult = DialogResult.Cancel
        self.Controls.Add(self.cancel_button)

        self.CancelButton = self.cancel_button

# Run form
form = RenameForm(get_assemblies())
form.ShowDialog()
