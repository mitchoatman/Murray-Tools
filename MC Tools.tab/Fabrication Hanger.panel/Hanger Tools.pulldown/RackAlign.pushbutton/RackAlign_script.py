import clr
clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
clr.AddReference("PresentationCore")
clr.AddReference("PresentationFramework")
clr.AddReference("WindowsBase")

from Autodesk.Revit.DB import Transaction, BuiltInParameter
from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter
import re
from System.Windows import Window, Thickness, WindowStartupLocation, ResizeMode, HorizontalAlignment
from System.Windows.Controls import StackPanel, Label, TextBox, CheckBox, Button, Orientation
from System.Windows.Media import FontFamily

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
app = doc.Application
RevitVersion = app.VersionNumber
RevitINT = int(RevitVersion)

def parse_elevation(input_str):
    input_str = input_str.strip().replace('"', '')
    if re.match(r"^\d+(\.\d+)?$", input_str):
        return float(input_str)
    match = re.match(r"(\d+)['\s\-]*(?:(?:(\d+)\s+(\d+/\d+)|(\d*(?:\.\d+)?))|\d*)?", input_str)
    if match:
        feet = float(match.group(1) or 0)
        inches = 0
        if match.group(2) and match.group(3):
            whole_inches = float(match.group(2))
            fraction = match.group(3)
            inches = whole_inches + eval(fraction)
        elif match.group(4):
            inches = float(match.group(4))
        elif match.group(2):
            inches = float(match.group(2))
        return feet + (inches / 12)
    raise ValueError("Invalid elevation format.")

class FabricationPartFilter(ISelectionFilter):
    def AllowElement(self, element):
        return element.LookupParameter("Fabrication Service") is not None
    def AllowReference(self, reference, point):
        return False

# Selection prompt
try:
    selected_refs = uidoc.Selection.PickObjects(
        ObjectType.Element, 
        FabricationPartFilter(),
        "Select Fabrication Rack Parts to align"
    )
    selected_elements = [doc.GetElement(ref.ElementId) for ref in selected_refs]
except:
    selected_elements = []

if selected_elements:
    # Get initial elevation estimate
    try:
        if RevitINT > 2022:
            ElevationEstimate = selected_elements[0].LookupParameter('Lower End Bottom Elevation').AsValueString() or "0"
        else:
            ElevationEstimate = selected_elements[0].LookupParameter('Bottom').AsValueString() or "0"
    except:
        ElevationEstimate = "0"

    class AlignmentDialog(Window):
        def __init__(self, default_elev):
            super(AlignmentDialog, self).__init__()
            self.Title = "Rack Alignment"
            self.Width = 380
            self.Height = 260                    # Reduced height
            self.WindowStartupLocation = WindowStartupLocation.CenterScreen
            self.ResizeMode = ResizeMode.NoResize

            stack = StackPanel()
            stack.Orientation = Orientation.Vertical
            stack.Margin = Thickness(15, 15, 15, 8)   # Reduced bottom margin

            def create_label(text):
                lbl = Label()
                lbl.Content = text
                lbl.FontSize = 13
                lbl.FontFamily = FontFamily("Segoe UI")
                return lbl

            # Checkboxes
            self.chk_top = CheckBox()
            self.chk_top.Content = "Align TOP"
            self.chk_top.IsChecked = False
            self.chk_top.Margin = Thickness(0, 5, 0, 5)
            stack.Children.Add(self.chk_top)

            self.chk_bottom = CheckBox()
            self.chk_bottom.Content = "Align BOTTOM"
            self.chk_bottom.IsChecked = True
            self.chk_bottom.Margin = Thickness(0, 5, 0, 5)
            stack.Children.Add(self.chk_bottom)

            self.chk_ignore_ins = CheckBox()
            self.chk_ignore_ins.Content = "Ignore Insulation"
            self.chk_ignore_ins.IsChecked = True
            self.chk_ignore_ins.Margin = Thickness(0, 5, 0, 10)
            stack.Children.Add(self.chk_ignore_ins)

            stack.Children.Add(create_label("Reference Bottom Elevation:"))

            self.txt_elev = TextBox()
            self.txt_elev.Text = str(default_elev) if default_elev else "0"
            self.txt_elev.Width = 220
            self.txt_elev.HorizontalAlignment = HorizontalAlignment.Left
            self.txt_elev.Margin = Thickness(0, 5, 0, 15)
            stack.Children.Add(self.txt_elev)

            btn_ok = Button()
            btn_ok.Content = "OK"
            btn_ok.Width = 80
            btn_ok.Height = 30
            btn_ok.HorizontalAlignment = HorizontalAlignment.Center
            btn_ok.Margin = Thickness(0, 12, 0, 8)   # Reduced bottom spacing
            btn_ok.Click += self.on_ok_clicked
            stack.Children.Add(btn_ok)

            self.Content = stack

        def on_ok_clicked(self, sender, event):
            self.DialogResult = True
            self.Close()

    # Show WPF dialog
    form = AlignmentDialog(ElevationEstimate)
    if form.ShowDialog():
        try:
            PRTElevation = parse_elevation(form.txt_elev.Text)
            TOP = form.chk_top.IsChecked
            BTM = form.chk_bottom.IsChecked
            INS = form.chk_ignore_ins.IsChecked

            t = Transaction(doc, "Rack Align")
            t.Start()

            for elem in selected_elements:
                isfabpart = elem.LookupParameter("Fabrication Service")
                if isfabpart and hasattr(elem, 'ItemCustomId') and elem.ItemCustomId in [2041, 866, 40]:
                    if RevitINT > 2022:
                        if BTM:
                            elem.get_Parameter(BuiltInParameter.MEP_LOWER_BOTTOM_ELEVATION).Set(PRTElevation)
                        if INS and getattr(elem, 'InsulationThickness', 0) > 0:
                            INSthickness = elem.LookupParameter('Insulation Thickness').AsDouble()
                            elem.get_Parameter(BuiltInParameter.MEP_LOWER_BOTTOM_ELEVATION).Set(PRTElevation - INSthickness)
                        if TOP:
                            elem.get_Parameter(BuiltInParameter.MEP_LOWER_TOP_ELEVATION).Set(PRTElevation)
                        if INS and getattr(elem, 'InsulationThickness', 0) > 0:
                            INSthickness = elem.LookupParameter('Insulation Thickness').AsDouble()
                            elem.get_Parameter(BuiltInParameter.MEP_LOWER_TOP_ELEVATION).Set(PRTElevation + INSthickness)
                    else:
                        if BTM:
                            elem.get_Parameter(BuiltInParameter.FABRICATION_BOTTOM_OF_PART).Set(PRTElevation)
                        if INS and getattr(elem, 'InsulationThickness', 0) > 0:
                            INSthickness = elem.LookupParameter('Insulation Thickness').AsDouble()
                            elem.get_Parameter(BuiltInParameter.FABRICATION_BOTTOM_OF_PART).Set(PRTElevation - INSthickness)
                        if TOP:
                            elem.get_Parameter(BuiltInParameter.FABRICATION_TOP_OF_PART).Set(PRTElevation)
                        if INS and getattr(elem, 'InsulationThickness', 0) > 0:
                            INSthickness = elem.LookupParameter('Insulation Thickness').AsDouble()
                            elem.get_Parameter(BuiltInParameter.FABRICATION_TOP_OF_PART).Set(PRTElevation + INSthickness)

            t.Commit()
        except:
            pass