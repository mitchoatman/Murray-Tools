import Autodesk
from Autodesk.Revit.DB import Transaction, BuiltInParameter, FamilyInstance, FamilySymbol, XYZ, ElementTransformUtils, BoundingBoxXYZ, Line, FilteredElementCollector, BuiltInCategory
from Autodesk.Revit.UI.Selection import ObjectType
import math
import os
import re
from fractions import Fraction
import clr
clr.AddReference("System.Windows.Forms")
clr.AddReference("System.Drawing")
clr.AddReference("System")

from System.Windows.Forms import *
from System.Drawing import Point, Size, Font
from System import Array

# Define some variables for easy use
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
curview = doc.ActiveView
app = doc.Application
RevitVersion = app.VersionNumber
RevitINT = float(RevitVersion)

# Function to convert fraction string to float
def frac2string(s):
    i, f = s.groups(0)
    f = Fraction(f)
    return str(int(i) + float(f))

try:
    selected_element = uidoc.Selection.PickObject(ObjectType.Element, 'Select OUTSIDE Pipe')
    if doc.GetElement(selected_element.ElementId).ItemCustomId != 916:
        selected_element1 = uidoc.Selection.PickObject(ObjectType.Element, 'Select OPPOSITE OUTSIDE Pipe')
        element = doc.GetElement(selected_element.ElementId)
        element1 = doc.GetElement(selected_element1.ElementId)
        selected_elements = [element, element1]

        level_id = element.LevelId
        level = doc.GetElement(level_id)
        level_elevation = level.Elevation if level else 0

        # Function to get parameter value
        def get_parameter_value(element, parameterName):
            param = element.LookupParameter(parameterName)
            if param and param.HasValue:
                return param.AsDouble()
            return None

        # Get bottom elevation of selected pipe
        PRTElevation = None
        if element and RevitINT > 2022:
            PRTElevation = get_parameter_value(element, 'Lower End Bottom Elevation')
        if element and RevitINT < 2023:
            PRTElevation = get_parameter_value(element, 'Bottom')

        # Fallback for PRTElevation using curve or connectors
        if PRTElevation is None:
            curve = element.get_Curve()
            if curve:
                PRTElevation = min(curve.GetEndPoint(0).Z, curve.GetEndPoint(1).Z)
            else:
                connectors = element.ConnectorManager.Connectors
                connector_list = list(connectors)
                if connector_list:
                    PRTElevation = min([conn.Origin.Z for conn in connector_list])
                else:
                    raise Exception("Cannot determine pipe elevation.")

        # Validate PRTElevation
        if PRTElevation < level_elevation - 1.0:
            bbox = element.get_BoundingBox(None)
            if bbox:
                PRTElevation = bbox.Min.Z

        # Get Outside Diameter of the first selected pipe
        outside_diameter = None
        od_param = element.LookupParameter("Overall Size")
        if od_param and od_param.HasValue:
            outside_diameter = od_param.AsString()
        else:
            try:
                outside_diameter = element.Diameter
            except:
                outside_diameter = None

        # Calculate the angle of the first selected pipe
        pipe_angle = 0.0
        try:
            curve = element.get_Curve()
            if curve:
                direction = curve.Direction
                pipe_angle = math.atan2(direction.Y, direction.X)
        except:
            try:
                connectors = element.ConnectorManager.Connectors
                connector_list = list(connectors)
                if len(connector_list) >= 2:
                    start_connector = connector_list[0]
                    end_connector = connector_list[1]
                    vector = end_connector.Origin - start_connector.Origin
                    pipe_angle = math.atan2(vector.Y, vector.X)
            except:
                pass

        # Collect all pipe accessory families
        pipe_accessory_symbols = FilteredElementCollector(doc).OfClass(FamilySymbol).OfCategory(BuiltInCategory.OST_PipeAccessory).ToElements()
        family_names = []
        family_symbol_dict = {}

        for symbol in pipe_accessory_symbols:
            family_name = symbol.Family.Name
            type_name = symbol.LookupParameter("Type Name").AsString() if symbol.LookupParameter("Type Name") else symbol.Name
            display_name = family_name + " - " + type_name
            if display_name not in family_names:
                family_names.append(display_name)
                family_symbol_dict[display_name] = symbol

        family_names.sort()

        folder_name = "c:\\Temp"
        filepath = os.path.join(folder_name, 'Ribbon_PlaceAccessory.txt')
        if not os.path.exists(folder_name):
            os.makedirs(folder_name)
        if not os.path.exists(filepath):
            with open(filepath, 'w') as the_file:
                line1 = (family_names[0] + '\n') if family_names else ('Unknown Family - Unknown Type' + '\n')
                line2 = ('1.0' + '\n')
                line3 = ('8.0' + '\n')
                line4 = ('True' + '\n')
                line5 = ('True' + '\n')
                the_file.writelines([line1, line2, line3, line4, line5])

        with open(filepath, 'r') as file:
            lines = file.readlines()
            lines = [line.rstrip() for line in lines]

        if len(lines) < 5:
            with open(filepath, 'w') as the_file:
                line1 = (family_names[0] + '\n') if family_names else ('Unknown Family - Unknown Type' + '\n')
                line2 = ('1.0' + '\n')
                line3 = ('8.0' + '\n')
                line4 = ('True' + '\n')
                line5 = ('True' + '\n')
                the_file.writelines([line1, line2, line3, line4, line5])

        with open(filepath, 'r') as file:
            lines = file.readlines()
            lines = [line.rstrip() for line in lines]

        if lines[3] == 'False':
            checkboxdefBOI = False
        else:
            checkboxdefBOI = True

        if len(lines) > 4 and lines[4] == 'False':
            checkboxdefRotate = False
        else:
            checkboxdefRotate = True

        class HangerSpacingDialog(Form):
            def __init__(self, family_names, lines, checkboxdefBOI, checkboxdefRotate):
                self.Text = "Hanger and Spacing"
                self.Size = Size(350, 310)
                self.StartPosition = FormStartPosition.CenterScreen
                self.FormBorderStyle = FormBorderStyle.FixedDialog

                label_hanger = Label()
                label_hanger.Text = "Choose Hanger Family:"
                label_hanger.Location = Point(10, 10)
                label_hanger.Size = Size(300, 20)
                label_hanger.Font = Font("Arial", 10)
                self.Controls.Add(label_hanger)

                self.combobox_hanger = ComboBox()
                self.combobox_hanger.Location = Point(10, 31)
                self.combobox_hanger.Size = Size(300, 20)
                self.combobox_hanger.DropDownStyle = ComboBoxStyle.DropDownList
                self.combobox_hanger.Items.AddRange(Array[object](family_names))
                if lines[0] in family_names:
                    self.combobox_hanger.SelectedItem = lines[0]
                else:
                    self.combobox_hanger.SelectedItem = family_names[0] if family_names else None
                self.Controls.Add(self.combobox_hanger)

                label_end_dist = Label()
                label_end_dist.Text = "Distance from End (Ft):"
                label_end_dist.Font = Font("Arial", 10)
                label_end_dist.Location = Point(10, 60)
                label_end_dist.Size = Size(300, 20)
                self.Controls.Add(label_end_dist)

                self.textbox_end_dist = TextBox()
                self.textbox_end_dist.Location = Point(10, 80)
                self.textbox_end_dist.Text = lines[1]
                self.Controls.Add(self.textbox_end_dist)

                label_spacing = Label()
                label_spacing.Text = "Hanger Spacing (Ft):"
                label_spacing.Font = Font("Arial", 10)
                label_spacing.Location = Point(10, 110)
                label_spacing.Size = Size(300, 20)
                self.Controls.Add(label_spacing)

                self.textbox_spacing = TextBox()
                self.textbox_spacing.Location = Point(10, 130)
                self.textbox_spacing.Text = lines[2]
                self.Controls.Add(self.textbox_spacing)

                self.checkbox_boi = CheckBox()
                self.checkbox_boi.Text = "Align Trapeze to Bottom of Insulation"
                self.checkbox_boi.Font = Font("Arial", 10)
                self.checkbox_boi.Location = Point(10, 160)
                self.checkbox_boi.Size = Size(300, 20)
                self.checkbox_boi.Checked = checkboxdefBOI
                self.Controls.Add(self.checkbox_boi)

                self.checkbox_rotate = CheckBox()
                self.checkbox_rotate.Text = "Rotate Family"
                self.checkbox_rotate.Font = Font("Arial", 10)
                self.checkbox_rotate.Location = Point(10, 190)
                self.checkbox_rotate.Size = Size(300, 20)
                self.checkbox_rotate.Checked = checkboxdefRotate
                self.Controls.Add(self.checkbox_rotate)

                self.button_ok = Button()
                self.button_ok.Text = "OK"
                self.button_ok.Font = Font("Arial", 10)
                self.button_ok.Location = Point(((self.Width / 2) - 50), 230)
                self.button_ok.Click += self.ok_button_clicked
                self.Controls.Add(self.button_ok)

            def ok_button_clicked(self, sender, event):
                self.DialogResult = DialogResult.OK
                self.Close()

        form = HangerSpacingDialog(family_names, lines, checkboxdefBOI, checkboxdefRotate)
        if form.ShowDialog() == DialogResult.OK:
            SelectedFamily = form.combobox_hanger.Text
            distancefromend = form.textbox_end_dist.Text
            Spacing = form.textbox_spacing.Text
            BOITrap = form.checkbox_boi.Checked
            RotateFamily = form.checkbox_rotate.Checked

            selected_symbol = family_symbol_dict.get(SelectedFamily)
            if not selected_symbol:
                raise Exception("Selected family '" + SelectedFamily + "' not found.")

            if not selected_symbol.IsActive:
                t = Transaction(doc, 'Activate Family Symbol')
                t.Start()
                selected_symbol.Activate()
                t.Commit()

            with open(filepath, 'w') as the_file:
                line1 = (str(SelectedFamily) + '\n')
                line2 = (str(distancefromend) + '\n')
                line3 = (str(Spacing) + '\n')
                line4 = (str(BOITrap) + '\n')
                line5 = (str(RotateFamily) + '\n')
                the_file.writelines([line1, line2, line3, line4, line5])

            def GetCenterPoint(ele):
                bBox = doc.GetElement(ele).get_BoundingBox(None)
                if bBox is None:
                    return XYZ(0, 0, 0)
                center = (bBox.Max + bBox.Min) / 2
                return center

            def myround(x, multiple):
                return multiple * math.ceil(x/multiple)

            first_pipe_bounding_box = element.get_BoundingBox(curview)

            # Determine Rack Direction
            delta_x = abs(first_pipe_bounding_box.Max.X - first_pipe_bounding_box.Min.X)
            delta_y = abs(first_pipe_bounding_box.Max.Y - first_pipe_bounding_box.Min.Y)

            # Calculate combined bounding box with insulation
            combined_min = first_pipe_bounding_box.Min
            combined_max = first_pipe_bounding_box.Max

            if (delta_x) > (delta_y):
                for pipe in selected_elements:
                    pipe_bounding_box = pipe.get_BoundingBox(curview)
                    if hasattr(pipe, 'InsulationThickness') and pipe.InsulationThickness > 0:
                        pipe_bounding_box.Min = XYZ(pipe_bounding_box.Min.X,
                                                    pipe_bounding_box.Min.Y - pipe.InsulationThickness,
                                                    pipe_bounding_box.Min.Z)
                        pipe_bounding_box.Max = XYZ(pipe_bounding_box.Max.X,
                                                    pipe_bounding_box.Max.Y + pipe.InsulationThickness,
                                                    pipe_bounding_box.Max.Z)
                    combined_min = XYZ(min(combined_min.X, pipe_bounding_box.Min.X),
                                       min(combined_min.Y, pipe_bounding_box.Min.Y),
                                       min(combined_min.Z, pipe_bounding_box.Min.Z))
                    combined_max = XYZ(max(combined_max.X, pipe_bounding_box.Max.X),
                                       max(combined_max.Y, pipe_bounding_box.Max.Y),
                                       max(combined_max.Z, pipe_bounding_box.Max.Z))

            if (delta_y) > (delta_x):
                for pipe in selected_elements:
                    pipe_bounding_box = pipe.get_BoundingBox(curview)
                    if hasattr(pipe, 'InsulationThickness') and pipe.InsulationThickness > 0:
                        pipe_bounding_box.Min = XYZ(pipe_bounding_box.Min.X - pipe.InsulationThickness,
                                                    pipe_bounding_box.Min.Y,
                                                    pipe_bounding_box.Min.Z)
                        pipe_bounding_box.Max = XYZ(pipe_bounding_box.Max.X + pipe.InsulationThickness,
                                                    pipe_bounding_box.Max.Y,
                                                    pipe_bounding_box.Max.Z)
                    combined_min = XYZ(min(combined_min.X, pipe_bounding_box.Min.X),
                                       min(combined_min.Y, pipe_bounding_box.Min.Y),
                                       min(combined_min.Z, pipe_bounding_box.Min.Z))
                    combined_max = XYZ(max(combined_max.X, pipe_bounding_box.Max.X),
                                       max(combined_max.Y, pipe_bounding_box.Max.Y),
                                       max(combined_max.Z, pipe_bounding_box.Max.Z))

            def get_reference_level(hanger):
                level_id = hanger.LevelId
                level = doc.GetElement(level_id)
                return level

            def get_level_elevation(level):
                if level:
                    return level.Elevation
                else:
                    return 0

            combined_bounding_box = BoundingBoxXYZ()
            combined_bounding_box.Min = combined_min
            combined_bounding_box.Max = combined_max
            combined_bounding_box_Center = (combined_bounding_box.Max + combined_bounding_box.Min) / 2

            X_side_xyz = XYZ(combined_bounding_box.Min.X + float(distancefromend), 
                             combined_bounding_box_Center.Y, 
                             PRTElevation)
            Y_side_xyz = XYZ(combined_bounding_box_Center.X, 
                             combined_bounding_box.Min.Y + float(distancefromend), 
                             PRTElevation)

            delta_x = abs(combined_bounding_box.Max.X - combined_bounding_box.Min.X)
            delta_y = abs(combined_bounding_box.Max.Y - combined_bounding_box.Min.Y)

            # Calculate how many hangers in the run
            if (delta_x) > (delta_y):
                qtyofhgrs = int(math.ceil(delta_x / float(Spacing)))
            if (delta_y) > (delta_x):
                qtyofhgrs = int(math.ceil(delta_y / float(Spacing)))
                
            IncrementSpacing = float(distancefromend)

            # Place family instances
            hangers = []
            t = Transaction(doc, 'Place Pipe Accessory Hanger')
            t.Start()
            
            for hgr in range(qtyofhgrs):
                if (delta_x) > (delta_y):
                    location = XYZ(combined_bounding_box.Min.X + IncrementSpacing, 
                                   combined_bounding_box_Center.Y, 
                                   PRTElevation)
                else:
                    location = XYZ(combined_bounding_box_Center.X, 
                                   combined_bounding_box.Min.Y + IncrementSpacing, 
                                   PRTElevation)
                
                hanger = doc.Create.NewFamilyInstance(location, selected_symbol, doc.GetElement(level_id), Autodesk.Revit.DB.Structure.StructuralType.NonStructural)
                hangers.append(hanger)
                IncrementSpacing += float(Spacing)
            
            for hanger in hangers:
                reference_level = get_reference_level(hanger)
                level_elevation = get_level_elevation(reference_level)

                # Set DIM C parameter for width
                if (delta_x) > (delta_y):
                    base_width = myround((delta_y * 12), 2) / 12
                else:
                    base_width = myround((delta_x * 12), 2) / 12

                is_figure_109 = SelectedFamily == "CEAS Stiffy Figure 109 - 109"
                newwidth = base_width if is_figure_109 else base_width + (4.0 / 12.0)

                dim_c_param = hanger.LookupParameter("DIM C")
                if dim_c_param:
                    dim_c_param.Set(newwidth)

                if is_figure_109:
                    dim_b_param = hanger.LookupParameter("DIM B")
                    if dim_b_param:
                        dim_b_param.Set(newwidth)

                z_axis_direction = XYZ(0, 0, 1)
                center = GetCenterPoint(hanger.Id)
                curve_points = [center, center + z_axis_direction * 2]
                curve = Line.CreateBound(curve_points[0], curve_points[1])
                ElementTransformUtils.RotateElement(doc, hanger.Id, curve, pipe_angle)

                if RotateFamily:
                    ElementTransformUtils.RotateElement(doc, hanger.Id, curve, (90.0 * (math.pi / 180.0)))

                if outside_diameter is not None:
                    tier_1_param = hanger.LookupParameter("Tier_1 Restraint")
                    if tier_1_param:
                        if '/' in outside_diameter:
                            tier_1_param.Set(float(re.sub(r'(?:(\d+)[-\s])?(\d+/\d+)[^\d.]', frac2string, outside_diameter)) / 12 + 0.02083333)
                        else:
                            tier_1_param.Set(float(re.sub(r'[^\d.]', '', outside_diameter)) / 12 + 0.02083333)

                offset_param = hanger.get_Parameter(BuiltInParameter.INSTANCE_FREE_HOST_OFFSET_PARAM)
                if offset_param:
                    pipe_elevation = PRTElevation
                    offset = pipe_elevation - level_elevation

                    if is_figure_109:
                        dim_a_param = hanger.LookupParameter("DIM A")
                        if dim_a_param and dim_a_param.HasValue:
                            dim_a_value = dim_a_param.AsDouble()
                            offset += dim_a_value - (0.5 / 12.0)
                        else:
                            offset = pipe_elevation - level_elevation

                    if not is_figure_109 and BOITrap:
                        offset = pipe_elevation - level_elevation

                    if offset < 0:
                        bbox = element.get_BoundingBox(None)
                        if bbox:
                            pipe_elevation = bbox.Min.Z
                            offset = pipe_elevation - level_elevation
                        else:
                            raise Exception("Cannot determine valid pipe elevation for offset.")

                    offset_param.Set(offset)

            t.Commit()
except Exception as e:
    print "Error: " + str(e)