import Autodesk
from Autodesk.Revit.DB import Transaction, FabricationConfiguration, BuiltInParameter, FabricationPart, FabricationServiceButton, FabricationService, XYZ, ConnectorType
from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType
import math
import os
import clr
clr.AddReference("PresentationCore")
clr.AddReference("PresentationFramework")
clr.AddReference("WindowsBase")
from System.Windows import Window, Thickness, WindowStartupLocation, ResizeMode, HorizontalAlignment
from System.Windows.Controls import StackPanel, Label, ComboBox, TextBox, CheckBox, Button, Orientation
from System.Windows.Media import FontFamily
from System import Array

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
curview = doc.ActiveView
app = doc.Application
RevitVersion = app.VersionNumber
RevitINT = float(RevitVersion)

# Custom selection filter for Fabrication Parts
class FabricationPartSelectionFilter(ISelectionFilter):
    def AllowElement(self, element):
        return isinstance(element, FabricationPart)
   
    def AllowReference(self, reference, point):
        return False

# Prompt user to select pipes (with filter)
class CustomISelectionFilter(ISelectionFilter):
    def __init__(self, nom_categorie):
        self.nom_categorie = nom_categorie
    def AllowElement(self, e):
        if e.LookupParameter('Fabrication Service').AsValueString() == self.nom_categorie:
            return True
        else:
            return False
    def AllowReference(self, ref, point):
        return True

# Function to order selected elements by connector proximity and track entry connectors
def order_selected_elements(selected_elements, start_element, start_connector):
    ordered_elements = [start_element]
    entry_connectors = {start_element.Id: start_connector}
    remaining_elements = set(e.Id for e in selected_elements) - {start_element.Id}
    current_element = start_element
    current_connector = start_connector
   
    while remaining_elements:
        min_distance = float('inf')
        next_element = None
        next_connector = None
       
        for element_id in remaining_elements:
            element = doc.GetElement(element_id)
            connector_manager = element.ConnectorManager
            if not connector_manager:
                continue
            for conn in connector_manager.Connectors:
                distance = current_connector.Origin.DistanceTo(conn.Origin)
                if distance < min_distance:
                    min_distance = distance
                    next_element = element
                    next_connector = conn
       
        if next_element:
            ordered_elements.append(next_element)
            entry_connectors[next_element.Id] = next_connector
            remaining_elements.remove(next_element.Id)
            current_element = next_element
            # Find the opposite connector on the current element
            for conn in current_element.ConnectorManager.Connectors:
                if conn != next_connector:
                    current_connector = conn
                    break
        else:
            break
   
    return ordered_elements, entry_connectors

# Function to determine if an element is vertical
def vertical_fab(element):
    connectors = element.ConnectorManager.Connectors
    connector_points = [connector.Origin for connector in connectors]
    if len(connector_points) >= 2:
        point1 = connector_points[0]
        point2 = connector_points[1]
        if abs(point1.X - point2.X) < 0.001 and abs(point1.Y - point2.Y) < 0.001:
            return True
    return False

# Function to find the nearest connector to a point
def find_nearest_connector(element, pick_point):
    connector_manager = element.ConnectorManager
    if not connector_manager:
        return None
   
    nearest_connector = None
    min_distance = float('inf')
   
    for connector in connector_manager.Connectors:
        connector_point = connector.Origin
        distance = pick_point.DistanceTo(connector_point)
        if distance < min_distance:
            min_distance = distance
            nearest_connector = connector
   
    return nearest_connector

# Function to analyze run for start, end, and angles
def analyze_run(selected_elements, start_element, start_connector):
    pipe_directions = []
    fitting_angles = []
    total_run_length = 0
    start_point = start_connector.Origin
    end_point = start_point
    segments = []
    current_position = 0
    # Use the selected_elements as the ordered run
    ordered_run = selected_elements
    # Collect segment info and calculate length
    for e in ordered_run:
        length = e.CenterlineLength
        is_fabpart = e.LookupParameter('Part Pattern Number').AsInteger() in (2041, 866, 40)
        segments.append((e, current_position, length, is_fabpart))
        total_run_length += length
        current_position += length
        if e == ordered_run[-1]:
            connectors = e.ConnectorManager.Connectors
            for conn in connectors:
                if conn.Origin.DistanceTo(start_point) > end_point.DistanceTo(start_point):
                    end_point = conn.Origin
    # Analyze directions and angles (kept for potential future use)
    for i, (e, start_pos, length, is_fabpart) in enumerate(segments):
        if is_fabpart:
            connectors = e.ConnectorManager.Connectors
            conn_points = [conn.Origin for conn in connectors]
            if len(conn_points) >= 2:
                p1, p2 = conn_points[:2]
                direction = (p2 - p1).Normalize()
                pipe_directions.append((e, direction))
                if i < len(segments) - 1:
                    next_e, _, _, next_is_fabpart = segments[i + 1]
                    if next_is_fabpart:
                        next_connectors = next_e.ConnectorManager.Connectors
                        next_points = [conn.Origin for conn in next_connectors]
                        if len(next_points) >= 2:
                            np1, np2 = next_points[:2]
                            next_direction = (np2 - np1).Normalize()
                            dot_product = direction.DotProduct(next_direction)
                            dot_product = min(1.0, max(-1.0, dot_product))
                            angle_rad = math.acos(dot_product)
                            angle_deg = math.degrees(angle_rad)
                            if i + 1 < len(segments) and not segments[i + 1][3]:
                                fitting_angles.append((segments[i + 1][0], angle_deg))
    return {
        'start_point': start_point,
        'end_point': end_point,
        'total_length': total_run_length,
        'segments': segments,
        'fitting_angles': fitting_angles
    }

# Windows Forms dialog for hanger and spacing
class HangerSpacingDialog(Window):
    def __init__(self, button_names, lines):
        super(HangerSpacingDialog, self).__init__()
        self.Title = "Hanger and Spacing"
        self.Width = 390
        self.Height = 300
        self.WindowStartupLocation = WindowStartupLocation.CenterScreen
        self.ResizeMode = ResizeMode.NoResize
        stack = StackPanel()
        stack.Orientation = Orientation.Vertical
        stack.Margin = Thickness(10)
        label_hanger = Label()
        label_hanger.Content = "Choose Hanger:"
        label_hanger.FontSize = 12
        label_hanger.FontFamily = FontFamily("Arial")
        label_hanger.Margin = Thickness(0, 0, 0, 0)
        stack.Children.Add(label_hanger)
        self.combobox_hanger = ComboBox()
        self.combobox_hanger.Width = 350
        self.combobox_hanger.Height = 20
        self.combobox_hanger.FontSize = 12
        self.combobox_hanger.FontFamily = FontFamily("Arial")
        self.combobox_hanger.ItemsSource = Array[object](button_names)
        if lines[0] in button_names:
            self.combobox_hanger.SelectedItem = lines[0]
        else:
            self.combobox_hanger.SelectedItem = button_names[0]
        self.combobox_hanger.Margin = Thickness(0, 0, 0, 10)
        self.combobox_hanger.HorizontalAlignment = HorizontalAlignment.Left
        stack.Children.Add(self.combobox_hanger)
        label_end_dist = Label()
        label_end_dist.Content = "Distance from End (Ft):"
        label_end_dist.FontSize = 12
        label_end_dist.FontFamily = FontFamily("Arial")
        label_end_dist.Margin = Thickness(0, 0, 0, 0)
        stack.Children.Add(label_end_dist)
        self.textbox_end_dist = TextBox()
        self.textbox_end_dist.Width = 200
        self.textbox_end_dist.Height = 20
        self.textbox_end_dist.FontSize = 12
        self.textbox_end_dist.FontFamily = FontFamily("Arial")
        self.textbox_end_dist.Text = lines[1]
        self.textbox_end_dist.Margin = Thickness(0, 0, 0, 10)
        self.textbox_end_dist.HorizontalAlignment = HorizontalAlignment.Left
        stack.Children.Add(self.textbox_end_dist)
        label_spacing = Label()
        label_spacing.Content = "Hanger Spacing (Ft):"
        label_spacing.FontSize = 12
        label_spacing.FontFamily = FontFamily("Arial")
        label_spacing.Margin = Thickness(0, 0, 0, 0)
        stack.Children.Add(label_spacing)
        self.textbox_spacing = TextBox()
        self.textbox_spacing.Width = 200
        self.textbox_spacing.Height = 20
        self.textbox_spacing.FontSize = 12
        self.textbox_spacing.FontFamily = FontFamily("Arial")
        self.textbox_spacing.Text = lines[2]
        self.textbox_spacing.Margin = Thickness(0, 0, 0, 10)
        self.textbox_spacing.HorizontalAlignment = HorizontalAlignment.Left
        stack.Children.Add(self.textbox_spacing)
        self.checkbox_atos = CheckBox()
        self.checkbox_atos.Content = "Attach to Structure"
        self.checkbox_atos.FontSize = 12
        self.checkbox_atos.FontFamily = FontFamily("Arial")
        self.checkbox_atos.IsChecked = True
        self.checkbox_atos.Margin = Thickness(0, 0, 0, 5)
        stack.Children.Add(self.checkbox_atos)
        self.checkbox_support_joints = CheckBox()
        self.checkbox_support_joints.Content = "Support Joints"
        self.checkbox_support_joints.FontSize = 12
        self.checkbox_support_joints.FontFamily = FontFamily("Arial")
        self.checkbox_support_joints.IsChecked = lines[3].lower() == 'true'
        self.checkbox_support_joints.Margin = Thickness(0, 0, 0, 10)
        stack.Children.Add(self.checkbox_support_joints)
        self.button_ok = Button()
        self.button_ok.Content = "OK"
        self.button_ok.FontSize = 12
        self.button_ok.FontFamily = FontFamily("Arial")
        self.button_ok.Width = 74
        self.button_ok.Height = 25
        self.button_ok.HorizontalAlignment = HorizontalAlignment.Center
        self.button_ok.Click += self.ok_button_clicked
        stack.Children.Add(self.button_ok)
        self.Content = stack
    def ok_button_clicked(self, sender, event):
        self.DialogResult = True
        self.Close()

try:
    # Selection of starting fabrication part
    selected_ref = uidoc.Selection.PickObject(ObjectType.Element, FabricationPartSelectionFilter(), 'Select the starting Fabrication Part')
    element = doc.GetElement(selected_ref.ElementId)
    if not isinstance(element, FabricationPart):
        raise Exception("Selected element is not a FabricationPart.")
    pick_point = selected_ref.GlobalPoint
    start_connector = find_nearest_connector(element, pick_point)
    if not start_connector:
        raise Exception("Could not find a valid connector on the selected element.")
    
    # Get fabrication service and button setup
    parameters = element.LookupParameter('Fabrication Service').AsValueString()
    servicenamelist = []
    Config = FabricationConfiguration.GetFabricationConfiguration(doc)
    LoadedServices = Config.GetAllLoadedServices()
    for Item1 in LoadedServices:
        try:
            servicenamelist.append(Item1.Name)
        except:
            servicenamelist.append([])
    Servicenum = servicenamelist.index(parameters)
    FabricationService = LoadedServices[Servicenum]
    buttonnames = []
    button_data = []
    if RevitINT > 2022:
        palette_count = FabricationService.PaletteCount
        for group_idx in range(palette_count):
            buttoncount = FabricationService.GetButtonCount(group_idx)
            for btn_idx in range(buttoncount):
                bt = FabricationService.GetButton(group_idx, btn_idx)
                if bt.IsAHanger:
                    buttonnames.append(bt.Name)
                    button_data.append((group_idx, btn_idx, bt.Name))
    else:
        group_count = FabricationService.GroupCount
        for group_idx in range(group_count):
            buttoncount = FabricationService.GetButtonCount(group_idx)
            for btn_idx in range(buttoncount):
                bt = FabricationService.GetButton(group_idx, btn_idx)
                if bt.IsAHanger:
                    buttonnames.append(bt.Name)
                    button_data.append((group_idx, btn_idx, bt.Name))
    
    folder_name = "c:\\Temp"
    filepath = os.path.join(folder_name, 'Ribbon_PlaceHangers.txt')
    if not os.path.exists(folder_name):
        os.makedirs(folder_name)
    if not os.path.exists(filepath):
        with open(filepath, 'w') as the_file:
            line1 = (str(buttonnames[0]) + '\n')
            line2 = ('1' + '\n')
            line3 = ('4' + '\n')
            line4 = 'True'
            the_file.writelines([line1, line2, line3, line4])
    with open(filepath, 'r') as file:
        lines = file.readlines()
        lines = [line.rstrip() for line in lines]
    if len(lines) < 4:
        with open(filepath, 'w') as the_file:
            line1 = (str(buttonnames[0]) + '\n')
            line2 = ('1' + '\n')
            line3 = '4' + '\n'
            line4 = 'True'
            the_file.writelines([line1, line2, line3, line4])
        with open(filepath, 'r') as file:
            lines = file.readlines()
            lines = [line.rstrip() for line in lines]
    
    # Show the Windows Forms dialog
    form = HangerSpacingDialog(buttonnames, lines)
    if form.ShowDialog():
        Selectedbutton = form.combobox_hanger.SelectedItem
        distancefromend = float(form.textbox_end_dist.Text)
        Spacing = float(form.textbox_spacing.Text)
        AtoS = form.checkbox_atos.IsChecked
        SupportJoint = form.checkbox_support_joints.IsChecked
        
        # Write values to text file for future retrieval
        with open(filepath, 'w') as the_file:
            line1 = (Selectedbutton + '\n')
            line2 = (str(distancefromend) + '\n')
            line3 = str(Spacing) + '\n'
            line4 = str(SupportJoint)
            the_file.writelines([line1, line2, line3, line4])
        
        for group_idx, btn_idx, btn_name in button_data:
            if btn_name == Selectedbutton:
                Servicegroupnum = group_idx
                Buttonnum = btn_idx
                break
        
        validbutton = FabricationService.IsValidButtonIndex(Servicegroupnum, Buttonnum)
        FabricationServiceButton = FabricationService.GetButton(Servicegroupnum, Buttonnum)
        
        selected_refs = uidoc.Selection.PickObjects(ObjectType.Element, CustomISelectionFilter(parameters), "Select additional Fabrication Parts to place hangers on")
        selected_elements = [doc.GetElement( elId ) for elId in selected_refs]
       
        # Ensure start element is included
        if element.Id not in [e.Id for e in selected_elements]:
            selected_elements.insert(0, element)
       
        # Order selected elements by connector proximity and retrieve entry connectors
        ordered_elements, entry_connectors = order_selected_elements(selected_elements, element, start_connector)
        if not ordered_elements:
            raise Exception("No valid elements selected for the run.")
        
        run_info = analyze_run(ordered_elements, element, start_connector)
        start_point = run_info['start_point']
        end_point = run_info['end_point']
        total_run_length = run_info['total_length']
        segments = run_info['segments']
        fitting_angles = run_info['fitting_angles']
        
        t = Transaction(doc, 'Place Hangers')
        t.Start()
        
        if SupportJoint:
            # Original per pipe logic unchanged works as expected
            for e in ordered_elements:
                if e.LookupParameter('Part Pattern Number').AsInteger() in (2041, 866, 40):
                    if not vertical_fab(e):
                        pipelen = e.CenterlineLength
                        pipe_connectors = e.ConnectorManager.Connectors
                        if not pipe_connectors or pipe_connectors.Size == 0:
                            continue
                        if pipelen < 2 * distancefromend:
                            try:
                                center_position = pipelen / 2
                                pipe_connector = next(iter(pipe_connectors))
                                FabricationPart.CreateHanger(doc, FabricationServiceButton, e.Id, pipe_connector, center_position, AtoS)
                            except Exception as ex:
                                pass
                        else:
                            try:
                                for connector in pipe_connectors:
                                    FabricationPart.CreateHanger(
                                        doc,
                                        FabricationServiceButton,
                                        e.Id,
                                        connector,
                                        distancefromend,
                                        AtoS
                                    )
                                if pipelen > (Spacing + (2 * distancefromend)):
                                    qtyofhgrs = range(int((math.floor(pipelen) - (2 * distancefromend)) / Spacing))
                                    IncrementSpacing = distancefromend
                                    pipe_connector = next(iter(pipe_connectors))
                                    for hgr in qtyofhgrs:
                                        IncrementSpacing += Spacing
                                        FabricationPart.CreateHanger(doc, FabricationServiceButton, e.Id, pipe_connector, IncrementSpacing, AtoS)
                            except Exception as ex:
                                pass
        else:
            # Continuous run logic with improved end handling
            placed_positions = []
            last_placed = 0.0
            
            def place_hanger_at_global(target_global):
                """Attempt to place hanger at or near target_global position.
                   Returns actual global position placed or None."""
                for seg_idx, (e, seg_start, seg_len, is_fabpart) in enumerate(segments):
                    seg_end = seg_start + seg_len
                    if seg_start <= target_global < seg_end:
                        if is_fabpart and not vertical_fab(e):
                            local_pos = target_global - seg_start
                            connector = entry_connectors.get(e.Id)
                            if connector and 0.01 <= local_pos <= seg_len - 0.01:
                                try:
                                    FabricationPart.CreateHanger(
                                        doc, FabricationServiceButton, e.Id, connector, local_pos, AtoS)
                                    return target_global
                                except:
                                    pass
                        # Shift to nearest valid pipe edge if on fitting or placement failed
                        prev_pipe = None
                        prev_end_global = seg_start
                        prev_idx = seg_idx - 1
                        while prev_idx >= 0:
                            prev_e, prev_start, prev_len, prev_is_fabpart = segments[prev_idx]
                            if prev_is_fabpart and not vertical_fab(prev_e):
                                prev_pipe = prev_e
                                prev_end_global = prev_start + prev_len
                                break
                            prev_idx -= 1
                        next_pipe = None
                        next_start_global = seg_end
                        next_idx = seg_idx + 1
                        while next_idx < len(segments):
                            next_e, next_start, _, next_is_fabpart = segments[next_idx]
                            if next_is_fabpart and not vertical_fab(next_e):
                                next_pipe = next_e
                                next_start_global = next_start
                                break
                            next_idx += 1
                        
                        dist_to_prev = abs(target_global - prev_end_global) if prev_pipe else float('inf')
                        dist_to_next = abs(next_start_global - target_global) if next_pipe else float('inf')
                        
                        if dist_to_prev <= dist_to_next and prev_pipe:
                            local_pos = prev_pipe.CenterlineLength - 0.1
                            connector = entry_connectors.get(prev_pipe.Id)
                            if connector:
                                try:
                                    FabricationPart.CreateHanger(
                                        doc, FabricationServiceButton, prev_pipe.Id, connector, local_pos, AtoS)
                                    return prev_end_global - 0.1
                                except:
                                    pass
                        elif next_pipe:
                            local_pos = 0.1
                            connector = entry_connectors.get(next_pipe.Id)
                            if connector:
                                try:
                                    FabricationPart.CreateHanger(
                                        doc, FabricationServiceButton, next_pipe.Id, connector, local_pos, AtoS)
                                    return next_start_global + 0.1
                                except:
                                    pass
                return None
            
            # Force hanger near start on first pipe segment
            first_pipe_idx = next((i for i, (_, _, _, is_fabpart) in enumerate(segments) if is_fabpart and not vertical_fab(segments[i][0])), None)
            if first_pipe_idx is not None:
                e, seg_start, seg_len, _ = segments[first_pipe_idx]
                connector = entry_connectors.get(e.Id)
                if connector and seg_len > 0.5:  # Minimum practical length
                    if seg_len < 2 * distancefromend:
                        local_pos = seg_len / 2
                    else:
                        local_pos = distancefromend
                    try:
                        FabricationPart.CreateHanger(doc, FabricationServiceButton, e.Id, connector, local_pos, AtoS)
                        actual_global = seg_start + local_pos
                        placed_positions.append(actual_global)
                        last_placed = actual_global
                    except:
                        pass
            
            # Place intermediate hangers
            proposed = last_placed + Spacing
            while proposed < total_run_length - distancefromend:
                actual = place_hanger_at_global(proposed)
                if actual is not None and abs(actual - last_placed) > 1.0:
                    placed_positions.append(actual)
                    last_placed = actual
                proposed = last_placed + Spacing
            
            # Force hanger near end on last pipe segment, if not already close
            last_pipe_idx = next((i for i in range(len(segments)-1, -1, -1) if segments[i][3] and not vertical_fab(segments[i][0])), None)
            if last_pipe_idx is not None:
                e, seg_start, seg_len, _ = segments[last_pipe_idx]
                connector = entry_connectors.get(e.Id)
                if connector and seg_len > 0.5:
                    if seg_len < 2 * distancefromend:
                        local_pos = seg_len / 2
                    else:
                        local_pos = seg_len - distancefromend
                    try:
                        global_pos = seg_start + local_pos
                        if not placed_positions or abs(global_pos - placed_positions[-1]) > 1.0:
                            FabricationPart.CreateHanger(doc, FabricationServiceButton, e.Id, connector, local_pos, AtoS)
                            placed_positions.append(global_pos)
                    except:
                        pass
        
        t.Commit()

except Exception as ex:
    pass