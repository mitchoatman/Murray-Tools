import Autodesk
from Autodesk.Revit.DB import Transaction, FabricationConfiguration, BuiltInParameter, FabricationPart, FabricationServiceButton, FabricationService, XYZ, ConnectorType
from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType
from rpw.ui.forms import FlexForm, Label, ComboBox, TextBox, Separator, Button, CheckBox
import math
import os

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

# Function to order selected elements by connector proximity
def order_selected_elements(selected_elements, start_element, start_connector):
    ordered_elements = [start_element]
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
            remaining_elements.remove(next_element.Id)
            current_element = next_element
            # Find the opposite connector on the current element
            for conn in current_element.ConnectorManager.Connectors:
                if conn != next_connector:
                    current_connector = conn
                    break
        else:
            break
    
    # print("Ordered elements: {}".format([e.Id.IntegerValue for e in ordered_elements]))
    return ordered_elements

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
    # print("Ordered run includes elements: {}".format([e.Id.IntegerValue for e in ordered_run]))

    # Collect segment info and calculate length
    for e in ordered_run:
        length = e.CenterlineLength
        is_pipe = e.LookupParameter('Part Pattern Number').AsInteger() in (2041, 866, 40)
        segments.append((e, current_position, length, is_pipe))
        total_run_length += length
        current_position += length
        # print("Segment: Element ID {}, Start Pos {}, Length {}, Is Pipe {}".format(
            # e.Id.IntegerValue, current_position - length, length, is_pipe))
        if e == ordered_run[-1]:
            connectors = e.ConnectorManager.Connectors
            for conn in connectors:
                if conn.Origin.DistanceTo(start_point) > end_point.DistanceTo(start_point):
                    end_point = conn.Origin

    # Analyze directions and angles
    for i, (e, start_pos, length, is_pipe) in enumerate(segments):
        if is_pipe:
            connectors = e.ConnectorManager.Connectors
            conn_points = [conn.Origin for conn in connectors]
            if len(conn_points) >= 2:
                p1, p2 = conn_points[:2]
                direction = (p2 - p1).Normalize()
                pipe_directions.append((e, direction))
                if i < len(segments) - 1:
                    next_e, _, _, next_is_pipe = segments[i + 1]
                    if next_is_pipe:
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
            line3 = '4' + '\n'
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

    if lines[0] in buttonnames:
        components = [
            Label('Choose Hanger:'),
            ComboBox('Buttonnum', buttonnames, sort=False, default=lines[0]),
            Label('Distance from End (Ft):'),
            TextBox('EndDist', lines[1]),
            Label('Hanger Spacing (Ft):'),
            TextBox('Spacing', lines[2]),
            CheckBox('checkboxvalue', 'Attach to Structure', default=True),
            CheckBox('checkboxjointvalue', 'Support Joints', default=lines[3].lower() == 'true'),
            Button('Ok')
        ]
        form = FlexForm('Hanger and Spacing', components)
        form.show()
    else:
        components = [
            Label('Choose Hanger:'),
            ComboBox('Buttonnum', buttonnames, sort=False),
            Label('Distance from End (Ft):'),
            TextBox('EndDist', lines[1]),
            Label('Hanger Spacing (Ft):'),
            TextBox('Spacing', lines[2]),
            CheckBox('checkboxvalue', 'Attach to Structure', default=True),
            CheckBox('checkboxjointvalue', 'Support Joints', default=lines[3].lower() == 'true'),
            Button('Ok')
        ]
        form = FlexForm('Hanger and Spacing', components)
        form.show()

    Selectedbutton = form.values['Buttonnum']
    for group_idx, btn_idx, btn_name in button_data:
        if btn_name == Selectedbutton:
            Servicegroupnum = group_idx
            Buttonnum = btn_idx
            break

    distancefromend = float(form.values['EndDist'])
    Spacing = float(form.values['Spacing'])
    AtoS = form.values['checkboxvalue']
    SupportJoint = form.values['checkboxjointvalue']

    with open(filepath, 'w') as the_file:
        line1 = (Selectedbutton + '\n')
        line2 = (str(distancefromend) + '\n')
        line3 = str(Spacing) + '\n'
        line4 = str(SupportJoint)
        the_file.writelines([line1, line2, line3, line4])

    validbutton = FabricationService.IsValidButtonIndex(Servicegroupnum, Buttonnum)
    FabricationServiceButton = FabricationService.GetButton(Servicegroupnum, Buttonnum)

    selected_refs = uidoc.Selection.PickObjects(ObjectType.Element, CustomISelectionFilter(parameters), "Select additional Fabrication Parts to place hangers on")            
    selected_elements = [doc.GetElement( elId ) for elId in selected_refs]
    
    # Ensure start element is included
    if element.Id not in [e.Id for e in selected_elements]:
        selected_elements.insert(0, element)
    
    # Order selected elements by connector proximity
    ordered_elements = order_selected_elements(selected_elements, element, start_connector)
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
        for e in ordered_elements:
            if e.LookupParameter('Part Pattern Number').AsInteger() in (2041, 866, 40):
                if not vertical_fab(e):
                    pipelen = e.CenterlineLength
                    pipe_connectors = e.ConnectorManager.Connectors
                    if not pipe_connectors or pipe_connectors.Size == 0:
                        print("No valid connectors found for pipe ID {}".format(e.Id.IntegerValue))
                        continue
                    # Check if pipe is shorter than 2 * distancefromend
                    if pipelen < 2 * distancefromend:
                        try:
                            # Place one hanger in the center
                            center_position = pipelen / 2
                            pipe_connector = next(iter(pipe_connectors))
                            FabricationPart.CreateHanger(doc, FabricationServiceButton, e.Id, pipe_connector, center_position, AtoS)
                        except Exception as ex:
                            print("Failed to place center hanger at position {} on pipe ID {}: {}".format(
                                center_position, e.Id.IntegerValue, str(ex)))
                    else:
                        try:
                            # Place hangers at distancefromend from each connector
                            for connector in pipe_connectors:
                                FabricationPart.CreateHanger(
                                    doc,
                                    FabricationServiceButton,
                                    e.Id,
                                    connector,
                                    distancefromend,
                                    AtoS
                                )
                            # Place additional hangers at Spacing intervals
                            if pipelen > (Spacing + (2 * distancefromend)):
                                qtyofhgrs = range(int((math.floor(pipelen) - (2 * distancefromend)) / Spacing))
                                IncrementSpacing = distancefromend
                                pipe_connector = next(iter(pipe_connectors))
                                for hgr in qtyofhgrs:
                                    IncrementSpacing += Spacing
                                    FabricationPart.CreateHanger(doc, FabricationServiceButton, e.Id, pipe_connector, IncrementSpacing, AtoS)
                        except Exception as ex:
                            print("Failed to place hanger on pipe ID {}: {}".format(e.Id.IntegerValue, str(ex)))
    else:
        placed_hangers = []
        last_hanger_position = None
        last_hanger_point = None
        last_connector = None
        last_pipe_id = None

        # Helper function to find nearest connector to a point
        def find_nearest_connector_to_point(element, target_point):
            connector_manager = element.ConnectorManager
            if not connector_manager:
                return None
            nearest_connector = None
            min_distance = float('inf')
            for connector in connector_manager.Connectors:
                distance = connector.Origin.DistanceTo(target_point)
                if distance < min_distance:
                    min_distance = distance
                    nearest_connector = connector
            return nearest_connector

        # Place first hanger
        if element.LookupParameter('Part Pattern Number').AsInteger() in (2041, 866, 40) and not vertical_fab(element):
            pipelen = element.CenterlineLength
            if pipelen > distancefromend:
                local_position = distancefromend
                segment_start = next((start for e, start, _, _ in segments if e.Id == element.Id), 0)
                global_position = segment_start + local_position
                fitting_zones = [(start, start + length) for e, start, length, is_pipe in segments if not is_pipe]
                is_valid_position = True
                for start, end in fitting_zones:
                    if (start - 0.5) <= global_position <= (end + 0.5):
                        is_valid_position = False
                        global_position = end + 0.5
                        local_position = global_position - segment_start
                        break
                try:
                    # Calculate physical position of the hanger
                    connectors = element.ConnectorManager.Connectors
                    connector_points = [conn.Origin for conn in connectors]
                    if len(connector_points) >= 2:
                        p1, p2 = connector_points[:2]
                        direction = (p2 - p1).Normalize()
                        hanger_point = start_connector.Origin + direction * local_position
                    else:
                        hanger_point = start_connector.Origin
                    FabricationPart.CreateHanger(
                        doc,
                        FabricationServiceButton,
                        element.Id,
                        start_connector,
                        local_position,
                        AtoS
                    )
                    placed_hangers.append(global_position)
                    last_hanger_position = global_position
                    last_hanger_point = hanger_point
                    last_connector = start_connector
                    last_pipe_id = element.Id

                except:
                    print("Failed to place first hanger at position {} on pipe ID {}".format(
                        global_position, element.Id.IntegerValue))

        # Place subsequent hangers
        if last_hanger_position is not None:
            proposed_position = last_hanger_position + Spacing
            bend_hangers_placed = set()  # Track bends with hangers placed

            while proposed_position < total_run_length:
                # Find the segment containing the proposed position
                current_segment = None
                segment_start = 0
                local_position = 0
                is_pipe_segment = False
                for e, start, length, is_pipe in segments:
                    segment_end = start + length
                    if start <= proposed_position < segment_end:
                        current_segment = e
                        segment_start = start
                        local_position = proposed_position - start
                        is_pipe_segment = is_pipe
                        break

                # Handle non-pipe segments (fittings, couplings, tees)
                if not is_pipe_segment and current_segment:
                    # Check if it's a bend fitting
                    is_bend = any(fitting.Id == current_segment.Id for fitting, _ in fitting_angles)
                    if is_bend and current_segment.Id not in bend_hangers_placed:
                        # Find the previous and next pipe segments
                        fitting_end = segment_start + length
                        prev_segment_idx = next((i for i in range(len(segments)-1, -1, -1) if segments[i][1] < segment_start and segments[i][3]), None)
                        next_segment_idx = next((i for i, (e, start, _, is_pipe) in enumerate(segments) if start > fitting_end and is_pipe), None)
                        
                        # Place hanger on previous pipe (if available)
                        if prev_segment_idx is not None:
                            prev_pipe, prev_start, prev_length, _ = segments[prev_segment_idx]
                            if prev_length > distancefromend:
                                # Select connector closest to the fitting
                                fitting_connectors = current_segment.ConnectorManager.Connectors
                                prev_pipe_connectors = prev_pipe.ConnectorManager.Connectors
                                prev_pipe_end_point = next(iter(prev_pipe_connectors)).Origin
                                fitting_start_connector = find_nearest_connector_to_point(current_segment, prev_pipe_end_point)
                                if fitting_start_connector:
                                    pipe_connector = find_nearest_connector_to_point(prev_pipe, fitting_start_connector.Origin)
                                    if pipe_connector:
                                        hanger_position = prev_start + prev_length - distancefromend
                                        local_position = prev_length - distancefromend
                                        try:
                                            # Calculate physical position
                                            connectors = prev_pipe.ConnectorManager.Connectors
                                            connector_points = [conn.Origin for conn in connectors]
                                            if len(connector_points) >= 2:
                                                p1, p2 = connector_points[:2]
                                                direction = (p2 - p1).Normalize()
                                                hanger_point = pipe_connector.Origin + direction * local_position
                                            else:
                                                hanger_point = pipe_connector.Origin
                                            FabricationPart.CreateHanger(
                                                doc,
                                                FabricationServiceButton,
                                                prev_pipe.Id,
                                                pipe_connector,
                                                local_position,
                                                AtoS
                                            )
                                            placed_hangers.append(hanger_position)
                                            if hanger_position > last_hanger_position:
                                                last_hanger_position = hanger_position
                                                last_hanger_point = hanger_point
                                                last_connector = pipe_connector
                                                last_pipe_id = prev_pipe.Id
                                        except:
                                            print("Failed to place bend hanger at position {} on pipe ID {}".format(
                                                hanger_position, prev_pipe.Id.IntegerValue))

                        # Place hanger on next pipe (if available)
                        if next_segment_idx is not None:
                            next_pipe, next_start, next_length, _ = segments[next_segment_idx]
                            if next_length > distancefromend:
                                # Select connector closest to the fitting
                                fitting_connectors = current_segment.ConnectorManager.Connectors
                                next_pipe_connectors = next_pipe.ConnectorManager.Connectors
                                next_pipe_start_point = next(iter(next_pipe_connectors)).Origin
                                fitting_end_connector = find_nearest_connector_to_point(current_segment, next_pipe_start_point)
                                if fitting_end_connector:
                                    pipe_connector = find_nearest_connector_to_point(next_pipe, fitting_end_connector.Origin)
                                    if pipe_connector:
                                        hanger_position = next_start + distancefromend
                                        local_position = distancefromend
                                        try:
                                            # Calculate physical position
                                            connectors = next_pipe.ConnectorManager.Connectors
                                            connector_points = [conn.Origin for conn in connectors]
                                            if len(connector_points) >= 2:
                                                p1, p2 = connector_points[:2]
                                                direction = (p2 - p1).Normalize()
                                                hanger_point = pipe_connector.Origin + direction * local_position
                                            else:
                                                hanger_point = pipe_connector.Origin
                                            FabricationPart.CreateHanger(
                                                doc,
                                                FabricationServiceButton,
                                                next_pipe.Id,
                                                pipe_connector,
                                                local_position,
                                                AtoS
                                            )
                                            placed_hangers.append(hanger_position)
                                            print("Placed bend hanger at position {} (local {}) on pipe ID {} after bend fitting ID {}, connector at X={}, Y={}, Z={}, spacing from last: {}".format(
                                                hanger_position, local_position, next_pipe.Id.IntegerValue, current_segment.Id.IntegerValue,
                                                pipe_connector.Origin.X, pipe_connector.Origin.Y, pipe_connector.Origin.Z,
                                                hanger_position - last_hanger_position))
                                            # Update last hanger info if this is the latest position
                                            if hanger_position > last_hanger_position:
                                                last_hanger_position = hanger_position
                                                last_hanger_point = hanger_point
                                                last_connector = pipe_connector
                                                last_pipe_id = next_pipe.Id
                                        except:
                                            print("Failed to place bend hanger at position {} on pipe ID {}".format(
                                                hanger_position, next_pipe.Id.IntegerValue))
                        
                        # Mark bend as processed and advance proposed_position
                        bend_hangers_placed.add(current_segment.Id)
                        proposeddrawn_position = last_hanger_position + Spacing
                    else:
                        # Skip couplings/tees or bends already processed
                        proposed_position = segment_end + Spacing
                    continue

                # Place hanger on pipe segment
                if is_pipe_segment and local_position <= current_segment.CenterlineLength:
                    # Check if proposed position is too close to an existing hanger
                    too_close = any(abs(proposed_position - p) < 0.5 for p in placed_hangers)
                    if too_close:
                        proposed_position += Spacing
                        continue
                    
                    # Use the same connector as the last hanger if on the same pipe, else find nearest
                    if current_segment.Id == last_pipe_id and last_connector:
                        pipe_connector = last_connector
                    else:
                        pipe_connector = find_nearest_connector_to_point(current_segment, last_hanger_point)
                    
                    if pipe_connector:
                        try:
                            # Calculate physical position
                            connectors = current_segment.ConnectorManager.Connectors
                            connector_points = [conn.Origin for conn in connectors]
                            if len(connector_points) >= 2:
                                p1, p2 = connector_points[:2]
                                direction = (p2 - p1).Normalize()
                                hanger_point = pipe_connector.Origin + direction * local_position
                            else:
                                hanger_point = pipe_connector.Origin
                            FabricationPart.CreateHanger(
                                doc,
                                FabricationServiceButton,
                                current_segment.Id,
                                pipe_connector,
                                local_position,
                                AtoS
                            )
                            placed_hangers.append(proposed_position)
                            last_hanger_position = proposed_position
                            last_hanger_point = hanger_point
                            last_connector = pipe_connector
                            last_pipe_id = current_segment.Id
                            proposed_position = last_hanger_position + Spacing
                        except:
                            print("Failed to place hanger at position {} on pipe ID {}".format(
                                proposed_position, current_segment.Id.IntegerValue))
                            proposed_position = last_hanger_position + Spacing
                    else:
                        proposed_position = segment_end + Spacing
                else:
                    if current_segment:
                        proposed_position = segment_end + Spacing
                    else:
                        print("No segment found for position {}; total run length: {}".format(proposed_position, total_run_length))
                        break

    t.Commit()

except Exception as ex:
    print("Error: {}".format(str(ex)))