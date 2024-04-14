
import Autodesk
from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory, Transaction, FabricationPart, FabricationConfiguration
from SharedParam.Add_Parameters import Shared_Params
from Parameters.Get_Set_Params import set_parameter_by_name, get_parameter_value_by_name_AsString, get_parameter_value_by_name_AsInteger

Shared_Params()

DB = Autodesk.Revit.DB
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
curview = doc.ActiveView
fec = FilteredElementCollector
app = doc.Application
RevitVersion = app.VersionNumber
RevitINT = float (RevitVersion)
Config = FabricationConfiguration.GetFabricationConfiguration(doc)

# --------------SELECTED ELEMENTS-------------------

selection = [doc.GetElement(id) for id in __revit__.ActiveUIDocument.Selection.GetElementIds()]

if selection:
    t = Transaction(doc, "Update FP Parameters")
    t.Start()
    for x in selection:
        isfabpart = x.LookupParameter("Fabrication Service")
        if isfabpart:
            set_parameter_by_name(x, 'FP_CID', x.ItemCustomId)
            set_parameter_by_name(x, 'FP_Service Type', Config.GetServiceTypeName(x.ServiceType))
            set_parameter_by_name(x, 'FP_Service Name', get_parameter_value_by_name_AsString(x, 'Fabrication Service Name'))
            set_parameter_by_name(x, 'FP_Service Abbreviation', get_parameter_value_by_name_AsString(x, 'Fabrication Service Abbreviation'))
            if x.ItemCustomId == 838:
                set_parameter_by_name(x, 'FP_Hanger Diameter', get_parameter_value_by_name_AsString(x, 'Size of Primary End'))
                set_parameter_by_name(x, 'FP_Rod Attached', 'Yes') if x.GetRodInfo().IsAttachedToStructure else set_parameter_by_name(x, 'FP_Rod Attached', 'No')
                [set_parameter_by_name(x, 'FP_Rod Size', n.AncillaryWidthOrDiameter) for n in x.GetPartAncillaryUsage() if n.AncillaryWidthOrDiameter > 0]
            if x.ItemCustomId != 838:
                set_parameter_by_name(x, 'FP_Centerline Length', x.CenterlineLength)
            try:
                if (x.GetRodInfo().RodCount) < 2:
                    ItmDims = x.GetDimensions()
                    for dta in ItmDims:
                        if dta.Name == 'Rod Extn Below':
                            RB = x.GetDimensionValue(dta)
                        if dta.Name == 'Rod Length':
                            RL = x.GetDimensionValue(dta)
                    TRL = RL + RB
                    set_parameter_by_name(x, 'FP_Rod Length', TRL)
            except:
                pass

            try:
                if (x.GetRodInfo().RodCount) > 1:
                    ItmDims = x.GetDimensions()
                    for dta in ItmDims:
                        if dta.Name == 'Length A':
                            RL = x.GetDimensionValue(dta)
                        if dta.Name == 'Width':
                            TrapWidth = x.GetDimensionValue(dta)
                        if dta.Name == 'Bearer Extn':
                            TrapExtn = x.GetDimensionValue(dta)
                        if dta.Name == 'Right Rod Offset':
                            TrapRRod = x.GetDimensionValue(dta)
                        if dta.Name == 'Left Rod Offset':
                            TrapLRod = x.GetDimensionValue(dta)
                    BearerLength = TrapWidth + TrapExtn + TrapExtn
                    set_parameter_by_name(x, 'FP_Bearer Length', BearerLength)
                    set_parameter_by_name(x, 'FP_Rod Length', RL)
            except:
                pass
    t.Commit()

# --------------SELECTED ELEMENTS-------------------
else:
    # --------------ACTIVE VIEW-------------------

    # Creating collector instance and collecting all the fabrication hangers from the model
    hanger_collector = FilteredElementCollector(doc, curview.Id).OfCategory(BuiltInCategory.OST_FabricationHangers) \
                       .WhereElementIsNotElementType() \
                       .ToElements()

    # Create a FilteredElementCollector to get all FabricationPart elements
    AllElements = FilteredElementCollector(doc, curview.Id).OfClass(FabricationPart) \
                       .WhereElementIsNotElementType() \
                       .ToElements()

    t = Transaction(doc, "Update FP Parameters")
    t.Start()

    for hanger in hanger_collector:
        try:
            if (hanger.GetRodInfo().RodCount) < 2:
                # Get the host element's size
                hosted_info = hanger.GetHostedInfo().HostId
                HostSize = get_parameter_value_by_name_AsString(doc.GetElement(hosted_info), 'Size')
                # Get the hanger's size
                HangerSize = get_parameter_value_by_name_AsString(hanger, 'Size of Primary End')
                # Set the 'FP_Hanger Shield' parameter based on whether the host and hanger sizes match
                if HostSize == HangerSize:
                    set_parameter_by_name(hanger, 'FP_Hanger Shield', 'No')
                else:
                    set_parameter_by_name(hanger, 'FP_Hanger Shield', 'Yes')
        except:
            pass

        try:
            if (hanger.GetRodInfo().RodCount) < 2:
                ItmDims = hanger.GetDimensions()
                for dta in ItmDims:
                    if dta.Name == 'Rod Extn Below':
                        RB = hanger.GetDimensionValue(dta)
                    if dta.Name == 'Rod Length':
                        RL = hanger.GetDimensionValue(dta)
                TRL = RL + RB
                set_parameter_by_name(hanger, 'FP_Rod Length', TRL)
        except:
            pass

        try:
            if (hanger.GetRodInfo().RodCount) > 1:
                ItmDims = hanger.GetDimensions()
                for dta in ItmDims:
                    if dta.Name == 'Length A':
                        RL = hanger.GetDimensionValue(dta)
                    if dta.Name == 'Width':
                        TrapWidth = hanger.GetDimensionValue(dta)
                    if dta.Name == 'Bearer Extn':
                        TrapExtn = hanger.GetDimensionValue(dta)
                    if dta.Name == 'Right Rod Offset':
                        TrapRRod = hanger.GetDimensionValue(dta)
                    if dta.Name == 'Left Rod Offset':
                        TrapLRod = hanger.GetDimensionValue(dta)
                BearerLength = TrapWidth + TrapExtn + TrapExtn
                set_parameter_by_name(hanger, 'FP_Bearer Length', BearerLength)
                set_parameter_by_name(hanger, 'FP_Rod Length', RL)
        except:
            pass


    try:
        list(map(lambda x: set_parameter_by_name(x, 'FP_CID', x.ItemCustomId), AllElements))
        list(map(lambda x: set_parameter_by_name(x, 'FP_Centerline Length', x.CenterlineLength), AllElements))
        list(map(lambda x: set_parameter_by_name(x, 'FP_Service Type', Config.GetServiceTypeName(x.ServiceType)), AllElements))
        list(map(lambda x: set_parameter_by_name(x, 'FP_Service Name', get_parameter_value_by_name_AsString(x, 'Fabrication Service Name')), AllElements))
        list(map(lambda x: set_parameter_by_name(x, 'FP_Service Abbreviation', get_parameter_value_by_name_AsString(x, 'Fabrication Service Abbreviation')), AllElements))
        list(map(lambda x: set_parameter_by_name(x, 'FP_Hanger Diameter', get_parameter_value_by_name_AsString(x, 'Size of Primary End')), hanger_collector))
        list(map(lambda x: set_parameter_by_name(x, 'FP_Rod Attached', 'Yes') if x.GetRodInfo().IsAttachedToStructure else set_parameter_by_name(x, 'FP_Rod Attached', 'No'), hanger_collector))
        list(map(lambda hanger: [set_parameter_by_name(hanger, 'FP_Rod Size', n.AncillaryWidthOrDiameter) for n in hanger.GetPartAncillaryUsage() if n.AncillaryWidthOrDiameter > 0], hanger_collector))
    except:
        pass
    t.Commit()
    # --------------ACTIVE VIEW-------------------