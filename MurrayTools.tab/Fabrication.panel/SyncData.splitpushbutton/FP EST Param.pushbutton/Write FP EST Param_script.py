
import Autodesk
from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory, Transaction, FabricationPart, FabricationConfiguration, BuiltInParameter
from Parameters.Add_SharedParameters import Shared_Params
from Parameters.Get_Set_Params import set_parameter_by_name, get_parameter_value_by_name_AsString, get_parameter_value_by_name_AsInteger, get_parameter_value_by_name_AsValueString, get_parameter_value_by_name_AsDouble
import math

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
            try:
                set_parameter_by_name(x, 'FP_CID', x.ItemCustomId)
            except:
                pass

            try:
                set_parameter_by_name(x, 'FP_Service Type', Config.GetServiceTypeName(x.ServiceType))
            except:
                pass

            try:
                set_parameter_by_name(x, 'FP_Service Name', get_parameter_value_by_name_AsString(x, 'Fabrication Service Name'))
            except:
                pass

            try:
                service_abbreviation = get_parameter_value_by_name_AsString(x, 'Fabrication Service Abbreviation')
                if service_abbreviation:
                    set_parameter_by_name(x, 'FP_Service Abbreviation', service_abbreviation)
            except:
                pass

            try:
                if x.LookupParameter('Product Entry'):
                    set_parameter_by_name(x, 'FP_Product Entry', get_parameter_value_by_name_AsString(x, 'Product Entry'))
                else:
                    set_parameter_by_name(x, 'FP_Product Entry', get_parameter_value_by_name_AsString(x, 'Size'))
            except:
                pass

            try:
                if x.ItemCustomId == 838:
                    set_parameter_by_name(x, 'FP_Rod Attached', 'Yes') if x.GetRodInfo().IsAttachedToStructure else set_parameter_by_name(x, 'FP_Rod Attached', 'No')
                    [set_parameter_by_name(x, 'FP_Rod Size', n.AncillaryWidthOrDiameter) for n in x.GetPartAncillaryUsage() if n.AncillaryWidthOrDiameter > 0]

                # Galv nut and flat washer
                mappings = {
                    0.03125: {'rod': '334BL3150X', 'nut': '334BL3048X', 'washer': '334BL3048H'}, # 3/8
                    0.04167: {'rod': '334BL3150Y', 'nut': '334BL3048Y', 'washer': '334BL3048I'}, # 1/2
                    0.05208: {'rod': '334BL3150Z', 'nut': '334BL3048Z', 'washer': '334BL3048J'}, # 5/8
                    0.06250: {'rod': '334BL3150P', 'nut': '334BL3049A', 'washer': '334BL3048K'}, # 3/4
                    0.07292: {'rod': '334BL3150X', 'nut': '334BL3049B', 'washer': '334BL3048L'}, # 7/8
                    0.08333: {'rod': '334BL3151B', 'nut': '334BL3049C', 'washer': '334BL3048M'}  # 1
                }

                # Set all parameters in one loop
                for n in x.GetPartAncillaryUsage():
                    if n.AncillaryWidthOrDiameter > 0:
                        rounded_size = round(n.AncillaryWidthOrDiameter, 5)
                        if rounded_size in mappings:
                            codes = mappings[rounded_size]
                            set_parameter_by_name(x, 'FP_Code Rod', codes['rod'])
                            set_parameter_by_name(x, 'FP_Code Nut', codes['nut'])
                            set_parameter_by_name(x, 'FP_Code Washer', codes['washer'])

                    ProductEntry = x.LookupParameter('Product Entry')
                    if ProductEntry:
                        if (x.GetRodInfo().RodCount) < 2:
                            # Get the host element's size
                            hosted_info = x.GetHostedInfo().HostId
                            try:
                                HostSize = get_parameter_value_by_name_AsString(doc.GetElement(hosted_info), 'Size').strip('"')
                                # Get the hanger's size
                                HangerSize = get_parameter_value_by_name_AsString(x, 'Product Entry')
                                set_parameter_by_name(x, 'FP_Product Entry', HangerSize)
                                # Set the 'FP_Hanger Shield' parameter based on whether the host and hanger sizes match
                                if HostSize == HangerSize:
                                    set_parameter_by_name(x, 'FP_Hanger Shield', 'No')
                                else:
                                    set_parameter_by_name(x, 'FP_Hanger Shield', 'Yes')
                                    set_parameter_by_name(x, 'Comments', HostSize)
                                    set_parameter_by_name(x, 'FP_Hanger Host Diameter', HostSize)
                            except:
                                pass

                            ItmDims = x.GetDimensions()
                            for dta in ItmDims:
                                if dta.Name == 'Length A':
                                    RLA = x.GetDimensionValue(dta)
                                if dta.Name == 'Length B':
                                    RLB = x.GetDimensionValue(dta)
                            set_parameter_by_name(x, 'FP_Rod Length', RLA)
                            set_parameter_by_name(x, 'FP_Rod Length A', RLA)
                            set_parameter_by_name(x, 'FP_Rod Length B', RLB)

                    else:
                        ItmDims = x.GetDimensions()
                        for dta in ItmDims:
                            if dta.Name == 'Length A':
                                RLA = x.GetDimensionValue(dta)
                            if dta.Name == 'Length B':
                                RLB = x.GetDimensionValue(dta)
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
                        set_parameter_by_name(x, 'FP_Rod Length', RLA)
                        set_parameter_by_name(x, 'FP_Rod Length A', RLA)
                        set_parameter_by_name(x, 'FP_Rod Length B', RLB)
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


    t = Transaction(doc, "Update FP Parameters")
    t.Start()

    for hanger in hanger_collector:
        if (hanger.GetRodInfo().RodCount) < 2:
            # Get the host element's size
            hosted_info = hanger.GetHostedInfo().HostId
            try:
                HostSize = get_parameter_value_by_name_AsString(doc.GetElement(hosted_info), 'Size').strip('"')
                # Get the hanger's size
                HangerSize = get_parameter_value_by_name_AsString(hanger, 'Product Entry')
                set_parameter_by_name(hanger, 'FP_Product Entry', HangerSize)
                set_parameter_by_name(hanger, 'FP_Hanger Host Diameter', HostSize)

                # Set the 'FP_Hanger Shield' parameter based on whether the host and hanger sizes match
                if HostSize == HangerSize:
                    set_parameter_by_name(hanger, 'FP_Hanger Shield', 'No')
                else:
                    set_parameter_by_name(hanger, 'FP_Hanger Shield', 'Yes')
            except:
                pass
            ItmDims = hanger.GetDimensions()
            for dta in ItmDims:
                if dta.Name == 'Length A':
                    RLA = hanger.GetDimensionValue(dta)
            set_parameter_by_name(hanger, 'FP_Rod Length A', RLA)

        try:
            if (hanger.GetRodInfo().RodCount) > 1:
                ItmDims = hanger.GetDimensions()
                for dta in ItmDims:
                    if dta.Name == 'Length A':
                        RLA = hanger.GetDimensionValue(dta)
                    if dta.Name == 'Length B':
                        RLB = hanger.GetDimensionValue(dta)
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
                set_parameter_by_name(hanger, 'FP_Rod Length', RLA)
                set_parameter_by_name(hanger, 'FP_Rod Length A', RLA)
                set_parameter_by_name(hanger, 'FP_Rod Length B', RLB)
        except:
            pass

        try:
            set_parameter_by_name(hanger, 'FP_Rod Attached', 'Yes') if hanger.GetRodInfo().IsAttachedToStructure else set_parameter_by_name(hanger, 'FP_Rod Attached', 'No')
            [set_parameter_by_name(hanger, 'FP_Rod Size', n.AncillaryWidthOrDiameter) for n in hanger.GetPartAncillaryUsage() if n.AncillaryWidthOrDiameter > 0]
        except:
            pass

        try:
            # galv nuts and flat washer
            mappings = {
                0.03125: {'rod': '334BL3150X', 'nut': '334BL3048X', 'washer': '334BL3048H'}, # 3/8
                0.04167: {'rod': '334BL3150Y', 'nut': '334BL3048Y', 'washer': '334BL3048I'}, # 1/2
                0.05208: {'rod': '334BL3150Z', 'nut': '334BL3048Z', 'washer': '334BL3048J'}, # 5/8
                0.06250: {'rod': '334BL3150P', 'nut': '334BL3049A', 'washer': '334BL3048K'}, # 3/4
                0.07292: {'rod': '334BL3150X', 'nut': '334BL3049B', 'washer': '334BL3048L'}, # 7/8
                0.08333: {'rod': '334BL3151B', 'nut': '334BL3049C', 'washer': '334BL3048M'}  # 1
            }

            # Set all parameters in one loop
            for n in hanger.GetPartAncillaryUsage():
                if n.AncillaryWidthOrDiameter > 0:
                    rounded_size = round(n.AncillaryWidthOrDiameter, 5)
                    if rounded_size in mappings:
                        codes = mappings[rounded_size]
                        set_parameter_by_name(hanger, 'FP_Code Rod', codes['rod'])
                        set_parameter_by_name(hanger, 'FP_Code Nut', codes['nut'])
                        set_parameter_by_name(hanger, 'FP_Code Washer', codes['washer'])
        except:
            pass


    t.Commit()
    
    # --------------ACTIVE VIEW-------------------