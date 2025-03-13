import Autodesk
from Autodesk.Revit import DB
from Autodesk.Revit.DB import Transaction, FilteredElementCollector, BuiltInCategory, FabricationConfiguration, FabricationPart
from pyrevit import revit, DB, forms
from SharedParam.Add_Parameters import Shared_Params
from Parameters.Get_Set_Params import get_parameter_value_by_name_AsString, get_parameter_value_by_name_AsValueString, get_parameter_value_by_name_AsInteger

Shared_Params()

#define the active Revit application and document
DB = Autodesk.Revit.DB
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
curview = doc.ActiveView
fec = FilteredElementCollector
app = doc.Application
RevitVersion = app.VersionNumber
RevitINT = float (RevitVersion)
Config = FabricationConfiguration.GetFabricationConfiguration(doc)

servicetypenameList = []
CIDList = []
NameList = []
SNameList = []
SNameABBRList = []
STRATUSAssemList = []
SizeList = []
ValveNumList = []
LineNumList = []
STRATUSStatusList = []
RefLevelList = []
ItemNumList = []
BundleList = []
REFLineNumList = []
REFBSDesigList = []
elementList = []
CommentsList = []
SpecificationList = []
HangerRodSizeList = []
BeamHangerList = []

selection = revit.get_selection()

preselection = [doc.GetElement(id) for id in __revit__.ActiveUIDocument.Selection.GetElementIds()]

if preselection:

    try:
        for x in preselection:
            isfabpart = x.LookupParameter("Fabrication Service")
            if isfabpart:
                servicetypenameList.append(Config.GetServiceTypeName(x.ServiceType))
                CIDList.append(x.ItemCustomId)
                NameList.append(get_parameter_value_by_name_AsValueString(x, 'Family'))
                SNameList.append(get_parameter_value_by_name_AsString(x, 'Fabrication Service Name'))
                SNameABBRList.append(get_parameter_value_by_name_AsString(x, 'Fabrication Service Abbreviation'))
                STRATUSAssemList.append(get_parameter_value_by_name_AsString(x, 'STRATUS Assembly'))
                STRATUSStatusList.append(get_parameter_value_by_name_AsString(x, 'STRATUS Status'))
                ValveNumList.append(get_parameter_value_by_name_AsString(x, 'FP_Valve Number'))
                LineNumList.append(get_parameter_value_by_name_AsString(x, 'FP_Line Number'))
                RefLevelList.append(get_parameter_value_by_name_AsValueString(x, 'Reference Level'))
                ItemNumList.append(get_parameter_value_by_name_AsString(x, 'Item Number'))
                BundleList.append(get_parameter_value_by_name_AsString(x, 'FP_Bundle'))
                REFLineNumList.append(get_parameter_value_by_name_AsString(x, 'FP_REF Line Number'))
                REFBSDesigList.append(get_parameter_value_by_name_AsString(x, 'FP_REF BS Designation'))
                SizeList.append(get_parameter_value_by_name_AsString(x, 'Size of Primary End'))
                CommentsList.append(get_parameter_value_by_name_AsString(x, 'Comments'))
                SpecificationList.append (Config.GetSpecificationName(x.Specification))
                HangerRodSizeList.append (get_parameter_value_by_name_AsValueString(x, 'FP_Rod Size'))
                BeamHangerList.append (get_parameter_value_by_name_AsString(x, 'FP_Beam Hanger'))

    except:
        pass
        
    # Creating dictionaries for faster lookup
    CID_set = set(CIDList)
    service_type_set = set(servicetypenameList)
    Name_set = set(NameList)
    SNameList_set = set(SNameList)
    SNameABBRList_set = set(SNameABBRList)
    STRATUSAssemList_set = set(STRATUSAssemList)
    LineNumList_set = set(LineNumList)
    STRATUSStatusList_set = set(STRATUSStatusList)
    RefLevelList_set = set(RefLevelList)
    ItemNumList_set = set(ItemNumList)
    BundleList_set = set(BundleList)
    REFBSDesigList_set = set(REFBSDesigList)
    REFLineNumList_set = set(REFLineNumList)
    ValveNumList_set = set(ValveNumList)
    CommentsList_set = set(CommentsList)
    SpecificationList_set = set(SpecificationList)
    HangerRodSizeList_set = set(HangerRodSizeList)    
    BeamHangerList_set = set(BeamHangerList) 

    GroupOptions = {'CID': sorted(CID_set),
        'ServiceType': sorted(service_type_set),
        'Name': sorted(Name_set),
        'Service Name': sorted(SNameList_set),
        'Service Abbreviation': sorted(SNameABBRList_set),
        'Size': sorted(set(SizeList)),
        'STRATUS Assembly': sorted(STRATUSAssemList_set),
        'Line Number': sorted(LineNumList),
        'STRATUS Status': sorted(STRATUSStatusList_set),
        'Reference Level': sorted(RefLevelList_set),
        'Item Number': sorted(ItemNumList_set),
        'Bundle Number': sorted(BundleList_set),
        'REF BS Designation': sorted(REFBSDesigList_set),
        'REF Line Number': sorted(REFLineNumList_set),
        'Comments': sorted(CommentsList_set),
        'Specification': sorted(SpecificationList_set),
        'Hanger Rod Size': sorted(HangerRodSizeList_set),
        'Valve Number': sorted(ValveNumList_set),
        'Beam Hanger': sorted(BeamHangerList_set)}

    res = forms.SelectFromList.show(GroupOptions,group_selector_title='Property Type:', multiselect=True, button_name='Select Item(s)', exitscript = True)
    
    if res:  # Check if user selected any filters
        elementList = []  # Reset the list
        processed_ids = set()  # To avoid duplicates
        try:        
            for elem in preselection:
                element_added = False
                for fil in res:
                    # Check each property and add element ID if it matches
                    if (fil in CID_set and elem.ItemCustomId == fil) or \
                       (fil in service_type_set and Config.GetServiceTypeName(elem.ServiceType) == fil) or \
                       (fil in Name_set and get_parameter_value_by_name_AsValueString(elem, 'Family') == fil) or \
                       (fil in SNameList_set and get_parameter_value_by_name_AsString(elem, 'Fabrication Service Name') == fil) or \
                       (fil in SNameABBRList_set and get_parameter_value_by_name_AsString(elem, 'Fabrication Service Abbreviation') == fil) or \
                       (fil in STRATUSAssemList_set and get_parameter_value_by_name_AsString(elem, 'STRATUS Assembly') == fil) or \
                       (fil in STRATUSStatusList_set and get_parameter_value_by_name_AsString(elem, 'STRATUS Status') == fil) or \
                       (fil in LineNumList_set and get_parameter_value_by_name_AsString(elem, 'FP_Line Number') == fil) or \
                       (fil in RefLevelList_set and get_parameter_value_by_name_AsValueString(elem, 'Reference Level') == fil) or \
                       (fil in ItemNumList_set and get_parameter_value_by_name_AsString(elem, 'Item Number') == fil) or \
                       (fil in BundleList_set and get_parameter_value_by_name_AsString(elem, 'FP_Bundle') == fil) or \
                       (fil in REFBSDesigList_set and get_parameter_value_by_name_AsString(elem, 'FP_REF BS Designation') == fil) or \
                       (fil in REFLineNumList_set and get_parameter_value_by_name_AsString(elem, 'FP_REF Line Number') == fil) or \
                       (fil in ValveNumList_set and get_parameter_value_by_name_AsString(elem, 'FP_Valve Number') == fil) or \
                       (fil in BeamHangerList_set and get_parameter_value_by_name_AsString(elem, 'FP_Beam Hanger') == fil) or \
                       (fil in HangerRodSizeList_set and get_parameter_value_by_name_AsValueString(elem, 'FP_Rod Size') == fil) or \
                       (get_parameter_value_by_name_AsString(elem, 'Size of Primary End') == fil) or \
                       (get_parameter_value_by_name_AsString(elem, 'Comments') == fil) or \
                       (Config.GetSpecificationName(elem.Specification) == fil):
                        if elem.Id not in processed_ids:
                            elementList.append(elem.Id)
                            processed_ids.add(elem.Id)
                        element_added = True
                        break  # Move to next element if we found a match
        except:
            pass            
        if elementList:
            selection.set_to(elementList)
        else:
            print("No elements matched the selected filters")
    else:
        print("No filters selected")
else:
    # Create a FilteredElementCollector to get all FabricationPart elements
    part_collector = FilteredElementCollector(doc, curview.Id).OfClass(FabricationPart) \
                       .WhereElementIsNotElementType() \
                       .ToElements()
    # collector for size only
    hanger_collector = FilteredElementCollector(doc, curview.Id).OfCategory(BuiltInCategory.OST_FabricationHangers) \
                       .WhereElementIsNotElementType() \
                       .ToElements()
    pipeduct_collector = FilteredElementCollector(doc, curview.Id).OfCategory(BuiltInCategory.OST_FabricationPipework) \
                       .UnionWith(FilteredElementCollector(doc, curview.Id).OfCategory(BuiltInCategory.OST_FabricationDuctwork)) \
                       .WhereElementIsNotElementType() \
                       .ToElements()

    if part_collector:
        try:
            CIDList = list(map(lambda x: get_parameter_value_by_name_AsInteger(x, 'Part Pattern Number'), part_collector))
            NameList = list(map(lambda x: get_parameter_value_by_name_AsValueString(x, 'Family'), part_collector))
            SNameList = list(map(lambda x: get_parameter_value_by_name_AsString(x, 'Fabrication Service Name'), part_collector))
            SNameABBRList = list(map(lambda x: get_parameter_value_by_name_AsString(x, 'Fabrication Service Abbreviation'), part_collector))
            STRATUSAssemList = list(map(lambda x: get_parameter_value_by_name_AsString(x, 'STRATUS Assembly'), part_collector))
            STRATUSStatusList = list(map(lambda x: get_parameter_value_by_name_AsString(x, 'STRATUS Status'), part_collector))
            ValveNumList = list(map(lambda x: get_parameter_value_by_name_AsString(x, 'FP_Valve Number'), part_collector))
            servicetypenameList = list(map(lambda x: Config.GetServiceTypeName(x.ServiceType), part_collector))
            LineNumList = list(map(lambda x: get_parameter_value_by_name_AsString(x, 'FP_Line Number'), part_collector))
            RefLevelList = list(map(lambda x: get_parameter_value_by_name_AsValueString(x, 'Reference Level'), part_collector))
            ItemNumList = list(map(lambda x: get_parameter_value_by_name_AsString(x, 'Item Number'), part_collector))
            BundleList = list(map(lambda x: get_parameter_value_by_name_AsString(x, 'FP_Bundle'), part_collector))
            REFLineNumList = list(map(lambda x: get_parameter_value_by_name_AsString(x, 'FP_REF Line Number'), part_collector))
            REFBSDesigList = list(map(lambda x: get_parameter_value_by_name_AsString(x, 'FP_REF BS Designation'), part_collector))
            SizeList = list(map(lambda x: get_parameter_value_by_name_AsString(x, 'Size'), pipeduct_collector))
            CommentsList = list(map(lambda x: get_parameter_value_by_name_AsString(x, 'Comments'), part_collector))
            SpecificationList = list(map(lambda x: Config.GetSpecificationName(x.Specification), part_collector))
            HangerRodSizeList = list(map(lambda x: get_parameter_value_by_name_AsValueString(x, 'FP_Rod Size'), hanger_collector))
            BeamHangerList = list(map(lambda x: get_parameter_value_by_name_AsString(x, 'FP_Beam Hanger'), hanger_collector))

        except:
            print('No Fabrication Parts in View')
        try:
            # Size filter only
            for elem in hanger_collector:
                    SizeList.append(get_parameter_value_by_name_AsString(elem, 'Size of Primary End'))
        except:
            pass

        # Creating dictionaries for faster lookup
        CID_set = set(CIDList)
        service_type_set = set(servicetypenameList)
        Name_set = set(NameList)
        SNameList_set = set(SNameList)
        SNameABBRList_set = set(SNameABBRList)
        STRATUSAssemList_set = set(STRATUSAssemList)
        LineNumList_set = set(LineNumList)
        STRATUSStatusList_set = set(STRATUSStatusList)
        RefLevelList_set = set(RefLevelList)
        ItemNumList_set = set(ItemNumList)
        BundleList_set = set(BundleList)
        REFBSDesigList_set = set(REFBSDesigList)
        REFLineNumList_set = set(REFLineNumList)
        ValveNumList_set = set(ValveNumList)
        SizeList_set = set(SizeList)
        SpecificationList_set = set(SpecificationList)
        CommentsList_set = set(CommentsList)
        HangerRodSizeList_set = set(HangerRodSizeList)  
        BeamHangerList_set = set(BeamHangerList) 

        GroupOptions = {'CID': sorted(CID_set),
            'ServiceType': sorted(service_type_set),
            'Name': sorted(Name_set),
            'Service Name': sorted(SNameList_set),
            'Service Abbreviation': sorted(SNameABBRList_set),
            'Size': sorted(SizeList_set),
            'STRATUS Assembly': sorted(STRATUSAssemList_set),
            'Line Number': sorted(LineNumList_set),
            'STRATUS Status': sorted(STRATUSStatusList_set),
            'Reference Level': sorted(RefLevelList_set),
            'Item Number': sorted(ItemNumList_set),
            'Bundle Number': sorted(BundleList_set),
            'REF BS Designation': sorted(REFBSDesigList_set),
            'REF Line Number': sorted(REFLineNumList_set),
            'Comments': sorted(CommentsList_set),
            'Specification': sorted(SpecificationList_set),
            'Hanger Rod Size': sorted(HangerRodSizeList_set),
            'Valve Number': sorted(ValveNumList_set),
            'Beam Hanger': sorted(BeamHangerList_set)}

        res = forms.SelectFromList.show(GroupOptions,group_selector_title='Property Type:', multiselect=True, button_name='Select Item(s)', exitscript = True)

        if res:  # Check if user selected any filters
            elementList = []
            processed_ids = set()  # To avoid duplicates

            # Combine all collectors into one iterable for efficiency
            all_elements = list(part_collector) + list(pipeduct_collector) + list(hanger_collector)

            for elem in all_elements:
                for fil in res:
                    # Determine element category for size parameter
                    is_hanger = elem.Category.Id.IntegerValue == BuiltInCategory.OST_FabricationHangers.value__
                    is_pipeduct = (elem.Category.Id.IntegerValue == BuiltInCategory.OST_FabricationPipework.value__ or 
                                  elem.Category.Id.IntegerValue == BuiltInCategory.OST_FabricationDuctwork.value__)

                    # Check all properties in a single condition
                    if (fil in CID_set and hasattr(elem, 'ItemCustomId') and elem.ItemCustomId == fil) or \
                       (fil in service_type_set and hasattr(elem, 'ServiceType') and Config.GetServiceTypeName(elem.ServiceType) == fil) or \
                       (fil in Name_set and get_parameter_value_by_name_AsValueString(elem, 'Family') == fil) or \
                       (fil in SNameList_set and get_parameter_value_by_name_AsString(elem, 'Fabrication Service Name') == fil) or \
                       (fil in SNameABBRList_set and get_parameter_value_by_name_AsString(elem, 'Fabrication Service Abbreviation') == fil) or \
                       (fil in STRATUSAssemList_set and get_parameter_value_by_name_AsString(elem, 'STRATUS Assembly') == fil) or \
                       (fil in STRATUSStatusList_set and get_parameter_value_by_name_AsString(elem, 'STRATUS Status') == fil) or \
                       (fil in LineNumList_set and get_parameter_value_by_name_AsString(elem, 'FP_Line Number') == fil) or \
                       (fil in RefLevelList_set and get_parameter_value_by_name_AsValueString(elem, 'Reference Level') == fil) or \
                       (fil in ItemNumList_set and get_parameter_value_by_name_AsString(elem, 'Item Number') == fil) or \
                       (fil in BundleList_set and get_parameter_value_by_name_AsString(elem, 'FP_Bundle') == fil) or \
                       (fil in REFBSDesigList_set and get_parameter_value_by_name_AsString(elem, 'FP_REF BS Designation') == fil) or \
                       (fil in REFLineNumList_set and get_parameter_value_by_name_AsString(elem, 'FP_REF Line Number') == fil) or \
                       (fil in ValveNumList_set and get_parameter_value_by_name_AsString(elem, 'FP_Valve Number') == fil) or \
                       (fil in BeamHangerList_set and get_parameter_value_by_name_AsString(elem, 'FP_Beam Hanger') == fil) or \
                       (fil in HangerRodSizeList_set and get_parameter_value_by_name_AsValueString(elem, 'FP_Rod Size') == fil) or \
                       (fil in CommentsList_set and get_parameter_value_by_name_AsString(elem, 'Comments') == fil) or \
                       (fil in SpecificationList_set and hasattr(elem, 'Specification') and Config.GetSpecificationName(elem.Specification) == fil) or \
                       (fil in SizeList_set and is_pipeduct and get_parameter_value_by_name_AsString(elem, 'Size') == fil) or \
                       (fil in SizeList_set and is_hanger and get_parameter_value_by_name_AsString(elem, 'Size of Primary End') == fil):
                        if elem.Id not in processed_ids:
                            elementList.append(elem.Id)
                            processed_ids.add(elem.Id)
                        break  # Move to next element once we find a match

            if elementList:
                selection.set_to(elementList)
            else:
                print("No elements matched the selected filters")
        else:
            print("No filters selected")