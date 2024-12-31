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

servicetypenamelist = []
CIDlist = []
Namelist = []
SNamelist = []
SNameABBRlist = []
STRATUSAssemlist = []
Sizelist = []
ValveNumlist = []
LineNumlist = []
STRATUSStatuslist = []
RefLevelList = []
ItemNumList = []
BundleList = []
REFLineNumList = []
REFBSDesigList = []
elementlist = []
CommentsList = []
SpecificationList = []
HangerRodSizeList = []

selection = revit.get_selection()

preselection = [doc.GetElement(id) for id in __revit__.ActiveUIDocument.Selection.GetElementIds()]

if preselection:

    try:
        for x in preselection:
            isfabpart = x.LookupParameter("Fabrication Service")
            if isfabpart:
                servicetypenamelist.append(Config.GetServiceTypeName(x.ServiceType))

                CIDlist.append(x.ItemCustomId)

                Namelist.append(get_parameter_value_by_name_AsValueString(x, 'Family'))

                SNamelist.append(get_parameter_value_by_name_AsString(x, 'Fabrication Service Name'))

                SNameABBRlist.append(get_parameter_value_by_name_AsString(x, 'Fabrication Service Abbreviation'))

                STRATUSAssemlist.append(get_parameter_value_by_name_AsString(x, 'STRATUS Assembly'))

                STRATUSStatuslist.append(get_parameter_value_by_name_AsString(x, 'STRATUS Status'))

                ValveNumlist.append(get_parameter_value_by_name_AsString(x, 'FP_Valve Number'))

                LineNumlist.append(get_parameter_value_by_name_AsString(x, 'FP_Line Number'))

                RefLevelList.append(get_parameter_value_by_name_AsValueString(x, 'Reference Level'))

                ItemNumList.append(get_parameter_value_by_name_AsString(x, 'Item Number'))

                BundleList.append(get_parameter_value_by_name_AsString(x, 'FP_Bundle'))

                REFLineNumList.append(get_parameter_value_by_name_AsString(x, 'FP_REF Line Number'))

                REFBSDesigList.append(get_parameter_value_by_name_AsString(x, 'FP_REF BS Designation'))

                Sizelist.append(get_parameter_value_by_name_AsString(x, 'Size of Primary End'))
                
                CommentsList.append(get_parameter_value_by_name_AsString(x, 'Comments'))
                
                SpecificationList.append (Config.GetSpecificationName(x.Specification))
                
                HangerRodSizeList.append (get_parameter_value_by_name_AsValueString(x, 'FP_Rod Size'))

    except:
        pass
        
    # Creating dictionaries for faster lookup
    CID_set = set(CIDlist)
    service_type_set = set(servicetypenamelist)
    Name_set = set(Namelist)
    SNamelist_set = set(SNamelist)
    SNameABBRlist_set = set(SNameABBRlist)
    STRATUSAssemlist_set = set(STRATUSAssemlist)
    LineNumlist_set = set(LineNumlist)
    STRATUSStatuslist_set = set(STRATUSStatuslist)
    RefLevelList_set = set(RefLevelList)
    ItemNumList_set = set(ItemNumList)
    BundleList_set = set(BundleList)
    REFBSDesigList_set = set(REFBSDesigList)
    REFLineNumList_set = set(REFLineNumList)
    ValveNumlist_set = set(ValveNumlist)
    CommentsList_set = set(CommentsList)
    SpecificationList_set = set(SpecificationList)
    HangerRodSizeList_set = set(HangerRodSizeList)    

    try:
        GroupOptions = {'CID': sorted(CID_set),
            'ServiceType': sorted(service_type_set),
            'Name': sorted(Name_set),
            'Service Name': sorted(SNamelist_set),
            'Service Abbreviation': sorted(SNameABBRlist_set),
            'Size': sorted(set(Sizelist)),
            'STRATUS Assembly': sorted(STRATUSAssemlist_set),
            'Line Number': sorted(LineNumlist),
            'STRATUS Status': sorted(STRATUSStatuslist_set),
            'Reference Level': sorted(RefLevelList_set),
            'Item Number': sorted(ItemNumList_set),
            'Bundle Number': sorted(BundleList_set),
            'REF BS Designation': sorted(REFBSDesigList_set),
            'REF Line Number': sorted(REFLineNumList_set),
            'Comments': sorted(CommentsList_set),
            'Specification': sorted(SpecificationList_set),
            'Hanger Rod Size': sorted(HangerRodSizeList_set),
            'Valve Number': sorted(ValveNumlist_set)}

        res = forms.SelectFromList.show(GroupOptions,group_selector_title='Property Type:', multiselect=True, button_name='Select Item(s)', exitscript = True)
        
        for fil in res:
            for elem in preselection:
                if fil in CID_set:
                    if elem.ItemCustomId == fil:
                        elementlist.append(elem.Id)
                if fil in service_type_set:
                    if Config.GetServiceTypeName(elem.ServiceType) == fil:
                        elementlist.append(elem.Id)
                if fil in Name_set:
                    if get_parameter_value_by_name_AsValueString(elem, 'Family') == fil:
                        elementlist.append(elem.Id)
                if fil in SNamelist_set:
                    if get_parameter_value_by_name_AsString(elem, 'Fabrication Service Name') == fil:
                        elementlist.append(elem.Id)
                if fil in SNameABBRlist_set:
                    if get_parameter_value_by_name_AsString(elem, 'Fabrication Service Abbreviation') == fil:
                        elementlist.append(elem.Id)
                if fil in STRATUSAssemlist_set:
                    if get_parameter_value_by_name_AsString(elem, 'STRATUS Assembly') == fil:
                        elementlist.append(elem.Id)
                if fil in STRATUSStatuslist_set:
                    if get_parameter_value_by_name_AsString(elem, 'STRATUS Status') == fil:
                        elementlist.append(elem.Id)
                if fil in LineNumlist_set:
                    if get_parameter_value_by_name_AsString(elem, 'FP_Line Number') == fil:
                        elementlist.append(elem.Id)
                if fil in RefLevelList_set:
                    if get_parameter_value_by_name_AsValueString(elem, 'Reference Level') == fil:
                        elementlist.append(elem.Id)
                if fil in ItemNumList_set:
                    if get_parameter_value_by_name_AsString(elem, 'Item Number') == fil:
                        elementlist.append(elem.Id)
                if fil in BundleList_set:
                    if get_parameter_value_by_name_AsString(elem, 'FP_Bundle') == fil:
                        elementlist.append(elem.Id)
                if fil in REFBSDesigList_set:
                    if get_parameter_value_by_name_AsString(elem, 'FP_REF BS Designation') == fil:
                        elementlist.append(elem.Id)
                if fil in REFLineNumList_set:
                    if get_parameter_value_by_name_AsString(elem, 'FP_REF Line Number') == fil:
                        elementlist.append(elem.Id)
                if fil in ValveNumlist_set:
                    if get_parameter_value_by_name_AsString(elem, 'FP_Valve Number') == fil:
                        elementlist.append(elem.Id)
                if fil in HangerRodSizeList_set:
                    if get_parameter_value_by_name_AsValueString(elem, 'FP_Rod Size') == fil:
                        elementlist.append(elem.Id)
                if get_parameter_value_by_name_AsString(elem, 'Size of Primary End') == fil:
                    elementlist.append(elem.Id)
                if get_parameter_value_by_name_AsString(elem, 'Comments') == fil:
                    elementlist.append(elem.Id)
                if (Config.GetSpecificationName(elem.Specification)) == fil:
                    elementlist.append(elem.Id)
            selection.set_to(elementlist)
    except:
        pass
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
            CIDlist = list(map(lambda x: get_parameter_value_by_name_AsInteger(x, 'Part Pattern Number'), part_collector))
            Namelist = list(map(lambda x: get_parameter_value_by_name_AsValueString(x, 'Family'), part_collector))
            SNamelist = list(map(lambda x: get_parameter_value_by_name_AsString(x, 'Fabrication Service Name'), part_collector))
            SNameABBRlist = list(map(lambda x: get_parameter_value_by_name_AsString(x, 'Fabrication Service Abbreviation'), part_collector))
            STRATUSAssemlist = list(map(lambda x: get_parameter_value_by_name_AsString(x, 'STRATUS Assembly'), part_collector))
            STRATUSStatuslist = list(map(lambda x: get_parameter_value_by_name_AsString(x, 'STRATUS Status'), part_collector))
            ValveNumlist = list(map(lambda x: get_parameter_value_by_name_AsString(x, 'FP_Valve Number'), part_collector))
            servicetypenamelist = list(map(lambda x: Config.GetServiceTypeName(x.ServiceType), part_collector))
            LineNumlist = list(map(lambda x: get_parameter_value_by_name_AsString(x, 'FP_Line Number'), part_collector))
            RefLevelList = list(map(lambda x: get_parameter_value_by_name_AsValueString(x, 'Reference Level'), part_collector))
            ItemNumList = list(map(lambda x: get_parameter_value_by_name_AsString(x, 'Item Number'), part_collector))
            BundleList = list(map(lambda x: get_parameter_value_by_name_AsString(x, 'FP_Bundle'), part_collector))
            REFLineNumList = list(map(lambda x: get_parameter_value_by_name_AsString(x, 'FP_REF Line Number'), part_collector))
            REFBSDesigList = list(map(lambda x: get_parameter_value_by_name_AsString(x, 'FP_REF BS Designation'), part_collector))
            Sizelist = list(map(lambda x: get_parameter_value_by_name_AsString(x, 'Size'), pipeduct_collector))
            CommentsList = list(map(lambda x: get_parameter_value_by_name_AsString(x, 'Comments'), part_collector))
            SpecificationList = list(map(lambda x: Config.GetSpecificationName(x.Specification), part_collector))
            HangerRodSizeList = list(map(lambda x: get_parameter_value_by_name_AsValueString(x, 'FP_Rod Size'), hanger_collector))

        except:
            print 'No Fabrication Parts in View'
        try:
            # Size filter only
            for elem in hanger_collector:
                    Sizelist.append(get_parameter_value_by_name_AsString(elem, 'Size of Primary End'))
        except:
            pass

        # Creating dictionaries for faster lookup
        CID_set = set(CIDlist)
        service_type_set = set(servicetypenamelist)
        Name_set = set(Namelist)
        SNamelist_set = set(SNamelist)
        SNameABBRlist_set = set(SNameABBRlist)
        STRATUSAssemlist_set = set(STRATUSAssemlist)
        LineNumlist_set = set(LineNumlist)
        STRATUSStatuslist_set = set(STRATUSStatuslist)
        RefLevelList_set = set(RefLevelList)
        ItemNumList_set = set(ItemNumList)
        BundleList_set = set(BundleList)
        REFBSDesigList_set = set(REFBSDesigList)
        REFLineNumList_set = set(REFLineNumList)
        ValveNumlist_set = set(ValveNumlist)
        Sizelist_set = set(Sizelist)
        SpecificationList_set = set(SpecificationList)
        CommentsList_set = set(CommentsList)
        HangerRodSizeList_set = set(HangerRodSizeList)  

        try:
            GroupOptions = {'CID': sorted(CID_set),
                'ServiceType': sorted(service_type_set),
                'Name': sorted(Name_set),
                'Service Name': sorted(SNamelist_set),
                'Service Abbreviation': sorted(SNameABBRlist_set),
                'Size': sorted(Sizelist_set),
                'STRATUS Assembly': sorted(STRATUSAssemlist_set),
                'Line Number': sorted(LineNumlist_set),
                'STRATUS Status': sorted(STRATUSStatuslist_set),
                'Reference Level': sorted(RefLevelList_set),
                'Item Number': sorted(ItemNumList_set),
                'Bundle Number': sorted(BundleList_set),
                'REF BS Designation': sorted(REFBSDesigList_set),
                'REF Line Number': sorted(REFLineNumList_set),
                'Comments': sorted(CommentsList_set),
                'Specification': sorted(SpecificationList_set),
                'Hanger Rod Size': sorted(HangerRodSizeList_set),
                'Valve Number': sorted(ValveNumlist_set)}

            res = forms.SelectFromList.show(GroupOptions,group_selector_title='Property Type:', multiselect=True, button_name='Select Item(s)', exitscript = True)

            for fil in res:
                for elem in part_collector:
                    if fil in CID_set:
                        if elem.ItemCustomId == fil:
                            elementlist.append(elem.Id)
                    if fil in service_type_set:
                        if Config.GetServiceTypeName(elem.ServiceType) == fil:
                            elementlist.append(elem.Id)
                    if fil in Name_set:
                        if get_parameter_value_by_name_AsValueString(elem, 'Family') == fil:
                            elementlist.append(elem.Id)
                    if fil in SNamelist_set:
                        if get_parameter_value_by_name_AsString(elem, 'Fabrication Service Name') == fil:
                            elementlist.append(elem.Id)
                    if fil in SNameABBRlist_set:
                        if get_parameter_value_by_name_AsString(elem, 'Fabrication Service Abbreviation') == fil:
                            elementlist.append(elem.Id)
                    if fil in STRATUSAssemlist_set:
                        if get_parameter_value_by_name_AsString(elem, 'STRATUS Assembly') == fil:
                            elementlist.append(elem.Id)
                    if fil in STRATUSStatuslist_set:
                        if get_parameter_value_by_name_AsString(elem, 'STRATUS Status') == fil:
                            elementlist.append(elem.Id)
                    if fil in LineNumlist_set:
                        if get_parameter_value_by_name_AsString(elem, 'FP_Line Number') == fil:
                            elementlist.append(elem.Id)
                    if fil in RefLevelList_set:
                        if get_parameter_value_by_name_AsValueString(elem, 'Reference Level') == fil:
                            elementlist.append(elem.Id)
                    if fil in ItemNumList_set:
                        if get_parameter_value_by_name_AsString(elem, 'Item Number') == fil:
                            elementlist.append(elem.Id)
                    if fil in BundleList_set:
                        if get_parameter_value_by_name_AsString(elem, 'FP_Bundle') == fil:
                            elementlist.append(elem.Id)
                    if fil in REFBSDesigList_set:
                        if get_parameter_value_by_name_AsString(elem, 'FP_REF BS Designation') == fil:
                            elementlist.append(elem.Id)
                    if fil in REFLineNumList_set:
                        if get_parameter_value_by_name_AsString(elem, 'FP_REF Line Number') == fil:
                            elementlist.append(elem.Id)
                    if fil in ValveNumlist_set:
                        if get_parameter_value_by_name_AsString(elem, 'FP_Valve Number') == fil:
                            elementlist.append(elem.Id)
                    if fil in CommentsList_set:
                        if get_parameter_value_by_name_AsString(elem, 'Comments') == fil:
                            elementlist.append(elem.Id)
                    if fil in HangerRodSizeList_set:
                        if get_parameter_value_by_name_AsValueString(elem, 'FP_Rod Size') == fil:
                            elementlist.append(elem.Id)
                    if (Config.GetSpecificationName(elem.Specification)) == fil:
                        elementlist.append(elem.Id)
                for elem in pipeduct_collector:
                    if get_parameter_value_by_name_AsString(elem, 'Size') == fil:
                        elementlist.append(elem.Id)
                for elem in hanger_collector:
                    if get_parameter_value_by_name_AsString(elem, 'Size of Primary End') == fil:
                        elementlist.append(elem.Id)
            selection.set_to(elementlist)
        except:
            pass
    else:
        forms.alert('No Fabrication parts in view.', exitscript=True)

