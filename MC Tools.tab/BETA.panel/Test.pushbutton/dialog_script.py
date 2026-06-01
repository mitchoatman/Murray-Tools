import Autodesk
from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory, Transaction, FabricationPart, WorksharingUtils
from System.Collections.Generic import List
from Parameters.Add_SharedParameters import Shared_Params
from Parameters.Get_Set_Params import set_parameter_by_name, get_parameter_value_by_name_AsString

Shared_Params()

DB = Autodesk.Revit.DB
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
curview = doc.ActiveView

# Optional: check out elements if workshared
def checkout_elements(element_ids):
    if not doc.IsWorkshared or not element_ids:
        return
    try:
        id_list = List[DB.ElementId](element_ids)
        WorksharingUtils.CheckoutElements(doc, id_list)
    except:
        pass

# Collect fabrication ductwork in active view only
duct_collector = FilteredElementCollector(doc, curview.Id) \
    .OfCategory(BuiltInCategory.OST_FabricationDuctwork) \
    .WhereElementIsNotElementType() \
    .ToElements()

# Optional: try to check out non-editable elements
non_editable_ids = []
if doc.IsWorkshared:
    for duct in duct_collector:
        try:
            tooltip_info = WorksharingUtils.GetWorksharingTooltipInfo(doc, duct.Id)
            if not tooltip_info.Editable:
                non_editable_ids.append(duct.Id)
        except:
            pass

if non_editable_ids:
    checkout_elements(non_editable_ids)

t = Transaction(doc, "Copy Item Number to STRATUS Assembly")
t.Start()

for duct in duct_collector:
    try:
        if isinstance(duct, FabricationPart):
            item_number = get_parameter_value_by_name_AsString(duct, 'Item Number')

            if item_number:
                target_param = duct.LookupParameter('STRATUS Assembly')
                if target_param and not target_param.IsReadOnly:
                    set_parameter_by_name(duct, 'STRATUS Assembly', item_number)
    except:
        pass

t.Commit()