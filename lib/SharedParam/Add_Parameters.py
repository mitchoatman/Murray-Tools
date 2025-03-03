import os
from Autodesk.Revit.DB import BuiltInCategory, Transaction, TransactionGroup

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
app = doc.Application
RevitVersion = app.VersionNumber
RevitINT = float(RevitVersion)

def Shared_Params():
    path, filename = os.path.split(__file__)
    NewFilename = 'MC Shared Parameters.txt'
    fullPath = os.path.join(path, NewFilename)

    cat1 = doc.Settings.Categories.get_Item(BuiltInCategory.OST_FabricationPipework)
    cat2 = doc.Settings.Categories.get_Item(BuiltInCategory.OST_FabricationHangers)
    cat3 = doc.Settings.Categories.get_Item(BuiltInCategory.OST_FabricationDuctwork)
    cat4 = doc.Settings.Categories.get_Item(BuiltInCategory.OST_PlumbingFixtures)
    cat5 = doc.Settings.Categories.get_Item(BuiltInCategory.OST_PipeAccessory)
    cat6 = doc.Settings.Categories.get_Item(BuiltInCategory.OST_GenericModel)
    cat7 = doc.Settings.Categories.get_Item(BuiltInCategory.OST_StructuralFraming)
    cat8 = doc.Settings.Categories.get_Item(BuiltInCategory.OST_PipeFitting)
    cat9 = doc.Settings.Categories.get_Item(BuiltInCategory.OST_DuctAccessory)
    cat10 = doc.Settings.Categories.get_Item(BuiltInCategory.OST_FlexDuctCurves)

    FPcatSet = app.Create.NewCategorySet()
    FPcatSet.Insert(cat1)
    FPcatSet.Insert(cat2)
    FPcatSet.Insert(cat3)

    STRATUScatSet = app.Create.NewCategorySet()
    STRATUScatSet.Insert(cat1)
    STRATUScatSet.Insert(cat2)
    STRATUScatSet.Insert(cat3)
    STRATUScatSet.Insert(cat4)
    STRATUScatSet.Insert(cat5)
    STRATUScatSet.Insert(cat6)
    STRATUScatSet.Insert(cat7)
    STRATUScatSet.Insert(cat8)
    STRATUScatSet.Insert(cat9)
    STRATUScatSet.Insert(cat10)

    t = Transaction(doc, 'Add Parameters')
    t.Start()    

    app.SharedParametersFilename = fullPath
    spFile = app.OpenSharedParameterFile()

    def add_parameters(group_name, category_set, group_type):
        for dG in spFile.Groups:
            if dG.Name == group_name:
                definitions = sorted(dG.Definitions, key=lambda x: x.Name)  # Sort alphabetically by name
                for eD in definitions:
                    newIB = app.Create.NewInstanceBinding(category_set)
                    doc.ParameterBindings.Insert(eD, newIB, group_type)

    if RevitINT > 2024:
        from Autodesk.Revit.DB import GroupTypeId

        add_parameters('FP Parameters', STRATUScatSet, GroupTypeId.General)
        add_parameters('STRATUS Parameters', STRATUScatSet, GroupTypeId.General)
        add_parameters('MC_General Data', STRATUScatSet, GroupTypeId.General)

    else:
        from Autodesk.Revit.DB import BuiltInParameterGroup

        add_parameters('FP Parameters', STRATUScatSet, BuiltInParameterGroup.INVALID)
        add_parameters('STRATUS Parameters', STRATUScatSet, BuiltInParameterGroup.INVALID)
        add_parameters('MC_General Data', STRATUScatSet, BuiltInParameterGroup.INVALID)

    t.Commit()

