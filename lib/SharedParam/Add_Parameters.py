import os
from Autodesk.Revit.DB import BuiltInCategory, Transaction, TransactionGroup

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
app = doc.Application
RevitVersion = app.VersionNumber
RevitINT = float (RevitVersion)

def Shared_Params():
    path, filename = os.path.split(__file__)
    NewFilename = 'MC Shared Parameters.txt'
    fullPath = os.path.join(path, NewFilename)

    sel = uidoc.Selection
    cat1 = doc.Settings.Categories.get_Item(BuiltInCategory.OST_FabricationPipework)
    cat2 = doc.Settings.Categories.get_Item(BuiltInCategory.OST_FabricationHangers)
    cat3 = doc.Settings.Categories.get_Item(BuiltInCategory.OST_FabricationDuctwork)
    FPcatSet = app.Create.NewCategorySet()
    FPcatSet.Insert(cat1)
    FPcatSet.Insert(cat2)
    FPcatSet.Insert(cat3)

    cat4 = doc.Settings.Categories.get_Item(BuiltInCategory.OST_PlumbingFixtures)
    cat5 = doc.Settings.Categories.get_Item(BuiltInCategory.OST_PipeAccessory)
    cat6 = doc.Settings.Categories.get_Item(BuiltInCategory.OST_GenericModel)
    cat7 = doc.Settings.Categories.get_Item(BuiltInCategory.OST_StructuralFraming)
    cat8 = doc.Settings.Categories.get_Item(BuiltInCategory.OST_PipeFitting)
    cat9 = doc.Settings.Categories.get_Item(BuiltInCategory.OST_DuctAccessory)
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

    t = Transaction(doc, 'Add Parameters')
    t.Start()    

    if RevitINT > 2024:
        from Autodesk.Revit.DB import GroupTypeId

        app.SharedParametersFilename = fullPath

        spFile = app.OpenSharedParameterFile()
        for dG in spFile.Groups:
            if dG.Name == 'FP Parameters':
                d = dG.Definitions
                for eD in d:
                    # Use GroupTypeId.INVALID or appropriate GroupTypeId
                    newIB = app.Create.NewInstanceBinding(STRATUScatSet)
                    doc.ParameterBindings.Insert(eD, newIB, GroupTypeId.General)  # Adjust as needed

        spFile = app.OpenSharedParameterFile()
        for dG in spFile.Groups:
            if dG.Name == 'STRATUS Parameters':
                d = dG.Definitions
                for eD in d:
                    newIB = app.Create.NewInstanceBinding(STRATUScatSet)
                    doc.ParameterBindings.Insert(eD, newIB, GroupTypeId.General)  # Adjust as needed


        spFile = app.OpenSharedParameterFile()
        for dG in spFile.Groups:
            if dG.Name == 'MC_General Data':
                d = dG.Definitions
                for eD in d:
                    newIB = app.Create.NewInstanceBinding(STRATUScatSet)
                    doc.ParameterBindings.Insert(eD, newIB, GroupTypeId.General)  # Adjust as needed
    else:
        from Autodesk.Revit.DB import BuiltInParameterGroup

        app.SharedParametersFilename = fullPath

        spFile = app.OpenSharedParameterFile()
        for dG in spFile.Groups:
            if dG.Name == 'FP Parameters':
                d = dG.Definitions
                for eD in d:
                    # Use GroupTypeId.INVALID or appropriate GroupTypeId
                    newIB = app.Create.NewInstanceBinding(STRATUScatSet)
                    doc.ParameterBindings.Insert(eD, newIB, BuiltInParameterGroup.INVALID)  # Adjust as needed

        spFile = app.OpenSharedParameterFile()
        for dG in spFile.Groups:
            if dG.Name == 'STRATUS Parameters':
                d = dG.Definitions
                for eD in d:
                    newIB = app.Create.NewInstanceBinding(STRATUScatSet)
                    doc.ParameterBindings.Insert(eD, newIB, BuiltInParameterGroup.INVALID)  # Adjust as needed


        spFile = app.OpenSharedParameterFile()
        for dG in spFile.Groups:
            if dG.Name == 'MC_General Data':
                d = dG.Definitions
                for eD in d:
                    newIB = app.Create.NewInstanceBinding(STRATUScatSet)
                    doc.ParameterBindings.Insert(eD, newIB, BuiltInParameterGroup.INVALID)  # Adjust as needed
    t.Commit()

