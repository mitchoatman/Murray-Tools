import System
from Autodesk.Revit.DB import BuiltInCategory, Transaction, BuiltInParameterGroup, TransactionGroup
import os

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
app = doc.Application

def Shared_Params():
    path, filename = os.path.split(__file__)
    NewFilename = '\MC Shared Parameters.txt'

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
    STRATUScatSet = app.Create.NewCategorySet()
    STRATUScatSet.Insert(cat1)
    STRATUScatSet.Insert(cat2)
    STRATUScatSet.Insert(cat3)
    STRATUScatSet.Insert(cat4)
    STRATUScatSet.Insert(cat5)
    STRATUScatSet.Insert(cat6)
    STRATUScatSet.Insert(cat7)
    STRATUScatSet.Insert(cat8)
    
    tg = TransactionGroup(doc, "Add Parameters")
    tg.Start()

    t = Transaction(doc, 'Change SParam File')
    t.Start()	
    app.SharedParametersFilename = path + NewFilename
    t.Commit()

    spFile = app.OpenSharedParameterFile()
    for dG in spFile.Groups:
        if (dG.Name == 'FP Parameters'):
            d = dG.Definitions
            t = Transaction(doc, 'FP Parameters')
            t.Start()					
            for eD in d:
                if eD.Name == 'FP_Service Name' or 'FP_Valve Number' or 'FP_Line Number' or 'FP_Bundle':
                    newIB = app.Create.NewInstanceBinding(STRATUScatSet)
                    doc.ParameterBindings.Insert(eD,newIB,BuiltInParameterGroup.INVALID)   #BuiltInParameterGroup.INVALID for other   BuiltInParameterGroup.PG_IDENTITY_DATA for Identity Data
                else:
                    newIB = app.Create.NewInstanceBinding(FPcatSet)
                    doc.ParameterBindings.Insert(eD,newIB,BuiltInParameterGroup.INVALID)   #BuiltInParameterGroup.INVALID for other   BuiltInParameterGroup.PG_IDENTITY_DATA for Identity Data

            t.Commit()

    spFile = app.OpenSharedParameterFile()
    for dG in spFile.Groups:
        if (dG.Name == 'STRATUS Parameters'):
            d = dG.Definitions
            t = Transaction(doc, 'Add STRATUS Parameters')
            t.Start()					
            for eD in d:
                newIB = app.Create.NewInstanceBinding(STRATUScatSet)
                doc.ParameterBindings.Insert(eD,newIB,BuiltInParameterGroup.INVALID)   #BuiltInParameterGroup.INVALID for other   BuiltInParameterGroup.PG_IDENTITY_DATA for Identity Data

            t.Commit()

    spFile = app.OpenSharedParameterFile()
    for dG in spFile.Groups:
        if (dG.Name == 'Identity Data'):
            d = dG.Definitions
            t = Transaction(doc, 'Add TS Parameters')
            t.Start()					
            for eD in d:
                newIB = app.Create.NewInstanceBinding(STRATUScatSet)
                doc.ParameterBindings.Insert(eD,newIB,BuiltInParameterGroup.INVALID)   #BuiltInParameterGroup.INVALID for other   BuiltInParameterGroup.PG_IDENTITY_DATA for Identity Data

            t.Commit()
    #End Transaction Group
    tg.Assimilate()








