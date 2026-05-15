# -*- coding: utf-8 -*-

from Autodesk.Revit.UI import TaskDialog

uidoc = __revit__.ActiveUIDocument

if not uidoc:
    TaskDialog.Show('Error', "No active Revit UI document found.")

uidoc.RefreshActiveView()
TaskDialog.Show('Success', "Active view refreshed.")