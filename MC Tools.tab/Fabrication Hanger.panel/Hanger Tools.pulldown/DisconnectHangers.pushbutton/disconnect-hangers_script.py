# -*- coding: UTF-8 -*-
import os
import clr
clr.AddReference("System.Windows.Forms")
clr.AddReference("System.Drawing")
from System.Windows.Forms import NotifyIcon, ToolTipIcon
from System.Drawing import Icon, SystemIcons

from Autodesk.Revit.DB import Transaction, FabricationPart
from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType
from Autodesk.Revit.UI import TaskDialog

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument


def show_balloon_notification(title, message, icon_path=None, timeout=5000):
    """Displays a native Windows balloon notification using a custom .ico file if available."""
    notify_icon = NotifyIcon()
    try:
        if icon_path and os.path.exists(icon_path):
            notify_icon.Icon = Icon(icon_path)
        else:
            notify_icon.Icon = SystemIcons.Information
            
        notify_icon.Visible = True
        notify_icon.ShowBalloonTip(timeout, title, message, ToolTipIcon.Info)
    except Exception:
        pass


# Custom selection filter for MEP Fabrication Hangers
class CustomISelectionFilter(ISelectionFilter):
    def __init__(self, allowed_categories):
        self.allowed_categories = allowed_categories
    def AllowElement(self, e):
        if e.Category and e.Category.Name in self.allowed_categories:
            # Additional check to ensure it's a hanger
            if isinstance(e, FabricationPart) and e.IsAHanger():
                return True
        return False
    def AllowReference(self, ref, point):
        return True


# Define the category for hangers
allowed_categories = ["MEP Fabrication Hangers"]

# Try to get user selection with error handling for cancellation
try:
    selection_filter = CustomISelectionFilter(allowed_categories)
    selected_parts = uidoc.Selection.PickObjects(ObjectType.Element, selection_filter, "Select Fabrication Hangers to disconnect from hosts")
    fab_hangers = [doc.GetElement(elId.ElementId) if hasattr(elId, "ElementId") else doc.GetElement(elId) for elId in selected_parts]

    # Start a transaction to modify elements only if selection was successful
    if fab_hangers:  # Check if list is not empty
        t = Transaction(doc, "Disconnect Hangers from Hosts")
        t.Start()

        successful_disconnects = 0
        for hanger in fab_hangers:
            try:
                # Verify it's a hanger
                if hanger.IsAHanger():
                    # Get the hosted info for the hanger
                    hosted_info = hanger.GetHostedInfo()
                    if hosted_info:
                        # Disconnect using the hosted info
                        hosted_info.DisconnectFromHost()
                        successful_disconnects += 1
                    else:
                        TaskDialog.Show("Error", "No host information found for hanger {}".format(hanger.Id))
                else:
                    TaskDialog.Show("Error", "Element {} is not a hanger".format(hanger.Id))
                    
            except Exception as e:
                TaskDialog.Show("Error", "Failed to disconnect hanger {}: {}".format(hanger.Id, str(e)))

        t.Commit()
        
        # Optional: Set a custom icon path if desired, e.g., r"C:\path\to\Murray.ico"
        icon_file = None
        
        # Show native Windows balloon notification instead of final TaskDialog
        show_balloon_notification(
            "Hanger Disconnection Complete", 
            "Successfully disconnected {} hangers.".format(successful_disconnects),
            icon_path=icon_file
        )
    else:
        TaskDialog.Show("Error", "No hangers selected")

except Exception:
    # Silently exit if user cancels the selection
    pass