# -*- coding: utf-8 -*-
import os
import math
import clr

clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
clr.AddReference('System')
clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')

from Autodesk.Revit import DB
from Autodesk.Revit.DB import (
    FilteredElementCollector, Transaction, TransactionGroup,
    Family, FamilySymbol, BuiltInCategory, IndependentTag,
    TagOrientation, Reference, BuiltInParameter
)
from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType

from System.Windows.Forms import NotifyIcon, ToolTipIcon
from System.Drawing import Icon, SystemIcons

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
curview = doc.ActiveView

script_dir, script_file = os.path.split(__file__)
FAMILY_NAME = 'Fabrication Duct - Size Tag'
TYPE_NAME = 'NONE'

FAMILY_FILE_1 = 'Fabrication Duct - Size Tag.rfa'
FAMILY_PATH_1 = os.path.join(script_dir, FAMILY_FILE_1)

# Keeps NotifyIcon alive long enough for the balloon to render
_active_notifications = []


def show_balloon_notification(title, message, tooltip_icon=None, timeout=5000):
    """Displays a native Windows balloon notification."""
    if tooltip_icon is None:
        tooltip_icon = ToolTipIcon.Info

    notify_icon = NotifyIcon()
    _active_notifications.append(notify_icon)
    try:
        icon_path = os.path.join(script_dir, 'Murray.ico')
        if os.path.exists(icon_path):
            notify_icon.Icon = Icon(icon_path)
        elif tooltip_icon == ToolTipIcon.Error:
            notify_icon.Icon = SystemIcons.Error
        else:
            notify_icon.Icon = SystemIcons.Information

        notify_icon.Visible = True
        notify_icon.ShowBalloonTip(timeout, title, message, tooltip_icon)
    except Exception:
        pass


def show_error_notification(message, timeout=5000):
    show_balloon_notification("Fabrication Duct Size Tags", message, ToolTipIcon.Error, timeout)


class FamilyLoaderOptionsHandler(DB.IFamilyLoadOptions):
    def OnFamilyFound(self, familyInUse, overwriteParameterValues):
        overwriteParameterValues.Value = False
        return True

    def OnSharedFamilyFound(self, sharedFamily, familyInUse, source, overwriteParameterValues):
        source.Value = DB.FamilySource.Family
        overwriteParameterValues.Value = False
        return True


class FabricationStraightFilter(ISelectionFilter):
    """Restricts selection exclusively to Fabrication Straight Parts."""
    def AllowElement(self, element):
        try:
            if element.Category and element.Category.Id.IntegerValue == int(BuiltInCategory.OST_FabricationDuctwork):
                return True
        except:
            pass
        return False

    def AllowReference(self, reference, point):
        return False


def get_tag_symbol(doc, family_name, type_name):
    symbols = FilteredElementCollector(doc).OfClass(FamilySymbol).ToElements()
    for sym in symbols:
        try:
            sym_name = sym.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString()
            if sym.Family.Name == family_name and (not type_name or sym_name == type_name):
                return sym
        except:
            pass
    return None


def get_element_connector_position_and_orientation(element):
    """Calculates position near connector, determines orientation, and applies consistent spacing."""
    try:
        width_inches = 0.0
        depth_inches = 0.0

        w_param = element.LookupParameter("Main Primary Width") or element.LookupParameter("Width")
        if w_param and w_param.HasValue:
            width_inches = round(w_param.AsDouble() * 12.0)

        d_param = element.LookupParameter("Main Primary Depth") or element.LookupParameter("Height") or element.LookupParameter("Diameter")
        if d_param and d_param.HasValue:
            depth_inches = round(d_param.AsDouble() * 12.0)

        is_under_24 = (width_inches > 0 and width_inches < 24.0) or (depth_inches > 0 and depth_inches < 24.0)

        conn_mgr = element.ConnectorManager
        if conn_mgr and conn_mgr.Connectors:
            connectors = [c for c in conn_mgr.Connectors]
            if len(connectors) >= 2:
                p1 = connectors[0].Origin
                p2 = connectors[1].Origin

                midpoint = DB.XYZ((p1.X + p2.X) / 2.0, (p1.Y + p2.Y) / 2.0, (p1.Z + p2.Z) / 2.0)

                vec = p2 - p1
                dx = abs(vec.X)
                dy = abs(vec.Y)

                # Determine orientation
                if dx > dy:
                    orientation = TagOrientation.Horizontal if is_under_24 else TagOrientation.Vertical
                else:
                    orientation = TagOrientation.Vertical if is_under_24 else TagOrientation.Horizontal

                # If the tag orientation matches the run direction (meaning text runs parallel along the duct length),
                # push it farther away so it doesn't crowd the connector. If it runs across the width, keep it close.
                is_parallel_to_run = (dx > dy and orientation == TagOrientation.Horizontal) or (dy > dx and orientation == TagOrientation.Vertical)
                offset_distance = 1.35 if is_parallel_to_run else 0.75

                direction_vector = (midpoint - p1).Normalize() if p1.DistanceTo(midpoint) > 0 else DB.XYZ(0, 0, 0)
                tag_position = p1 + direction_vector * offset_distance

                return tag_position, orientation
    except:
        pass

    # Fallback
    try:
        loc = element.Location
        if hasattr(loc, 'Point'):
            return loc.Point, TagOrientation.Horizontal
        elif hasattr(loc, 'Curve'):
            return loc.Curve.Evaluate(0.5, True), TagOrientation.Horizontal
    except:
        pass

    bbox = element.get_BoundingBox(curview)
    if bbox:
        return DB.XYZ((bbox.Min.X + bbox.Max.X)/2, (bbox.Min.Y + bbox.Max.Y)/2, (bbox.Min.Z + bbox.Max.Z)/2), TagOrientation.Horizontal

    return None, TagOrientation.Horizontal


def main():
    if curview.ViewType == DB.ViewType.ThreeD:
        show_error_notification("Cannot use this tool in a 3D view.")
        return

    families = FilteredElementCollector(doc).OfClass(Family)
    fam_in_project = any(f.Name == FAMILY_NAME for f in families)

    tg = TransactionGroup(doc, "Batch Tag Fabrication Duct Size")
    tg.Start()

    if not fam_in_project:
        t_load = Transaction(doc, "Load Tag Families")
        t_load.Start()
        try:
            load_handler = FamilyLoaderOptionsHandler()
            if os.path.exists(FAMILY_PATH_1):
                doc.LoadFamily(FAMILY_PATH_1, load_handler)
                doc.Regenerate()
            t_load.Commit()
        except Exception as e:
            if t_load.HasStarted():
                t_load.RollBack()
            tg.RollBack()
            show_error_notification("Failed to load families: {}".format(str(e)))
            return

    tag_symbol = get_tag_symbol(doc, FAMILY_NAME, TYPE_NAME)
    if not tag_symbol:
        tg.RollBack()
        show_error_notification("Could not find tag symbol '{}'.".format(TYPE_NAME))
        return

    selected_ids = list(uidoc.Selection.GetElementIds())
    elements_to_tag = []

    if selected_ids:
        elements_to_tag = [doc.GetElement(eid) for eid in selected_ids]
    else:
        try:
            selection_filter = FabricationStraightFilter()
            picked_refs = uidoc.Selection.PickObjects(
                ObjectType.Element,
                selection_filter,
                "Select fabrication straight parts to tag."
            )
            elements_to_tag = [doc.GetElement(ref.ElementId) for ref in picked_refs]
        except:
            tg.RollBack()
            return

    if not elements_to_tag:
        tg.RollBack()
        return

    t_tag = Transaction(doc, "Place Size Tags")
    try:
        t_tag.Start()
        if not tag_symbol.IsActive:
            tag_symbol.Activate()
            doc.Regenerate()

        tagged_count = 0
        for elem in elements_to_tag:
            tag_position, orientation = get_element_connector_position_and_orientation(elem)
            if not tag_position:
                continue

            ref = Reference(elem)
            new_tag = IndependentTag.Create(
                doc,
                tag_symbol.Id,
                curview.Id,
                ref,
                False,  # Include leader = False
                orientation,
                tag_position
            )

            if new_tag:
                doc.Regenerate()
                try:
                    new_tag.TagHeadPosition = tag_position
                    doc.Regenerate()
                except:
                    pass
                tagged_count += 1

        t_tag.Commit()
        tg.Assimilate()

        show_balloon_notification(
            "Fabrication Duct Size Tags",
            "Successfully placed {} size tag(s).".format(tagged_count)
        )

    except Exception as e:
        if t_tag.HasStarted():
            t_tag.RollBack()
        tg.RollBack()
        show_error_notification("Error during tagging: {}".format(str(e)))


if __name__ == '__main__':
    main()