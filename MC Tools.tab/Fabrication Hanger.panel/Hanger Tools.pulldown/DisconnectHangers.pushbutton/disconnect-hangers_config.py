# -*- coding: UTF-8 -*-
import os
import clr
clr.AddReference("System.Windows.Forms")
clr.AddReference("System.Drawing")
from System.Windows.Forms import NotifyIcon, ToolTipIcon
from System.Drawing import Icon, SystemIcons

from Autodesk.Revit.DB import Transaction, FabricationPart, XYZ, BuiltInParameter
from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType
from Autodesk.Revit.UI import TaskDialog

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

# Tolerance for whether a hanger counts as "within" a host segment's span,
# to account for gaps/couplings between adjacent straight segments.
SPAN_TOLERANCE_FT = 0.5


class HangerSelectionFilter(ISelectionFilter):
    def AllowElement(self, e):
        return isinstance(e, FabricationPart) and e.IsAHanger()

    def AllowReference(self, ref, point):
        return True


class HostSelectionFilter(ISelectionFilter):
    def AllowElement(self, e):
        return isinstance(e, FabricationPart) and not e.IsAHanger()

    def AllowReference(self, ref, point):
        return True


def get_location_point(element):
    loc = element.Location
    if hasattr(loc, "Point") and loc.Point is not None:
        return loc.Point
    bbox = element.get_BoundingBox(None)
    return XYZ(
        (bbox.Min.X + bbox.Max.X) / 2.0,
        (bbox.Min.Y + bbox.Max.Y) / 2.0,
        (bbox.Min.Z + bbox.Max.Z) / 2.0,
    )


def build_host_data(host_part):
    """Return (host_part, c0, c1, axis, length) for a straight host segment,
    or None if the part doesn't have exactly 2 connectors (not a straight run)."""
    connectors = list(host_part.ConnectorManager.Connectors)
    if len(connectors) != 2:
        return None
    c0, c1 = connectors[0], connectors[1]
    p0, p1 = c0.Origin, c1.Origin
    length = p0.DistanceTo(p1)
    if length < 1e-6:
        return None
    axis = (p1 - p0).Normalize()
    return (host_part, c0, c1, axis, length)


def match_hanger_to_host(hanger_point, host_data_list):
    """Find the host whose span the hanger falls within, preferring the
    closest perpendicular distance among candidates. Returns
    (host_part, host_connector, distance) or None if no host qualifies."""
    best = None
    best_perp = None
    for host_part, c0, c1, axis, length in host_data_list:
        vec = hanger_point - c0.Origin
        t = vec.DotProduct(axis)  # distance along axis from c0
        perp_vec = vec - axis.Multiply(t)
        perp_dist = perp_vec.GetLength()

        if -SPAN_TOLERANCE_FT <= t <= length + SPAN_TOLERANCE_FT:
            # Clamp into valid range for PlaceOnHost
            t_clamped = min(max(t, 0.001), length - 0.001)
            if best_perp is None or perp_dist < best_perp:
                # Use whichever end connector is closer along the axis
                if t_clamped <= length / 2.0:
                    best = (host_part, c0, t_clamped)
                else:
                    best = (host_part, c1, length - t_clamped)
                best_perp = perp_dist
    return best


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


try:
    hanger_filter = HangerSelectionFilter()
    selected_hanger_refs = uidoc.Selection.PickObjects(
        ObjectType.Element, hanger_filter, "Select disconnected Fabrication Hangers to reconnect"
    )
    fab_hangers = [doc.GetElement(r.ElementId) for r in selected_hanger_refs]

    if not fab_hangers:
        TaskDialog.Show("Error", "No hangers selected")
    else:
        host_filter = HostSelectionFilter()
        selected_host_refs = uidoc.Selection.PickObjects(
            ObjectType.Element, host_filter, "Select all host fabrication parts (duct/pipe run) that could apply"
        )
        host_parts = [doc.GetElement(r.ElementId) for r in selected_host_refs]

        # Build geometry data for each host once, before the transaction.
        host_data_list = []
        skipped_hosts = []
        for hp in host_parts:
            data = build_host_data(hp)
            if data is None:
                skipped_hosts.append(hp.Id)
            else:
                host_data_list.append(data)

        if skipped_hosts:
            TaskDialog.Show(
                "Note",
                "Skipped {} selected host part(s) that are not straight "
                "2-connector segments (fittings/couplings can't be a "
                "PlaceOnHost target): {}".format(
                    len(skipped_hosts), ", ".join(str(i) for i in skipped_hosts)
                ),
            )

        t = Transaction(doc, "Reconnect Hangers to Matched Hosts")
        t.Start()

        successful_connects = 0
        unmatched = []
        for hanger in fab_hangers:
            try:
                if not hanger.IsAHanger():
                    TaskDialog.Show("Error", "Element {} is not a hanger".format(hanger.Id))
                    continue

                hosted_info = hanger.GetHostedInfo()
                if hosted_info is None:
                    TaskDialog.Show(
                        "Error",
                        "Hanger {} returned no FabricationHostedInfo. It may not support hosting.".format(hanger.Id),
                    )
                    continue

                hanger_point = get_location_point(hanger)
                match = match_hanger_to_host(hanger_point, host_data_list)

                if match is None:
                    unmatched.append(hanger.Id)
                    continue

                host_part, host_connector, distance = match
                hosted_info.PlaceOnHost(host_part.Id, host_connector, distance)
                successful_connects += 1

            except Exception as e:
                TaskDialog.Show("Error", "Failed to reconnect hanger {}: {}".format(hanger.Id, str(e)))

        doc.Regenerate()
        t.Commit()

        msg = "Successfully reconnected {} hangers.".format(successful_connects)
        if unmatched:
            msg += "\nUnmatched hangers: {}".format(", ".join(str(i) for i in unmatched))

        # Optional: Set a custom icon path if desired, e.g., r"C:\path\to\Murray.ico"
        icon_file = None
        
        # Show native Windows balloon notification instead of TaskDialog
        show_balloon_notification(
            "Hanger Reconnection Complete",
            msg,
            icon_path=icon_file
        )

except Exception:
    # Silently exit if user cancels the selection
    pass