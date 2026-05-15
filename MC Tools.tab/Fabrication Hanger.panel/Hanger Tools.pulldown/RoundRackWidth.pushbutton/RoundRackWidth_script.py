# -*- coding: utf-8 -*-
import math

from Autodesk.Revit import DB
from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument


# -----------------------------
# Selection Filter
# -----------------------------
class HangerSelectionFilter(ISelectionFilter):
    def AllowElement(self, element):
        cat = element.Category
        return cat and cat.Name == "MEP Fabrication Hangers"

    def AllowReference(self, reference, point):
        return True


# -----------------------------
# Helpers
# -----------------------------
def round_up(value, multiple):
    return multiple * math.ceil(value / float(multiple))


# -----------------------------
# Select Hangers
# -----------------------------
try:
    refs = uidoc.Selection.PickObjects(
        ObjectType.Element,
        HangerSelectionFilter(),
        "Select Fabrication Hangers"
    )

except Exception:
    refs = []


if refs:

    hangers = [doc.GetElement(r) for r in refs]

    t = DB.Transaction(doc, "Round Trapeze Width")
    t.Start()

    for hanger in hangers:

        hosted_info = hanger.GetHostedInfo()
        if hosted_info:
            hosted_info.DisconnectFromHost()

        width_value = None
        bearer_value = None

        # Collect required dimensions
        for dim in hanger.GetDimensions():

            if dim.Name == "Width":
                width_value = hanger.GetDimensionValue(dim)

            elif dim.Name == "Bearer Extn":
                bearer_value = hanger.GetDimensionValue(dim)
                bearer_dim = dim

        # Skip invalid hangers
        if width_value is None or bearer_value is None:
            continue

        # Convert to inches
        width_inches = width_value * 12.0
        bearer_inches = bearer_value * 12.0

        extra_bearer = max(bearer_inches - 4.0, 0.0)

        rounded = round_up(
            width_inches + bearer_inches + extra_bearer,
            2.0
        )

        adjusted = rounded - width_inches
        half_diff = (adjusted - 4.0) / 2.0

        new_bearer = (adjusted - half_diff) / 12.0

        hanger.SetDimensionValue(bearer_dim, new_bearer)

    t.Commit()