import os
import Autodesk
from Autodesk.Revit import DB
from Autodesk.Revit.DB import (
    BuiltInCategory,
    Family,
    FamilySymbol,
    FilteredElementCollector,
    IndependentTag,
    TagOrientation,
    Transaction,
)
from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType

app = __revit__.Application
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

path, filename = os.path.split(__file__)
NewFilename1 = '\\Fabrication Duct - Duct Elevation Tag - Aligned.rfa'


# Custom selection filter to ONLY allow Fabrication Straight parts
class FabricationStraightFilter(ISelectionFilter):

  def AllowElement(self, elem):
    if isinstance(elem, DB.FabricationPart):
      try:
        return elem.IsAStraight()
      except Exception:
        return False
    return False

  def AllowReference(self, reference, pos):
    return False


# Search project for all Families
families = FilteredElementCollector(doc).OfClass(Family)
FamilyName = 'Fabrication Duct - Duct Elevation Tag - Aligned'
Fam_is_in_project = any(f.Name == FamilyName for f in families)

family_pathCC1 = path + NewFilename1

# Load family if not already in project
if not Fam_is_in_project:
  t = Transaction(doc, 'Load Duct Elev Family')
  t.Start()
  doc.LoadFamily(family_pathCC1)
  t.Commit()

# Family symbol name to look for
symbName = 'Fabrication Duct - Duct Elevation Tag - Aligned'

# Collect the tag FamilySymbol
collector = (
    FilteredElementCollector(doc)
    .OfCategory(BuiltInCategory.OST_FabricationDuctworkTags)
    .OfClass(FamilySymbol)
)

tag_symbol = None
for symb in collector:
  if symb.Family.Name == symbName:
    tag_symbol = symb
    break

if not tag_symbol:
  print("Error: Tag symbol '{}' not found in project!".format(symbName))
else:
  # Prompt user to multi-pick ONLY fabrication straight parts using the custom filter
  try:
    selection_filter = FabricationStraightFilter()
    references = uidoc.Selection.PickObjects(
        ObjectType.Element,
        selection_filter,
        'Select fabrication duct straights to tag',
    )
  except Exception:
    references = []

  if references:
    active_view = doc.ActiveView

    t2 = Transaction(doc, 'Batch Tag Ducts')
    t2.Start()

    created_count = 0
    for ref in references:
      element = doc.GetElement(ref)
      if element:
        connector_origins = []
        is_round = False
        diam_inches = 0.0

        try:
          conn_mgr = element.ConnectorManager
          if conn_mgr:
            conn_iter = conn_mgr.Connectors.GetEnumerator()
            while conn_iter.MoveNext():
              conn = conn_iter.Current
              connector_origins.append(conn.Origin)
              if conn.Shape == DB.ConnectorProfileType.Round:
                is_round = True
                # Connector radius is in feet, convert to inches
                diam_inches = conn.Radius * 2.0 * 12.0
        except Exception:
          pass

        if len(connector_origins) >= 2:
          p1 = connector_origins[0]
          p2 = connector_origins[1]
          center_pt = (p1 + p2) * 0.5

          # Check if round duct and diameter is less than 14 inches
          if is_round and diam_inches < 14.0:
            dx = abs(p2.X - p1.X)
            dy = abs(p2.Y - p1.Y)
            offset = 1.25  # Offset distance in feet (~15 inches)

            if dx >= dy:
              # Running X axis (left to right) -> Above (+Y)
              pt = DB.XYZ(center_pt.X, center_pt.Y + offset, center_pt.Z)
            else:
              # Running Y axis (top to bottom) -> Left and Above (-X, +Y)
              pt = DB.XYZ(center_pt.X - offset, center_pt.Y + offset, center_pt.Z)
          else:
            # 14" and larger or non-round: place directly in the center
            pt = center_pt
        else:
          loc = element.Location
          if hasattr(loc, 'Curve'):
            pt = loc.Curve.Evaluate(0.5, True)
          else:
            bbox = element.get_BoundingBox(active_view)
            if bbox:
              pt = (bbox.Min + bbox.Max) * 0.5
            else:
              pt = DB.XYZ(0, 0, 0)

        try:
          # Create the tag without a leader
          tag = IndependentTag.Create(
              doc,
              tag_symbol.Id,
              active_view.Id,
              ref,
              False,  # No leader
              TagOrientation.Horizontal,
              pt,
          )
          if tag:
            created_count += 1
        except Exception as ex:
          print('Error tagging element ID {}: {}'.format(element.Id, ex))

    t2.Commit()