# -*- coding: utf-8 -*-

import os
from Autodesk.Revit import DB
from Autodesk.Revit.DB import (
    FilteredElementCollector,
    BuiltInCategory,
    FamilySymbol,
    FabricationPart,
    IndependentTag,
    TagOrientation,
    Reference,
    Transaction
)
from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType
from Autodesk.Revit.Exceptions import OperationCanceledException
from Autodesk.Revit.UI import TaskDialog

from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType
from Autodesk.Revit.Exceptions import OperationCanceledException
from Autodesk.Revit.UI import TaskDialog

from Parameters.Get_Set_Params import set_parameter_by_name

from Parameters.Add_SharedParameters import Shared_Params
Shared_Params()

Shared_Params()

TITLE = 'Fab Duct Size'

# -----------------------------------------------------------------------------
# Revit handles
# -----------------------------------------------------------------------------
uidoc = __revit__.ActiveUIDocument
doc = uidoc.Document
view = doc.ActiveView

# -----------------------------------------------------------------------------
# Family names / paths
# -----------------------------------------------------------------------------
script_dir = os.path.dirname(__file__)

TOP_FAMILY_NAME = 'Fabrication Duct - Top Extension'
BOT_FAMILY_NAME = 'Fabrication Duct - Bottom Extension'

TOP_FAMILY_FILE = os.path.join(script_dir, 'Fabrication Duct - Top Extension.rfa')
BOT_FAMILY_FILE = os.path.join(script_dir, 'Fabrication Duct - Bottom Extension.rfa')


# -----------------------------------------------------------------------------
# Load options
# -----------------------------------------------------------------------------
class FamilyLoaderOptionsHandler(DB.IFamilyLoadOptions):
    def OnFamilyFound(self, familyInUse, overwriteParameterValues):
        overwriteParameterValues.Value = False
        return True

    def OnSharedFamilyFound(self, sharedFamily, familyInUse, source, overwriteParameterValues):
        source.Value = DB.FamilySource.Family
        overwriteParameterValues.Value = False
        return True


# -----------------------------------------------------------------------------
# Selection filter
# -----------------------------------------------------------------------------
class FabDuctFittingSelectionFilter(ISelectionFilter):
    def AllowElement(self, elem):
        if not isinstance(elem, FabricationPart):
            return False

        cat = elem.Category
        if cat is None:
            return False

        return cat.Id.IntegerValue == int(BuiltInCategory.OST_FabricationDuctwork)

    def AllowReference(self, reference, point):
        return False


# -----------------------------------------------------------------------------
# Helpers
# -----------------------------------------------------------------------------
def get_tag_symbol(doc, family_name):
    collector = (FilteredElementCollector(doc)
                 .OfCategory(BuiltInCategory.OST_FabricationDuctworkTags)
                 .OfClass(FamilySymbol))

    for sym in collector:
        try:
            if sym.Family and sym.Family.Name == family_name:
                return sym
        except:
            pass
    return None


def ensure_tag_symbol_loaded(doc, family_name, family_path):
    sym = get_tag_symbol(doc, family_name)
    if sym:
        return sym

    if not os.path.exists(family_path):
        raise Exception("Family file not found:\n{}".format(family_path))

    load_opts = FamilyLoaderOptionsHandler()
    doc.LoadFamily(family_path, load_opts)

    sym = get_tag_symbol(doc, family_name)
    if not sym:
        raise Exception("Family loaded but tag symbol not found:\n{}".format(family_name))

    return sym


def activate_symbol(doc, symbol):
    if not symbol.IsActive:
        symbol.Activate()
        doc.Regenerate()


def get_end_connectors(part):
    connectors = []

    try:
        conn_mgr = part.ConnectorManager
    except:
        return connectors

    if conn_mgr is None:
        return connectors

    for conn in conn_mgr.Connectors:
        try:
            _ = conn.Origin
            if conn.ConnectorType == DB.ConnectorType.End:
                connectors.append(conn)
        except:
            pass

    try:
        connectors = sorted(connectors, key=lambda c: c.Id)
    except:
        pass

    return connectors


def place_tag(doc, view, symbol, part, point):
    return IndependentTag.Create(
        doc,
        symbol.Id,
        view.Id,
        Reference(part),
        False,
        TagOrientation.Horizontal,
        point
    )


# -----------------------------------------------------------------------------
# Main
# -----------------------------------------------------------------------------
if view.IsTemplate:
    raise Exception("Active view cannot be a template view.")

try:
    picked_ref = uidoc.Selection.PickObject(
        ObjectType.Element,
        FabDuctFittingSelectionFilter(),
        "Select fabrication duct fitting"
    )
except OperationCanceledException:
    TaskDialog.Show(TITLE, "No fabrication duct fitting was selected.")
    raise SystemExit

t = Transaction(doc, "Place Fab Top/Bottom Extension Tags")
t.Start()

try:
    top_symbol = ensure_tag_symbol_loaded(doc, TOP_FAMILY_NAME, TOP_FAMILY_FILE)
    bot_symbol = ensure_tag_symbol_loaded(doc, BOT_FAMILY_NAME, BOT_FAMILY_FILE)

    activate_symbol(doc, top_symbol)
    activate_symbol(doc, bot_symbol)

    part = doc.GetElement(picked_ref.ElementId)
    
    if part is None:
        skipped.append("Selected element could not be read.")
    else:
        try:
            TOPE = None
            BOTE = None

            ItmDims = part.GetDimensions()
            for dta in ItmDims:
                if dta.Name == 'Top Extension':
                    TOPE = part.GetDimensionValue(dta)
                elif dta.Name == 'Bottom Extension':
                    BOTE = part.GetDimensionValue(dta)

            if TOPE is not None:
                set_parameter_by_name(part, 'FP_Extension Top', TOPE)

            if BOTE is not None:
                set_parameter_by_name(part, 'FP_Extension Bottom', BOTE)
        except:
            pass

        connectors = get_end_connectors(part)

        if len(connectors) < 2:
            skipped.append("Element {}: fewer than 2 end connectors".format(part.Id.IntegerValue))
        else:
            c1 = connectors[0]
            c2 = connectors[1]

            try:
                place_tag(doc, view, top_symbol, part, c1.Origin)
                place_tag(doc, view, bot_symbol, part, c2.Origin)
                placed_count = 1
            except Exception as ex:
                skipped.append("Element {}: {}".format(part.Id.IntegerValue, str(ex)))

    t.Commit()

except Exception:
    t.RollBack()
    raise

# msg = "Processed fittings: 1\nPlaced tag pairs: {}".format(placed_count)

# if skipped:
    # msg += "\n\nSkipped:\n- " + "\n- ".join(skipped[:20])

# TaskDialog.Show(TITLE, msg)