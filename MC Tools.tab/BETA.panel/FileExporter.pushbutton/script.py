# -*- coding: utf-8 -*-

import os, json, csv, datetime, re

import clr
clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")
clr.AddReference("System.Xaml")
clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
clr.AddReference("System.Windows.Forms")

import System
import System.Windows             as SW
import System.Windows.Controls   as SWC
import System.Windows.Media      as SWM
import System.Windows.Data       as SWD
import System.Windows.Input      as SWI
import System.Windows.Threading  as SWT
import System.Windows.Interop    as Interop
import System.Diagnostics        as Diagnostics
import Autodesk.Revit.DB         as DB

from System.Collections.ObjectModel import ObservableCollection
from System.Collections.Generic     import List as GList
from Autodesk.Revit.UI              import IExternalEventHandler, ExternalEvent
from System                         import Array, Action
import __main__

# ── Short aliases ─────────────────────────────────────────────────────────────
Window              = SW.Window
Thickness           = SW.Thickness
Visibility          = SW.Visibility
GridLength          = SW.GridLength
GridUnitType        = SW.GridUnitType
FontWeights         = SW.FontWeights
FontStyles          = SW.FontStyles
HorizontalAlignment = SW.HorizontalAlignment
VerticalAlignment   = SW.VerticalAlignment
Dock                = SWC.Dock

Grid                = SWC.Grid
RowDefinition       = SWC.RowDefinition
ColumnDefinition    = SWC.ColumnDefinition
StackPanel          = SWC.StackPanel
DockPanel           = SWC.DockPanel
Button              = SWC.Button
TextBox             = SWC.TextBox
ComboBox            = SWC.ComboBox
CheckBox            = SWC.CheckBox
RadioButton         = SWC.RadioButton
TextBlock           = SWC.TextBlock
Border              = SWC.Border
ScrollViewer        = SWC.ScrollViewer
GroupBox            = SWC.GroupBox
DataGrid            = SWC.DataGrid
DataGridRow              = SWC.DataGridRow
GridSplitter             = SWC.GridSplitter
GridResizeDirection      = SWC.GridResizeDirection
GridResizeBehavior       = SWC.GridResizeBehavior
DataGridTextColumn       = SWC.DataGridTextColumn
DataGridCheckBoxColumn   = SWC.DataGridCheckBoxColumn
DataGridLength           = SWC.DataGridLength
DataGridLengthUnitType   = SWC.DataGridLengthUnitType
DataGridSelectionMode    = SWC.DataGridSelectionMode
DataGridSelectionUnit    = SWC.DataGridSelectionUnit
DataGridGridLinesVisibility  = SWC.DataGridGridLinesVisibility
DataGridHeadersVisibility    = SWC.DataGridHeadersVisibility
ScrollBarVisibility          = SWC.ScrollBarVisibility
Orientation          = SWC.Orientation

SolidColorBrush  = SWM.SolidColorBrush
Color            = SWM.Color
Brushes          = SWM.Brushes
FontFamily       = SWM.FontFamily
VisualTreeHelper = SWM.VisualTreeHelper

Binding          = SWD.Binding
BindingMode      = SWD.BindingMode

Keyboard                = SWI.Keyboard
Key                     = SWI.Key
Cursors                 = SWI.Cursors
MouseButtonEventHandler = SWI.MouseButtonEventHandler
MouseEventHandler       = SWI.MouseEventHandler

# ── Revit handles ─────────────────────────────────────────────────────────────
doc   = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
app   = __revit__.Application

# ── Profile storage ───────────────────────────────────────────────────────────
PROFILES_DIR = os.path.join(
    os.environ.get("APPDATA", os.path.expanduser("~")),
    "pyRevit_SheetsExporter"
)
if not os.path.exists(PROFILES_DIR):
    os.makedirs(PROFILES_DIR)

# ── Brushes / theme ───────────────────────────────────────────────────────────
def rgb(r, g, b):
    return SolidColorBrush(Color.FromRgb(r, g, b))

BR_BG       = rgb(245, 246, 248)
BR_WHITE    = Brushes.White
BR_HEADER   = rgb(28,  28,  36)
BR_ACCENT   = rgb(22, 160, 133)
BR_GREEN    = rgb(40, 140,  90)
BR_RED      = rgb(190,  50,  50)
BR_GREY_BTN = rgb(210, 212, 218)
BR_TEXT     = rgb(30,  30,  38)
BR_SUBTEXT  = rgb(120, 120, 135)
BR_BORDER   = rgb(210, 212, 220)
BR_ROW_ALT  = rgb(248, 249, 251)
BR_PANEL    = rgb(240, 241, 244)
BR_TRANSP   = Brushes.Transparent

FONT_UI    = FontFamily("Segoe UI")

# ─────────────────────────────────────────────────────────────────────────────
#  DATA HELPERS
# ─────────────────────────────────────────────────────────────────────────────
def safe_fn(name):
    return re.sub(r'[\\/*?:"<>|]', "-", name or "").strip()

def get_sheets():
    return sorted(
        DB.FilteredElementCollector(doc).OfClass(DB.ViewSheet)
          .WhereElementIsNotElementType().ToElements(),
        key=lambda s: s.SheetNumber)

def get_views():
    """Get all non-sheet views suitable for NWC export."""
    exclude = [
        DB.ViewType.DrawingSheet,
        DB.ViewType.Schedule,
        DB.ViewType.ColumnSchedule,
        DB.ViewType.PanelSchedule,
        DB.ViewType.Undefined,
    ]
    views = DB.FilteredElementCollector(doc).OfClass(DB.View)              .WhereElementIsNotElementType().ToElements()
    result = []
    for v in views:
        if v.ViewType in exclude: continue
        if v.IsTemplate: continue
        result.append(v)
    return sorted(result, key=lambda v: (str(v.ViewType), v.Name))

def get_param(sheet, pname):
    p = sheet.LookupParameter(pname)
    if p and p.HasValue:
        return p.AsString() or ""
    if pname == "Sheet Number": return sheet.SheetNumber
    if pname == "Sheet Name":   return sheet.Name
    return ""

def get_size_label(sheet):
    try:
        w = sheet.LookupParameter("Sheet Width")
        h = sheet.LookupParameter("Sheet Height")
        if w and h and w.HasValue and h.HasValue:
            wi = w.AsDouble() * 304.8
            hi = h.AsDouble() * 304.8
            if abs(wi-431)<20 and abs(hi-279)<20: return "ANSI B"
            if abs(wi-279)<20 and abs(hi-216)<20: return "ANSI A"
            if abs(wi-559)<20 and abs(hi-432)<20: return "ANSI C"
            if abs(wi-864)<20 and abs(hi-559)<20: return "ANSI D"
            if abs(wi-297)<20 and abs(hi-210)<20: return "A4"
            return "{:.0f}x{:.0f}mm".format(wi, hi)
    except: pass
    return "Auto"

def build_filename(sheet, template, sep):
    tokens = {
        "{Sheet Number}"     : sheet.SheetNumber,
        "{Sheet Name}"       : sheet.Name,
        "{Current Revision}" : get_param(sheet, "Current Revision"),
        "{Date}"             : datetime.datetime.now().strftime("%Y%m%d"),
        "{Project Number}"   : (doc.ProjectInformation.Number if doc.ProjectInformation else ""),
        "{Project Name}"     : (doc.ProjectInformation.Name   if doc.ProjectInformation else ""),
    }
    r = template
    for t, v in tokens.items():
        r = r.replace(t, safe_fn(v))
    while sep * 2 in r:
        r = r.replace(sep * 2, sep)
    return r.strip(sep) or "Sheet"

def save_profile(name, data):
    with open(os.path.join(PROFILES_DIR, safe_fn(name) + ".json"), "w") as f:
        json.dump(data, f, indent=2)

def load_profile(name):
    p = os.path.join(PROFILES_DIR, safe_fn(name) + ".json")
    if os.path.exists(p):
        with open(p) as f: return json.load(f)
    return None

def list_profiles():
    return [f[:-5] for f in os.listdir(PROFILES_DIR) if f.endswith(".json")]

def del_profile(name):
    p = os.path.join(PROFILES_DIR, safe_fn(name) + ".json")
    if os.path.exists(p): os.remove(p)

def write_csv(folder, rows):
    path = os.path.join(folder, "ExportReport.csv")
    fields = ["Sheet Number","Sheet Name","Format","Filename","Status","Error","Timestamp"]
    ts = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    try:
        with open(path, "w") as f:
            w = csv.DictWriter(f, fieldnames=fields)
            w.writeheader()
            for r in rows:
                r["Timestamp"] = ts
                w.writerow(r)
        return path
    except: return None

# ─────────────────────────────────────────────────────────────────────────────
#  PRE-SELECTION HELPER
# ─────────────────────────────────────────────────────────────────────────────
def get_preselected_sheet_numbers():
    try:
        sel_ids = uidoc.Selection.GetElementIds()
        if not sel_ids or sel_ids.Count == 0:
            return None, None

        sheet_numbers   = set()
        view_unique_ids = set()
        assembly_names  = set()

        for eid in sel_ids:
            try:
                elem = doc.GetElement(eid)
                if elem is None: continue

                # Sheet selected
                if isinstance(elem, DB.ViewSheet):
                    sheet_numbers.add(elem.SheetNumber); continue

                # View selected in project browser
                if isinstance(elem, DB.View) and not isinstance(elem, DB.ViewSheet):
                    if not elem.IsTemplate:
                        view_unique_ids.add(elem.UniqueId)
                    continue

                # Assembly:Name param
                p = elem.LookupParameter("Assembly: Name")
                if p and p.HasValue and p.AsString():
                    assembly_names.add(p.AsString()); continue

                # m_spool_name param
                p2 = elem.LookupParameter("m_spool_name")
                if p2 and p2.HasValue and p2.AsString():
                    assembly_names.add(p2.AsString()); continue

                # Element name fallback
                try: assembly_names.add(elem.Name)
                except: pass

            except: continue

        # Resolve assembly names to sheet numbers
        if assembly_names:
            for sheet in get_sheets():
                p = sheet.LookupParameter("Assembly: Name")
                if p and p.HasValue and p.AsString() in assembly_names:
                    sheet_numbers.add(sheet.SheetNumber); continue
                if sheet.Name in assembly_names:
                    sheet_numbers.add(sheet.SheetNumber); continue
                if sheet.SheetNumber in assembly_names:
                    sheet_numbers.add(sheet.SheetNumber)

        return (sheet_numbers if sheet_numbers else None,
                view_unique_ids if view_unique_ids else None)
    except:
        return None, None

class SheetRow(object):
    def __init__(self, sheet):
        self.sheet       = sheet
        self.IsChecked   = False
        self.IsSelected  = False
        self.SheetNumber = sheet.SheetNumber
        self.SheetName   = sheet.Name
        self.Revision    = get_param(sheet, "Current Revision") or ""
        self.Size        = get_size_label(sheet)
        self.SpoolTag    = get_param(sheet, "SPOOL SHEET HCC IFF TAG") or ""

class ViewRow(object):
    def __init__(self, view):
        self.view       = view
        self.IsChecked  = False
        self.IsSelected = False
        self.SheetNumber = ""
        self.SheetName   = str(view.ViewType).replace("ViewType.", "")
        self.Revision    = ""
        self.Size        = ""
        self.SpoolTag    = view.Name
        self._chk_ref   = None

# ─────────────────────────────────────────────────────────────────────────────
#  EXTERNAL EVENT HANDLER
# ─────────────────────────────────────────────────────────────────────────────
class ExportSheetsHandler(IExternalEventHandler):

    def __init__(self):
        self.data = None

    def Execute(self, uiapp):
        import os, datetime, csv, re
        import Autodesk.Revit.DB as DB
        from System.Collections.Generic import List as GList
        def safe_fn(name):
            return re.sub(r'[\/*?:"<>|]', "-", name or "").strip()
        def write_csv(folder, rows):
            import csv as _csv
            path = os.path.join(folder, "ExportReport.csv")
            fields = ["Sheet Number","Sheet Name","Format","Filename","Status","Error","Timestamp"]
            ts = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            try:
                with open(path, "w") as f:
                    w = _csv.DictWriter(f, fieldnames=fields)
                    w.writeheader()
                    for r in rows:
                        r["Timestamp"] = ts
                        w.writerow(r)
            except: pass
        try:
            d          = self.data
            sheets     = d["sheets"]
            pairs      = d["pairs"]
            cfg        = d["settings"]
            out_folder = d["out_folder"]
            do_pdf     = d.get("do_pdf", False)
            do_dwg     = d.get("do_dwg", False)
            do_report  = d.get("do_report", False)
            callback   = d.get("callback")
            doc_l      = uiapp.ActiveUIDocument.Document

            def get_out(fmt):
                base = out_folder
                if d.get("date_sub"):
                    base = os.path.join(base, datetime.datetime.now().strftime("%Y-%m-%d"))
                if d.get("fmt_sub"):
                    base = os.path.join(base, fmt)
                if not os.path.exists(base): os.makedirs(base)
                return base

            # Clear Revit element selection so nothing prints highlighted/blue
            try:
                uidoc_l = uiapp.ActiveUIDocument
                uidoc_l.Selection.SetElementIds(GList[DB.ElementId]())
            except: pass

            def set_thin_lines(enable):
                """Temporarily override all line weights to 1 for thin lines export."""
                try:
                    lws = DB.LineWeightSettings.GetLineWeightSettings(doc_l)
                    cats = [
                        DB.LineWeightType.Cut,
                        DB.LineWeightType.Projection,
                    ]
                    saved = {}
                    for lwt in cats:
                        count = lws.GetLineWeightCount(lwt)
                        for i in range(1, count + 1):
                            saved[(lwt, i)] = lws.GetLineWeight(lwt, i)
                            if enable:
                                lws.SetLineWeight(lwt, i, 1)
                    return saved
                except:
                    return {}

            def restore_line_weights(saved):
                try:
                    lws = DB.LineWeightSettings.GetLineWeightSettings(doc_l)
                    for (lwt, i), w in saved.items():
                        lws.SetLineWeight(lwt, i, w)
                except: pass

            def pdf(pairs, out):
                opts = DB.PDFExportOptions()
                opts.Combine = cfg.get("combine", False)
                opts.StopOnError = False
                try:
                    cmap = {0: DB.PDFColorDepthType.Color,
                            1: DB.PDFColorDepthType.BlackAndWhite,
                            2: DB.PDFColorDepthType.GrayScale}
                    opts.ColorDepth = cmap.get(cfg.get("color_idx", 0),
                                               DB.PDFColorDepthType.Color)
                except: pass
                try:
                    qmap = {0: DB.PDFExportQualityType.DPI72,
                            1: DB.PDFExportQualityType.DPI150,
                            2: DB.PDFExportQualityType.DPI300,
                            3: DB.PDFExportQualityType.DPI600}
                    opts.ExportQuality = qmap.get(cfg.get("quality_idx", 2),
                                                  DB.PDFExportQualityType.DPI300)
                except: pass
                try: opts.AlwaysUseRaster          = cfg.get("raster_proc", False)
                except: pass
                try: opts.HideReferencePlane       = cfg.get("hide_ref", True)
                except: pass
                try: opts.HideUnreferencedViewTags = cfg.get("hide_unref", True)
                except: pass
                try: opts.HideScopeBoxes           = cfg.get("hide_scope", True)
                except: pass
                try: opts.HideCropBoundaries       = cfg.get("hide_crop", True)
                except: pass
                try:
                    # If thin lines checked, override ReplaceHalftoneWithThinLines to True
                    opts.ReplaceHalftoneWithThinLines = cfg.get("thin_lines", False) or cfg.get("replace_halftone", False)
                except: pass
                try: opts.MaskCoincidentLines = cfg.get("thin_lines", False)
                except: pass

                try:
                    if cfg.get("zoom_fit", True):
                        opts.ZoomType = DB.ZoomType.FitPage
                    else:
                        opts.ZoomType = DB.ZoomType.Zoom
                        opts.ZoomPercentage = cfg.get("zoom_pct", 100)
                except: pass

                # Thin lines: override halftone on every view on every sheet
                thin_t = None
                if cfg.get("thin_lines", False):
                    try:
                        thin_t = DB.Transaction(doc_l, "ThinLinesExport")
                        thin_t.Start()
                        # Process all view types that can have category overrides — skip only schedules
                        skip_types = [
                            DB.ViewType.Schedule,
                            DB.ViewType.ColumnSchedule,
                            DB.ViewType.PanelSchedule,
                            DB.ViewType.DrawingSheet,
                            DB.ViewType.Undefined,
                        ]
                        all_views = []
                        seen_ids = set()
                        for eid, _ in pairs:
                            sheet = doc_l.GetElement(eid)
                            if not sheet: continue
                            vps = DB.FilteredElementCollector(doc_l, sheet.Id)                                    .OfClass(DB.Viewport).ToElements()
                            for vp in vps:
                                try:
                                    v = doc_l.GetElement(vp.ViewId)
                                    if v and v.Id not in seen_ids and v.ViewType not in skip_types:
                                        all_views.append(v)
                                        seen_ids.add(v.Id)
                                except: pass

                        # For each view: detach template, apply halftone, reattach after export
                        cats = doc_l.Settings.Categories
                        saved_templates = {}  # view.Id -> original template Id
                        invalid_id = DB.ElementId.InvalidElementId

                        for view in all_views:
                            # Detach view template if one is applied
                            try:
                                tmpl_id = view.ViewTemplateId
                                if tmpl_id != invalid_id:
                                    saved_templates[view.Id] = tmpl_id
                                    view.ViewTemplateId = invalid_id
                            except: pass

                            # Apply halftone to all categories
                            for cat in cats:
                                try:
                                    ogs = view.GetCategoryOverrides(cat.Id)
                                    ogs.SetHalftone(True)
                                    view.SetCategoryOverrides(cat.Id, ogs)
                                except: pass
                                try:
                                    for subcat in cat.SubCategories:
                                        ogs2 = view.GetCategoryOverrides(subcat.Id)
                                        ogs2.SetHalftone(True)
                                        view.SetCategoryOverrides(subcat.Id, ogs2)
                                except: pass
                        opts.ReplaceHalftoneWithThinLines = True
                    except Exception as ex:
                        thin_t = None

                exported, failed = [], []
                if cfg.get("combine", False):
                    opts.FileName = safe_fn(cfg.get("combined_name", "Combined_Sheets"))
                    col = GList[DB.ElementId]()
                    for eid, _ in pairs: col.Add(eid)
                    try:
                        doc_l.Export(out, col, opts)
                        exported = [p[1] for p in pairs]
                    except Exception as ex:
                        failed.append(("ALL", str(ex)))
                else:
                    for eid, fname in pairs:
                        col = GList[DB.ElementId]()
                        col.Add(eid)
                        try:
                            target = os.path.join(out, fname + ".pdf")
                            # Find all pdfs in folder before export
                            before = set(f for f in os.listdir(out) if f.lower().endswith(".pdf"))
                            try:
                                _rs = d.get("row_started")
                                if _rs: _rs(fname)
                            except: pass
                            opts.FileName = fname
                            doc_l.Export(out, col, opts)
                            # Find what Revit actually created
                            after = set(f for f in os.listdir(out) if f.lower().endswith(".pdf"))
                            new_files = after - before
                            if new_files:
                                created = os.path.join(out, list(new_files)[0])
                                if created != target:
                                    if os.path.exists(target): os.remove(target)
                                    os.rename(created, target)
                            exported.append(fname)
                            try:
                                _rd = d.get("row_done")
                                if _rd: _rd(fname, True, len(exported))
                            except: pass
                        except Exception as ex:
                            failed.append((fname, str(ex)))

                # Restore view templates before rollback
                if thin_t is not None:
                    try:
                        for view in all_views:
                            try:
                                if view.Id in saved_templates:
                                    view.ViewTemplateId = saved_templates[view.Id]
                            except: pass
                    except: pass
                    try: thin_t.RollBack()
                    except: pass

                return exported, failed

            def dwg(pairs, out):
                exported, failed = [], []
                try: setups = DB.ExportDWGSettings.GetActivePredefinedSettings(doc_l)
                except: setups = None
                for eid, fname in pairs:
                    col = GList[DB.ElementId]()
                    col.Add(eid)
                    try:
                        if setups: doc_l.Export(out, fname, col, setups)
                        else:      doc_l.Export(out, fname, col, DB.DWGExportOptions())
                        exported.append(fname)
                        try:
                            _rd = d.get("row_done")
                            if _rd: _rd(fname, True, len(exported))
                        except: pass
                    except Exception as ex:
                        failed.append((fname, str(ex)))
                return exported, failed

            def nwc(view_ids, out):
                exported, failed = [], []
                try:
                    _proj = DB.NavisworksCoordinates.ProjectInternalOrigin
                except:
                    _proj = DB.NavisworksCoordinates.Internal
                coords_map = [
                    DB.NavisworksCoordinates.Shared,
                    DB.NavisworksCoordinates.Internal,
                    _proj,
                ]
                chosen_coords = coords_map[cfg.get("nwc_coords_idx", 0)]
                for vid, fname in view_ids:
                    try:
                        v_opts = DB.NavisworksExportOptions()
                        v_opts.ViewId      = vid
                        v_opts.ExportScope = DB.NavisworksExportScope.View
                        try: v_opts.Coordinates = chosen_coords
                        except: pass
                        # Apply user-controlled settings from UI
                        try: v_opts.ConvertElementIds        = cfg.get("nwc_convert_ids",    True)
                        except: pass
                        try: v_opts.ConvertElementProperties = cfg.get("nwc_convert_props",  False)
                        except: pass
                        try: v_opts.ExportLinks             = cfg.get("nwc_convert_links",  False)
                        except: pass
                        try: v_opts.ConvertLinkedFiles      = cfg.get("nwc_convert_links",  False)
                        except: pass
                        try: v_opts.ConvertRoomAsAttribute   = cfg.get("nwc_convert_room",   True)
                        except: pass
                        try: v_opts.ConvertUrls              = cfg.get("nwc_convert_urls",   True)
                        except: pass
                        try: v_opts.DivideFileIntoLevels     = cfg.get("nwc_divide_levels",  False)
                        except: pass
                        try: v_opts.ExportRoomGeometry       = cfg.get("nwc_export_rooms",   True)
                        except: pass
                        try: v_opts.FindMissingMaterials     = cfg.get("nwc_find_materials", True)
                        except: pass
                        try: v_opts.ConvertLinkedCADFormats  = cfg.get("nwc_convert_cad",    True)
                        except: pass
                        try: v_opts.ConvertLights            = cfg.get("nwc_convert_lights", False)
                        except: pass
                        try:
                            facet = float(cfg.get("nwc_faceting", "1") or "1")
                            v_opts.FacetingFactor = facet
                        except: pass
                        try:
                            _nwc_params = DB.NavisworksParameters
                            param_map = {
                                0: _nwc_params.All,
                                1: _nwc_params.Elements,
                                2: getattr(_nwc_params, "None", _nwc_params.All),
                            }
                            v_opts.Parameters = param_map.get(cfg.get("nwc_params_idx", 0), _nwc_params.All)
                        except: pass
                        target = os.path.join(out, fname + ".nwc")
                        before = set(f for f in os.listdir(out) if f.lower().endswith(".nwc"))
                        try:
                            _rs = d.get("row_started")
                            if _rs: _rs(fname)
                        except: pass
                        doc_l.Export(out, fname, v_opts)
                        after  = set(f for f in os.listdir(out) if f.lower().endswith(".nwc"))
                        new_files = after - before
                        if new_files:
                            created = os.path.join(out, list(new_files)[0])
                            if created != target:
                                if os.path.exists(target): os.remove(target)
                                os.rename(created, target)
                        exported.append(fname)
                        try:
                            _rd = d.get("row_done")
                            if _rd: _rd(fname, True, len(exported))
                        except: pass
                    except Exception as ex:
                        failed.append((fname, str(ex)))
                return exported, failed

            report_rows, all_exp, all_fail = [], [], []

            # Scheduled exports: reload latest + checkout before any export
            if d.get("is_scheduled", False):
                import System.Windows as _SWws
                # 1. Reload latest from central
                try:
                    if doc_l.IsWorkshared:
                        import Autodesk.Revit.DB as _DBpre
                        _rlo = _DBpre.ReloadLatestOptions()
                        doc_l.ReloadLatest(_rlo)
                except Exception as _rle:
                    _SWws.MessageBox.Show("Reload latest failed: {}".format(str(_rle)), "Scheduler")
                # 2. Check out all elements being exported
                try:
                    if doc_l.IsWorkshared:
                        from System.Collections.Generic import List as _GListPre
                        _co_ids = _GListPre[DB.ElementId]()
                        for _eid, _ in d.get("pairs", []):
                            try: _co_ids.Add(_eid)
                            except: pass
                        for _eid, _ in d.get("nwc_pairs", []):
                            try: _co_ids.Add(_eid)
                            except: pass
                        if _co_ids.Count > 0:
                            DB.WorksharingUtils.CheckoutElements(doc_l, _co_ids)
                except Exception as _coe:
                    _SWws.MessageBox.Show("Checkout failed: {}".format(str(_coe)), "Scheduler")

            if do_pdf:
                out = get_out("PDF")
                exp, fail = pdf(pairs, out)
                all_exp += exp; all_fail += fail
                for s, (_, fn) in zip(sheets, pairs):
                    ok = fn in exp or (cfg.get("combine") and bool(exp))
                    report_rows.append({
                        "Sheet Number": s.SheetNumber, "Sheet Name": s.Name,
                        "Format": "PDF", "Filename": fn + ".pdf",
                        "Status": "OK" if ok else "FAILED",
                        "Error": next((f[1] for f in fail if f[0] == fn), "")})

            if do_dwg:
                out = get_out("DWG")
                exp, fail = dwg(pairs, out)
                all_exp += exp; all_fail += fail
                for s, (_, fn) in zip(sheets, pairs):
                    report_rows.append({
                        "Sheet Number": s.SheetNumber, "Sheet Name": s.Name,
                        "Format": "DWG", "Filename": fn + ".dwg",
                        "Status": "OK" if fn in exp else "FAILED",
                        "Error": next((f[1] for f in fail if f[0] == fn), "")})

            do_nwc = d.get("do_nwc", False)
            if do_nwc:
                out = get_out("NWC")
                nwc_pairs = d.get("nwc_pairs", [])

                exp, fail = nwc(nwc_pairs, out)
                all_exp += exp; all_fail += fail


                for (vid, fn) in nwc_pairs:
                    report_rows.append({
                        "Sheet Number": fn, "Sheet Name": fn,
                        "Format": "NWC", "Filename": fn + ".nwc",
                        "Status": "OK" if fn in exp else "FAILED",
                        "Error": next((f[1] for f in fail if f[0] == fn), "")})

            if do_report and report_rows:
                write_csv(out_folder, report_rows)

            # Scheduled exports: sync to central after all exports complete
            if d.get("is_scheduled", False):
                import System.Windows as _SWsync
                try:
                    if doc_l.IsWorkshared:
                        import Autodesk.Revit.DB as _DBpost
                        _sync_opts = _DBpost.SynchronizeWithCentralOptions()
                        _relin = _DBpost.RelinquishOptions(True)
                        _sync_opts.SetRelinquishOptions(_relin)
                        _sync_opts.SaveLocalBefore = False
                        _sync_opts.SaveLocalAfter  = False
                        _sync_opts.Comment = "Scheduled Export - auto sync"
                        doc_l.SynchronizeWithCentral(
                            _DBpost.TransactWithCentralOptions(),
                            _sync_opts)
                except Exception as _synce:
                    _SWsync.MessageBox.Show("Sync failed: {}".format(str(_synce)), "Scheduler")

            msg = "Exported {} file(s) to:\n{}".format(len(all_exp), out_folder)
            if do_report: msg += "\n\nCSV report: ExportReport.csv"
            if all_fail:
                msg += "\n\nFailed ({}):\n".format(len(all_fail))
                msg += "\n".join("  {} - {}".format(*f) for f in all_fail[:8])

            if callback:
                callback(msg, len(all_fail) == 0)

        except Exception as ex:
            import traceback, System.Windows as _SWE
            _SWE.MessageBox.Show("Execute error: {}\n{}".format(str(ex), traceback.format_exc()), "Export Error")

    def GetName(self): return "ExportSheetsHandler"

# ─────────────────────────────────────────────────────────────────────────────
#  WPF WIDGET HELPERS
# ─────────────────────────────────────────────────────────────────────────────
def make_btn(text, bg=None, fg=None, width=None, height=28, bold=False):
    b = Button()
    b.Content         = text
    b.Background      = bg or BR_ACCENT
    b.Foreground      = fg or BR_WHITE
    b.BorderThickness = Thickness(0)
    b.Padding         = Thickness(12, 0, 12, 0)
    b.Height          = height
    b.FontFamily      = FONT_UI
    b.FontSize        = 12
    b.Cursor          = Cursors.Hand
    if bold: b.FontWeight = FontWeights.Bold
    if width: b.Width = width
    return b

def make_txt(val="", width=None, height=26, readonly=False):
    t = TextBox()
    t.Text        = val
    t.Height      = height
    t.FontFamily  = FONT_UI
    t.FontSize    = 12
    t.Padding     = Thickness(6, 2, 6, 2)
    t.Background  = BR_WHITE
    t.Foreground  = BR_TEXT
    t.IsReadOnly  = readonly
    if width: t.Width = width
    return t

def make_lbl(text, bold=False, size=12, fg=None):
    l = TextBlock()
    l.Text       = text
    l.FontFamily = FONT_UI
    l.FontSize   = size
    l.Foreground = fg or BR_TEXT
    l.VerticalAlignment = VerticalAlignment.Center
    if bold: l.FontWeight = FontWeights.Bold
    return l

def make_chk(text, checked=False):
    c = CheckBox()
    c.Content    = text
    c.IsChecked  = checked
    c.FontFamily = FONT_UI
    c.FontSize   = 12
    c.Foreground = BR_TEXT
    c.Margin     = Thickness(0, 3, 0, 3)
    return c

def make_rad(text, checked=False, group=None):
    r = RadioButton()
    r.Content    = text
    r.IsChecked  = checked
    r.FontFamily = FONT_UI
    r.FontSize   = 12
    r.Foreground = BR_TEXT
    r.Margin     = Thickness(0, 3, 0, 3)
    if group: r.GroupName = group
    return r

def make_cmb(items, selected=0, width=None):
    c = ComboBox()
    c.FontFamily = FONT_UI
    c.FontSize   = 12
    c.Height     = 26
    c.Padding    = Thickness(4, 0, 4, 0)
    for item in items: c.Items.Add(item)
    c.SelectedIndex = selected
    if width: c.Width = width
    return c

def make_grp(header, child, margin=None):
    g = GroupBox()
    g.Header      = header
    g.Content     = child
    g.FontFamily  = FONT_UI
    g.FontSize    = 11
    g.Foreground  = BR_SUBTEXT
    g.Background  = BR_WHITE
    g.Margin      = margin or Thickness(0, 0, 8, 8)
    g.Padding     = Thickness(8, 4, 8, 8)
    return g

def hsp(*children, **kw):
    s = StackPanel()
    s.Orientation = Orientation.Horizontal
    if kw.get("margin"): s.Margin = kw["margin"]
    for c in children:
        if c is not None: s.Children.Add(c)
    return s

def vsp(*children, **kw):
    s = StackPanel()
    s.Orientation = Orientation.Vertical
    if kw.get("margin"): s.Margin = kw["margin"]
    for c in children:
        if c is not None: s.Children.Add(c)
    return s

def hline():
    b = Border()
    b.Background = BR_BORDER
    b.Height     = 1
    b.Margin     = Thickness(0, 6, 0, 6)
    return b

# ─────────────────────────────────────────────────────────────────────────────
#  SCHEDULER PICKER WINDOW
# ─────────────────────────────────────────────────────────────────────────────
class SchedPickerWindow(Window):
    """Lightweight sheet/view picker for the scheduler."""

    def __init__(self, parent_window):
        self.parent_window  = parent_window
        self.selected_rows  = []
        self.output_folder  = ""
        self.all_sheets     = parent_window.all_sheets
        self.all_views      = parent_window.all_views
        self.sheet_rows     = []
        self.sheet_rows_list = []
        self.parent_window  = parent_window
        self._SheetRow      = parent_window._SheetRow
        self._ViewRow       = parent_window._ViewRow
        # Store module-level helper functions accessible at class creation time
        self._make_btn = parent_window._make_btn
        self._make_lbl = parent_window._make_lbl
        self._make_txt = parent_window._make_txt
        self._make_rad = parent_window._make_rad
        self._make_chk = parent_window._make_chk
        self._vsp      = parent_window._vsp
        self._hsp      = parent_window._hsp
        self.is_updating_selection = False
        self.last_clicked_row = None
        self.is_mouse_dragging = False
        self.drag_start_item   = None
        self.col_widths = parent_window.col_widths[:]

        import System.Windows as _SWP
        import System.Windows.Media as _SWMP
        from System.Windows.Media import FontFamily as _FFP
        self.Title  = "Configure Scheduled Export"
        self.Width  = 1000
        self.Height = 650
        self.WindowStartupLocation = _SWP.WindowStartupLocation.CenterOwner
        self.Background = parent_window.Background
        self.FontFamily = parent_window.FontFamily
        self.WindowStyle = _SWP.WindowStyle.SingleBorderWindow
        self.ResizeMode  = _SWP.ResizeMode.CanResize

        self._build()
        self._load_rows()

    def _build(self):
        import System.Windows as SW
        import System.Windows.Controls as SWC
        import System.Windows.Media as SWM
        import System.Windows.Input as SWI
        from System.Windows import Style, Setter
        from System.Windows.Controls import ListBoxItem
        BR_WHITE    = SWM.Brushes.White
        BR_GREY_BTN = SWM.SolidColorBrush(SWM.Color.FromRgb(210,212,218))
        BR_TEXT     = SWM.SolidColorBrush(SWM.Color.FromRgb(30,30,38))
        BR_SUBTEXT  = SWM.SolidColorBrush(SWM.Color.FromRgb(120,120,135))
        BR_BORDER   = SWM.SolidColorBrush(SWM.Color.FromRgb(210,212,220))
        BR_PANEL    = SWM.SolidColorBrush(SWM.Color.FromRgb(240,241,244))
        BR_ACCENT   = SWM.SolidColorBrush(SWM.Color.FromRgb(22,160,133))
        BR_RED      = SWM.SolidColorBrush(SWM.Color.FromRgb(190,50,50))
        BR_GREEN    = SWM.SolidColorBrush(SWM.Color.FromRgb(40,140,90))
        BR_ORANGE   = SWM.SolidColorBrush(SWM.Color.FromRgb(180,80,30))
        Thickness           = SW.Thickness
        GridLength          = SW.GridLength
        GridUnitType        = SW.GridUnitType
        FontWeights         = SW.FontWeights
        FontStyles          = SW.FontStyles
        HorizontalAlignment = SW.HorizontalAlignment
        VerticalAlignment   = SW.VerticalAlignment
        Visibility          = SW.Visibility
        Dock                = SWC.Dock
        Grid                = SWC.Grid
        RowDefinition       = SWC.RowDefinition
        ColumnDefinition    = SWC.ColumnDefinition
        StackPanel          = SWC.StackPanel
        DockPanel           = SWC.DockPanel
        Button              = SWC.Button
        CheckBox            = SWC.CheckBox
        RadioButton         = SWC.RadioButton
        TextBox             = SWC.TextBox
        TextBlock           = SWC.TextBlock
        Border              = SWC.Border
        ScrollViewer        = SWC.ScrollViewer
        GroupBox            = SWC.GroupBox
        GridSplitter        = SWC.GridSplitter
        Orientation         = SWC.Orientation
        SolidColorBrush     = SWM.SolidColorBrush
        Color               = SWM.Color
        Brushes             = SWM.Brushes
        FontFamily          = SWM.FontFamily
        Cursors             = SWI.Cursors
        _MBEH               = SWI.MouseButtonEventHandler
        _MEH                = SWI.MouseEventHandler
        import System.Windows.Controls as _SWC2
        import System.Windows.Media as _SWM2
        import System.Windows as _SW2

        def make_btn(text, bg=None, fg=None, width=None, height=28, bold=False):
            b = _SWC2.Button()
            b.Content = text
            b.Background = bg or _SWM2.SolidColorBrush(_SWM2.Color.FromRgb(22,160,133))
            b.Foreground = fg or _SWM2.Brushes.White
            b.BorderThickness = _SW2.Thickness(0)
            b.Padding = _SW2.Thickness(12,0,12,0)
            b.Height = height
            b.FontFamily = _SWM2.FontFamily("Segoe UI")
            b.FontSize = 12
            b.Cursor = _SW2.Input.Cursors.Hand
            if bold: b.FontWeight = _SW2.FontWeights.Bold
            if width: b.Width = width
            return b

        def make_lbl(text, bold=False, size=12, fg=None):
            l = _SWC2.TextBlock()
            l.Text = text
            l.FontFamily = _SWM2.FontFamily("Segoe UI")
            l.FontSize = size
            l.Foreground = fg or _SWM2.SolidColorBrush(_SWM2.Color.FromRgb(30,30,38))
            l.VerticalAlignment = _SW2.VerticalAlignment.Center
            if bold: l.FontWeight = _SW2.FontWeights.Bold
            return l

        def make_txt(val="", width=None, height=26, readonly=False):
            t = _SWC2.TextBox()
            t.Text = val
            t.Height = height
            t.FontFamily = _SWM2.FontFamily("Segoe UI")
            t.FontSize = 12
            t.Padding = _SW2.Thickness(6,2,6,2)
            t.Background = _SWM2.Brushes.White
            t.Foreground = _SWM2.SolidColorBrush(_SWM2.Color.FromRgb(30,30,38))
            t.IsReadOnly = readonly
            if width: t.Width = width
            return t
        root = Grid()
        root.RowDefinitions.Add(RowDefinition(Height=GridLength(48)))           # toolbar
        root.RowDefinitions.Add(RowDefinition(Height=GridLength(1, GridUnitType.Auto)))  # filter bar
        root.RowDefinitions.Add(RowDefinition(Height=GridLength(34)))           # col headers
        root.RowDefinitions.Add(RowDefinition(Height=GridLength(1, GridUnitType.Star)))  # listbox
        root.RowDefinitions.Add(RowDefinition(Height=GridLength(1, GridUnitType.Auto)))  # profile picker
        root.RowDefinitions.Add(RowDefinition(Height=GridLength(50)))           # footer

        # Toolbar
        tb = Border()
        tb.Background = BR_WHITE; tb.BorderBrush = BR_BORDER
        tb.BorderThickness = Thickness(0,0,0,1)
        tb.SetValue(Grid.RowProperty, 0)

        tr = StackPanel(); tr.Orientation = Orientation.Horizontal
        tr.VerticalAlignment = VerticalAlignment.Center
        tr.Margin = Thickness(12,0,12,0)

        self.rad_show_sheets = SWC.RadioButton()
        self.rad_show_sheets.Content   = "Sheets"
        self.rad_show_sheets.IsChecked = True
        self.rad_show_sheets.GroupName = "sp_mode"
        self.rad_show_sheets.FontFamily = SWM.FontFamily("Segoe UI")
        self.rad_show_sheets.FontSize  = 12
        self.rad_show_sheets.Foreground = BR_TEXT
        self.rad_show_sheets.Margin    = SW.Thickness(0,3,0,3)
        self.rad_show_views = SWC.RadioButton()
        self.rad_show_views.Content   = "Views"
        self.rad_show_views.IsChecked = False
        self.rad_show_views.GroupName = "sp_mode"
        self.rad_show_views.FontFamily = SWM.FontFamily("Segoe UI")
        self.rad_show_views.FontSize  = 12
        self.rad_show_views.Foreground = BR_TEXT
        self.rad_show_views.Margin    = SW.Thickness(0,3,0,3)
        self.rad_show_sheets.Margin = Thickness(0,0,8,0)
        self.rad_show_sheets.Click += lambda s,e: self._load_rows()
        self.rad_show_views.Click  += lambda s,e: self._load_rows()
        tr.Children.Add(make_lbl("Show:", bold=True, size=12))
        tr.Children.Add(self.rad_show_sheets)
        tr.Children.Add(self.rad_show_views)

        s1 = Border(); s1.Background=BR_BORDER; s1.Width=1; s1.Margin=Thickness(10,8,10,8)
        tr.Children.Add(s1)
        tr.Children.Add(make_lbl("Search:", size=12))
        self.txt_search = make_txt(width=220, height=28)
        self.txt_search.Margin = Thickness(6,0,0,0)
        self.txt_search.TextChanged += lambda s,e: self._load_rows(s.Text)
        tr.Children.Add(self.txt_search)

        s2 = Border(); s2.Background=BR_BORDER; s2.Width=1; s2.Margin=Thickness(10,8,10,8)
        tr.Children.Add(s2)
        ba = make_btn("Select All", bg=BR_PANEL, fg=BR_TEXT, width=90, height=30)
        bn = make_btn("Clear All",  bg=BR_PANEL, fg=BR_TEXT, width=80, height=30)
        ba.Margin=Thickness(0,0,6,0)
        ba.Click += lambda s,e: self._check_all(True)
        bn.Click += lambda s,e: self._check_all(False)
        tr.Children.Add(ba); tr.Children.Add(bn)


        tb.Child = tr
        root.Children.Add(tb)

        # Header
        hdr = Border()
        hdr.Background = SolidColorBrush(Color.FromRgb(238,240,244))
        hdr.BorderBrush = BR_BORDER; hdr.BorderThickness = Thickness(0,0,0,1)
        hdr.SetValue(Grid.RowProperty, 1)
        hdr_g = Grid()
        self.hdr_col_defs = []
        for w in self.col_widths:
            cd = ColumnDefinition(); cd.Width = GridLength(w); cd.MinWidth = 20
            hdr_g.ColumnDefinitions.Add(cd); self.hdr_col_defs.append(cd)
        for i, txt in enumerate(["","Sheet Number","Sheet Name","Revision","Size","Spool Tag"]):
            t = SWC.TextBlock(); t.Text=txt; t.FontWeight=SW.FontWeights.Bold; t.FontSize=11
            t.VerticalAlignment=SW.VerticalAlignment.Center; t.Margin=SW.Thickness(8,0,0,0)
            SWC.Grid.SetColumn(t,i); hdr_g.Children.Add(t)
        hdr.Child = hdr_g
        root.Children.Add(hdr)

        # ListBox
        import System.Windows.Controls as _SWC
        self.list_box = _SWC.ListBox()
        self.list_box.SetValue(Grid.RowProperty, 3)
        self.list_box.SelectionMode = _SWC.SelectionMode.Multiple
        self.list_box.IsEnabled = True
        _LBI = _SWC.ListBoxItem
        from System.Windows import Style, Setter, HorizontalAlignment as _HA
        item_style = Style(_LBI)
        item_style.Setters.Add(Setter(_LBI.HorizontalContentAlignmentProperty, _HA.Stretch))
        self.list_box.ItemContainerStyle = item_style
        self.list_box.SelectionChanged += self._on_sel
        self.list_box.PreviewMouseLeftButtonDown += _MBEH(self._on_mouse_down)
        root.Children.Add(self.list_box)

        # Profile picker panel
        prof_border = Border()
        prof_border.Background = SWM.SolidColorBrush(SWM.Color.FromRgb(245,246,248))
        prof_border.BorderBrush = SWM.SolidColorBrush(SWM.Color.FromRgb(210,212,220))
        prof_border.BorderThickness = SW.Thickness(0,1,0,1)
        prof_border.Padding = SW.Thickness(12,8,12,8)
        prof_border.SetValue(SWC.Grid.RowProperty, 4)

        prof_row = SWC.StackPanel(); prof_row.Orientation = SWC.Orientation.Horizontal

        prof_lbl = SWC.TextBlock(); prof_lbl.Text = "Use Export Profile:"
        prof_lbl.FontSize = 11; prof_lbl.FontWeight = SW.FontWeights.Bold
        prof_lbl.VerticalAlignment = SW.VerticalAlignment.Center
        prof_lbl.Margin = SW.Thickness(0,0,10,0)
        prof_row.Children.Add(prof_lbl)

        self.sched_cmb_profile = SWC.ComboBox()
        self.sched_cmb_profile.Height = 28; self.sched_cmb_profile.Width = 220
        self.sched_cmb_profile.FontSize = 12
        self.sched_cmb_profile.Margin = SW.Thickness(0,0,10,0)
        # Populate with available profiles
        import os as _osp, re as _rsp
        _prof_base = _osp.path.join(_osp.environ.get("APPDATA", _osp.path.expanduser("~")), "pyRevit_SheetsExporter")
        self.sched_cmb_profile.Items.Add("(none - use main settings)")
        _prof_names = []
        if _osp.path.exists(_prof_base):
            _prof_names = [f[:-5] for f in _osp.listdir(_prof_base) if f.endswith(".json") and not f.startswith("_")]
            for _pn in sorted(_prof_names): self.sched_cmb_profile.Items.Add(_pn)
        self.sched_cmb_profile.SelectedIndex = 0
        prof_row.Children.Add(self.sched_cmb_profile)

        prof_note = SWC.TextBlock()
        prof_note.Text = "Format, filename template and all settings will be taken from the selected profile."
        prof_note.FontSize = 10
        prof_note.Foreground = SWM.SolidColorBrush(SWM.Color.FromRgb(120,120,135))
        prof_note.VerticalAlignment = SW.VerticalAlignment.Center
        prof_row.Children.Add(prof_note)

        prof_border.Child = prof_row
        root.Children.Add(prof_border)

        # Footer
        foot = Border()
        foot.Background=BR_WHITE; foot.BorderBrush=BR_BORDER
        foot.BorderThickness=Thickness(0,1,0,0)
        foot.SetValue(Grid.RowProperty, 5)
        foot_dp = DockPanel(); foot_dp.Margin=Thickness(14,0,14,0)
        foot_dp.LastChildFill=False; foot.Child=foot_dp
        self.lbl_count = make_lbl("0 selected", fg=BR_SUBTEXT, size=11)
        self.lbl_count.VerticalAlignment = VerticalAlignment.Center
        DockPanel.SetDock(self.lbl_count, Dock.Left)
        foot_dp.Children.Add(self.lbl_count)

        # Output folder in footer
        folder_dp = DockPanel(); folder_dp.LastChildFill=True
        folder_dp.Margin = Thickness(0,6,16,0)
        btn_br = make_btn("Browse...", bg=BR_PANEL, fg=BR_TEXT, width=90, height=28)
        self.txt_folder = make_txt(
            self.parent_window.txt_folder.Text if hasattr(self.parent_window, "txt_folder") else "",
            readonly=True)
        def browse_folder(s, e):
            import clr as _clr
            _clr.AddReference("System.Windows.Forms")
            from System.Windows.Forms import FolderBrowserDialog, DialogResult
            dlg = FolderBrowserDialog()
            if dlg.ShowDialog() == DialogResult.OK:
                self.txt_folder.Text = dlg.SelectedPath
        btn_br.Click += browse_folder
        DockPanel.SetDock(btn_br, Dock.Right)
        folder_dp.Children.Add(btn_br)
        folder_dp.Children.Add(self.txt_folder)
        folder_lbl = make_lbl("Output Folder:", bold=True, size=11)
        folder_lbl.Margin = Thickness(0,0,6,0)
        folder_lbl.VerticalAlignment = VerticalAlignment.Center
        frow = DockPanel(); frow.LastChildFill = True
        DockPanel.SetDock(folder_lbl, Dock.Left)
        frow.Children.Add(folder_lbl)
        frow.Children.Add(folder_dp)
        DockPanel.SetDock(frow, Dock.Left)
        foot_dp.Children.Add(frow)

        btn_ok = make_btn("Save Schedule Config", width=160, height=34, bold=True)
        btn_cancel = make_btn("Cancel", bg=BR_GREY_BTN, fg=BR_TEXT, width=80, height=34)
        btn_ok.Margin = Thickness(6,0,0,0)
        btn_ok.Click += self._on_ok
        btn_cancel.Click += lambda s,e: self.Close()
        DockPanel.SetDock(btn_ok, Dock.Right)
        DockPanel.SetDock(btn_cancel, Dock.Right)
        foot_dp.Children.Add(btn_ok)
        foot_dp.Children.Add(btn_cancel)
        root.Children.Add(foot)
        self.Content = root

    def _on_sel(self, sender, e):
        if self.is_updating_selection: return
        for item in self.list_box.SelectedItems:
            if hasattr(item, "Tag") and item.Tag:
                self.last_clicked_row = item.Tag

    def _on_mouse_down(self, sender, e):
        from System.Windows.Input import Keyboard, Key
        from System.Windows.Media import VisualTreeHelper
        from System.Windows.Controls import ListBoxItem
        is_ctrl  = Keyboard.IsKeyDown(Key.LeftCtrl) or Keyboard.IsKeyDown(Key.RightCtrl)
        is_shift = Keyboard.IsKeyDown(Key.LeftShift) or Keyboard.IsKeyDown(Key.RightShift)
        hit = VisualTreeHelper.HitTest(self.list_box, e.GetPosition(self.list_box))
        clicked_item = clicked_lbi = None
        if hit:
            el = hit.VisualHit
            while el and not isinstance(el, ListBoxItem):
                el = VisualTreeHelper.GetParent(el)
            if el:
                clicked_lbi = el; clicked_item = el.Content
        if clicked_item is None: return
        try:
            pos = e.GetPosition(clicked_lbi)
            if pos.X <= 40:
                sr = clicked_item.Tag if hasattr(clicked_item, "Tag") else None
                if sr:
                    new_state = not sr.IsChecked
                    cur_sel = [self.list_box.Items[i].Tag
                               for i in range(self.list_box.Items.Count)
                               if self.list_box.Items[i] in self.list_box.SelectedItems
                               and hasattr(self.list_box.Items[i], "Tag")]
                    targets = cur_sel if sr in cur_sel else [sr]
                    for r in targets:
                        r.IsChecked = new_state
                        if hasattr(r, "_chk_ref"): r._chk_ref.IsChecked = new_state
                    self._upd_count()
                e.Handled = True; return
        except: pass
        if not is_ctrl and not is_shift:
            self.is_updating_selection = True
            self.list_box.SelectedItems.Clear()
            self.is_updating_selection = False

    def _build_row(self, sr, idx):
        import System.Windows.Controls as _SWC6
        import System.Windows.Media as _SWM6
        import System.Windows as _SW6
        rg = _SWC6.Grid()
        rg.Tag = sr
        rg.Background = _SWM6.SolidColorBrush(_SWM6.Color.FromArgb(1,255,255,255))
        rg.MinHeight = 26
        for w in self.col_widths:
            cd = _SWC6.ColumnDefinition(); cd.Width = _SW6.GridLength(w)
            rg.ColumnDefinitions.Add(cd)
        chk = _SWC6.CheckBox()
        chk.IsChecked = sr.IsChecked
        chk.VerticalAlignment = _SW6.VerticalAlignment.Center
        chk.HorizontalAlignment = _SW6.HorizontalAlignment.Center
        chk.IsHitTestVisible = False; chk.Tag = sr; sr._chk_ref = chk
        _SWC6.Grid.SetColumn(chk, 0); rg.Children.Add(chk)
        for i, val in enumerate([sr.SheetNumber, sr.SheetName, sr.Revision, sr.Size, sr.SpoolTag]):
            t = _SWC6.TextBlock(); t.Text = val or ""
            t.VerticalAlignment = _SW6.VerticalAlignment.Center
            t.Margin = _SW6.Thickness(8,0,0,0); t.FontSize = 12
            t.IsHitTestVisible = False
            _SWC6.Grid.SetColumn(t, i+1); rg.Children.Add(t)
        return rg

    def _load_rows(self, ft=""):
        import System.Windows as _SW
        ft = (ft or "").lower().strip()
        self.list_box.Items.Clear()
        self.sheet_rows = []; self.sheet_rows_list = []
        is_views = hasattr(self, "rad_show_views") and self.rad_show_views.IsChecked
        try:
            if is_views:
                for idx, v in enumerate(self.all_views):
                    try:
                        if ft and ft not in v.Name.lower() and ft not in str(v.ViewType).lower(): continue
                        vr = self._ViewRow.__new__(self._ViewRow)
                        vr.view=v; vr.IsChecked=False; vr.IsSelected=False
                        vr.SheetNumber=v.Name; vr.SheetName=str(v.ViewType).replace("ViewType.","")
                        vr.Revision=""; vr.Size=""; vr.SpoolTag=""; vr._chk_ref=None
                        self.sheet_rows.append(vr); self.sheet_rows_list.append(vr)
                        self.list_box.Items.Add(self._build_row(vr, idx))
                    except: pass
            else:
                for idx, s in enumerate(self.all_sheets):
                    try:
                        p = s.LookupParameter("SPOOL SHEET HCC IFF TAG")
                        tag = (p.AsString() or "").lower() if p and p.HasValue else ""
                        if ft and ft not in s.SheetNumber.lower() and ft not in s.Name.lower() and ft not in tag: continue
                        sr = self._SheetRow.__new__(self._SheetRow)
                        sr.sheet=s; sr.IsChecked=False; sr.IsSelected=False
                        sr.SheetNumber=s.SheetNumber; sr.SheetName=s.Name
                        p2 = s.LookupParameter("Current Revision")
                        sr.Revision = p2.AsString() if p2 and p2.HasValue else ""
                        sr.Size="Auto"
                        p3 = s.LookupParameter("SPOOL SHEET HCC IFF TAG")
                        sr.SpoolTag = p3.AsString() if p3 and p3.HasValue else ""
                        sr._chk_ref=None
                        self.sheet_rows.append(sr); self.sheet_rows_list.append(sr)
                        self.list_box.Items.Add(self._build_row(sr, idx))
                    except: pass
        except Exception as _le:
            _SW.MessageBox.Show("Load rows error: {}\nSheets: {}  Views: {}".format(
                str(_le), len(self.all_sheets), len(self.all_views)), "Scheduler List Error")
        self._upd_count()

    def _check_all(self, state):
        for sr in self.sheet_rows:
            sr.IsChecked = state
            if hasattr(sr, "_chk_ref"): sr._chk_ref.IsChecked = state
        self._upd_count()

    def _upd_count(self):
        checked = [r for r in self.sheet_rows if r.IsChecked]
        self.lbl_count.Text = "{} selected.  Total: {}".format(len(checked), len(self.sheet_rows))

    def _on_ok(self, sender, e):
        self.selected_rows = [r for r in self.sheet_rows if r.IsChecked]
        self.output_folder = self.txt_folder.Text if hasattr(self, "txt_folder") else ""
        if not self.selected_rows:
            import System.Windows as _SW
            _SW.MessageBox.Show("Check at least one sheet or view.", "Nothing Selected")
            return
        # Load settings from selected profile
        selected_prof = str(self.sched_cmb_profile.SelectedItem or "")
        if selected_prof and selected_prof != "(none - use main settings)":
            import os as _op, json as _jp, re as _rp
            _base = _op.path.join(_op.environ.get("APPDATA", _op.path.expanduser("~")), "pyRevit_SheetsExporter")
            _fname = _rp.sub(r'[/*?:"<>|]', "-", selected_prof).strip()
            _path = _op.path.join(_base, _fname + ".json")
            if _op.path.exists(_path):
                with open(_path) as _f: self.format_settings = _jp.load(_f)
                self.format_settings["_profile_name"] = selected_prof
            else:
                self.format_settings = None
        else:
            self.format_settings = None  # Use main window settings
        self.DialogResult = True
        self.Close()


# ─────────────────────────────────────────────────────────────────────────────
#  MAIN WINDOW
# ─────────────────────────────────────────────────────────────────────────────
class ProSheetsWindow(Window):

    PAGE_SEL = 0; PAGE_FMT = 1; PAGE_CRE = 2

    def __init__(self, handler, event, preselected=None, preselected_views=None):
        self.export_handler        = handler
        self.export_event          = event
        self.all_sheets            = get_sheets()
        self.sheet_rows            = []
        self.sort_col              = None
        # Store column widths - synced between header and rows
        self.col_widths = [40, 180, 400, 80, 100, 200]  # pixel widths
        self.sort_asc              = True
        self.sheet_rows_list       = []
        self.selected_rows         = []
        self.last_clicked_row      = None
        self.is_mouse_dragging     = False
        self.drag_start_item       = None
        self.is_updating_selection = False
        self.cur_page              = self.PAGE_SEL
        self.preselected             = preselected
        self.preselected_views       = preselected_views
        self.active_preselection     = preselected is not None or preselected_views is not None

        self.Title      = "Sheet Exporter"
        self.Width      = 1280
        self.Height     = 860
        self.MinWidth   = 1000
        self.MinHeight  = 650
        self.Background = BR_BG
        self.FontFamily = FONT_UI
        self.WindowStartupLocation = SW.WindowStartupLocation.CenterScreen

        # Store module refs on self so event handlers can access them
        import os as _os_ref
        import re as _re_ref
        import datetime as _dt_ref
        self._os = _os_ref
        self._re = _re_ref
        self._dt = _dt_ref
        self._SheetRow = SheetRow
        self._make_btn = make_btn
        self._make_lbl = make_lbl
        self._make_txt = make_txt
        self._make_rad = make_rad
        self._make_chk = make_chk
        self._vsp      = vsp
        self._hsp      = hsp
        self._sched_timer   = None
        self._SchedPickerWindow = SchedPickerWindow
        self._sched_running = False
        self._sched_sheets     = []  # SheetRow/ViewRow objects for scheduled export
        self._sched_out_folder    = ""   # output folder for scheduled export
        self._sched_fmt_settings  = None  # None = use main format settings
        self._sched_h = 8
        self._sched_m = 0
        self._ViewRow  = ViewRow
        self.all_views = get_views()
        # Load saved column widths if available
        saved_widths = self._load_col_widths()
        if saved_widths and len(saved_widths) == len(self.col_widths):
            self.col_widths = saved_widths
        self._build_ui()
        # Auto-switch to views mode if views were pre-selected
        if self.preselected_views and not self.preselected:
            try:
                self.rad_show_views.IsChecked  = True
                self.rad_show_sheets.IsChecked = False
            except: pass
        self._load_rows()
        self._go(self.PAGE_SEL)
        self._refresh_prof_cmb()
        # Auto-load last used profile - UI is fully built now
        last = self._load_last_profile()
        if last:
            try: self.cmb_profile.SelectedItem = last
            except: pass
        # Restore scheduler state
        self._restore_scheduler_state()



    # ─────────────────────────────────────────────────────────────────────────
    #  UI BUILD
    # ─────────────────────────────────────────────────────────────────────────
    def _build_ui(self):
        from System.Windows.Controls import RowDefinition as RD

        # Root grid - exact spool manager pattern using Grid.SetRow()
        root = Grid()
        root.RowDefinitions.Add(RD())  # 0 header
        root.RowDefinitions.Add(RD())  # 1 tabs
        root.RowDefinitions.Add(RD())  # 2 content
        root.RowDefinitions.Add(RD())  # 3 footer
        root.RowDefinitions[0].Height = GridLength(46)
        root.RowDefinitions[1].Height = GridLength(44)
        root.RowDefinitions[2].Height = GridLength(1, GridUnitType.Star)
        root.RowDefinitions[3].Height = GridLength(50)

        # ── Header ────────────────────────────────────────────────────────────
        hdr = Border()
        hdr.Background = BR_HEADER
        hdr_dp = DockPanel()
        hdr_dp.LastChildFill     = True
        hdr_dp.VerticalAlignment = VerticalAlignment.Center
        hdr_dp.Margin            = Thickness(14, 0, 14, 0)
        hdr.Child = hdr_dp

        prof_row = StackPanel()
        prof_row.Orientation       = Orientation.Horizontal
        prof_row.VerticalAlignment = VerticalAlignment.Center
        prof_lbl = make_lbl("Profile", size=11,
                             fg=SolidColorBrush(Color.FromRgb(150,152,165)))
        prof_lbl.Margin = Thickness(0,0,6,0)
        prof_row.Children.Add(prof_lbl)
        self.cmb_profile = make_cmb(["(none)"], width=180)
        self.cmb_profile.Margin = Thickness(0,0,6,0)
        self.cmb_profile.SelectionChanged += self._on_profile_changed
        prof_row.Children.Add(self.cmb_profile)
        btn_load = make_btn("Load", width=55, height=30)
        btn_save = make_btn("Save", bg=BR_GREEN, width=55, height=30)
        btn_del  = make_btn("Del",  bg=BR_RED,   width=45, height=30)
        btn_load.Margin = Thickness(0,0,4,0)
        btn_save.Margin = Thickness(0,0,4,0)
        btn_load.Click += self._prof_load_or_new
        btn_save.Click += self._prof_save
        btn_del.Click  += self._prof_del
        prof_row.Children.Add(btn_load)
        prof_row.Children.Add(btn_save)
        prof_row.Children.Add(btn_del)
        DockPanel.SetDock(prof_row, Dock.Right)
        hdr_dp.Children.Add(prof_row)

        left = StackPanel()
        left.Orientation       = Orientation.Horizontal
        left.VerticalAlignment = VerticalAlignment.Center
        logo = TextBlock()
        logo.Text = "Sheet Exporter"; logo.FontFamily = FONT_UI; logo.FontSize = 16
        logo.FontWeight = FontWeights.Bold; logo.FontStyle = FontStyles.Italic
        logo.Foreground = BR_ACCENT; logo.Margin = Thickness(0,0,8,0)
        logo.VerticalAlignment = VerticalAlignment.Center
        left.Children.Add(logo)
        sub = make_lbl("", size=11,
                        fg=SolidColorBrush(Color.FromRgb(160,162,170)))
        left.Children.Add(sub)
        hdr_dp.Children.Add(left)

        Grid.SetRow(hdr, 0)
        root.Children.Add(hdr)

        # ── Tab bar ───────────────────────────────────────────────────────────
        tab_border = Border()
        tab_border.Background      = BR_WHITE
        tab_border.BorderBrush     = BR_BORDER
        tab_border.BorderThickness = Thickness(0, 0, 0, 1)

        tab_row = StackPanel()
        tab_row.Orientation       = Orientation.Horizontal
        tab_row.VerticalAlignment = VerticalAlignment.Stretch
        self.tab_btns = []
        for i, name in enumerate(["Selection", "Format", "Create"]):
            b = Button()
            b.Content = name; b.Background = BR_WHITE
            b.BorderThickness = Thickness(0); b.FontFamily = FONT_UI
            b.FontSize = 13; b.Foreground = BR_SUBTEXT
            b.Padding = Thickness(16,0,16,0); b.Height = 44
            b.Cursor = Cursors.Hand; b.Tag = i
            b.Click += self._on_tab
            tab_row.Children.Add(b)
            self.tab_btns.append(b)
        tab_border.Child = tab_row

        Grid.SetRow(tab_border, 1)
        root.Children.Add(tab_border)

        # ── Content ───────────────────────────────────────────────────────────
        self.content_border = Border()
        Grid.SetRow(self.content_border, 2)
        root.Children.Add(self.content_border)

        self.page_sel = self._build_page_sel()
        self.page_fmt = self._build_page_fmt()
        self.page_cre = self._build_page_cre()

        # ── Footer ────────────────────────────────────────────────────────────
        foot = Border()
        foot.Background = BR_WHITE; foot.BorderBrush = BR_BORDER
        foot.BorderThickness = Thickness(0, 1, 0, 0)

        foot_dp = DockPanel()
        foot_dp.Margin = Thickness(14, 0, 14, 0)
        foot_dp.LastChildFill = False
        foot.Child = foot_dp

        self.lbl_count = make_lbl("0 sheets selected.  Total: 0",
                                   fg=BR_SUBTEXT, size=11)
        self.lbl_count.VerticalAlignment = VerticalAlignment.Center
        DockPanel.SetDock(self.lbl_count, Dock.Left)
        foot_dp.Children.Add(self.lbl_count)

        self.btn_export = make_btn("Export", width=110, height=34, bold=True)
        self.btn_next   = make_btn("Next",   width=110, height=34, bold=True)
        self.btn_back   = make_btn("Back",   bg=BR_GREY_BTN, fg=BR_TEXT,
                                   width=110, height=34)
        self.btn_export.Visibility = Visibility.Collapsed
        self.btn_back.Visibility   = Visibility.Collapsed

        self.btn_export.Click += self._on_export
        self.btn_next.Click   += self._on_next
        self.btn_back.Click += self._on_back

        self.btn_export.Margin = Thickness(6,0,0,0)
        self.btn_back.Margin   = Thickness(6,0,0,0)
        DockPanel.SetDock(self.btn_export, Dock.Right)
        DockPanel.SetDock(self.btn_next,   Dock.Right)
        DockPanel.SetDock(self.btn_back,   Dock.Right)
        foot_dp.Children.Add(self.btn_export)
        foot_dp.Children.Add(self.btn_next)
        foot_dp.Children.Add(self.btn_back)

        Grid.SetRow(foot, 3)
        root.Children.Add(foot)

        self.Content = root


    # ─────────────────────────────────────────────────────────────────────────
    #  PAGE: SELECTION
    # ─────────────────────────────────────────────────────────────────────────
    def _build_page_sel(self):
        from System.Windows import Style, Setter
        from System.Windows.Controls import ListBoxItem
        from System.Windows.Input import MouseButtonEventHandler as MBEH
        from System.Windows.Input import MouseEventHandler as MEH

        root = Grid()
        root.RowDefinitions.Add(RowDefinition(Height=GridLength(48)))
        root.RowDefinitions.Add(RowDefinition(Height=GridLength(1, GridUnitType.Auto)))
        root.RowDefinitions.Add(RowDefinition(Height=GridLength(34)))
        root.RowDefinitions.Add(RowDefinition(Height=GridLength(1, GridUnitType.Star)))

        # ── Toolbar ───────────────────────────────────────────────────────────
        tb = Border()
        tb.Background = BR_WHITE; tb.BorderBrush = BR_BORDER
        tb.BorderThickness = Thickness(0,0,0,1)
        tb.SetValue(Grid.RowProperty, 0)
        root.Children.Add(tb)

        tr = StackPanel(); tr.Orientation = Orientation.Horizontal
        tr.VerticalAlignment = VerticalAlignment.Center
        tr.Margin = Thickness(12,0,12,0)
        tr.Children.Add(make_lbl("Show:", bold=True, size=13))
        self.rad_show_sheets = make_rad("Sheets", checked=True, group="show_mode")
        self.rad_show_views  = make_rad("Views",  checked=False, group="show_mode")
        self.rad_show_sheets.Margin = Thickness(6,0,6,0)
        self.rad_show_views.Margin  = Thickness(0,0,0,0)
        self.rad_show_sheets.Click += self._on_show_mode_changed
        self.rad_show_views.Click  += self._on_show_mode_changed
        tr.Children.Add(self.rad_show_sheets)
        tr.Children.Add(self.rad_show_views)
        s1 = Border(); s1.Background=BR_BORDER; s1.Width=1; s1.Margin=Thickness(10,8,10,8)
        tr.Children.Add(s1)
        tr.Children.Add(make_lbl("Search:", size=12))
        self.txt_search = make_txt(width=260, height=28)
        self.txt_search.Margin = Thickness(6,0,0,0)
        self.txt_search.TextChanged += self._on_search
        tr.Children.Add(self.txt_search)
        s2 = Border(); s2.Background=BR_BORDER; s2.Width=1; s2.Margin=Thickness(10,8,10,8)
        tr.Children.Add(s2)
        ba = make_btn("Select All", bg=BR_PANEL, fg=BR_TEXT, width=90, height=30)
        bn = make_btn("Clear All",  bg=BR_PANEL, fg=BR_TEXT, width=80, height=30)
        ba.Margin=Thickness(0,0,6,0)
        ba.Click += lambda s,e: self._check_all(True)
        bn.Click += lambda s,e: self._check_all(False)
        tr.Children.Add(ba); tr.Children.Add(bn)

        tb.Child = tr

        # ── Filter indicator bar ──────────────────────────────────────────────
        filter_bar = Border()
        filter_bar.SetValue(Grid.RowProperty, 1)
        filter_bar.Background = SolidColorBrush(Color.FromRgb(255, 243, 220))
        filter_bar.BorderBrush = SolidColorBrush(Color.FromRgb(180, 80, 30))
        filter_bar.BorderThickness = Thickness(0, 0, 0, 1)
        filter_bar.Padding = Thickness(12, 4, 12, 4)
        filter_dp = DockPanel(); filter_dp.LastChildFill = False
        self.btn_clear_filter = make_btn(
            "X  Clear Selection Filter",
            bg=SolidColorBrush(Color.FromRgb(180,80,30)), fg=BR_WHITE, height=24)
        self.btn_clear_filter.FontSize = 11
        self.btn_clear_filter.Click += lambda s,e: self._on_clear_filter(s, e)
        DockPanel.SetDock(self.btn_clear_filter, Dock.Right)
        filter_dp.Children.Add(self.btn_clear_filter)
        self.lbl_filter_info = make_lbl("Selection filter active", size=11,
            fg=SolidColorBrush(Color.FromRgb(180,80,30)))
        filter_dp.Children.Add(self.lbl_filter_info)
        filter_bar.Child = filter_dp
        has_filter = bool(self.preselected) or bool(self.preselected_views)
        filter_bar.Visibility = Visibility.Visible if has_filter else Visibility.Collapsed
        self.filter_bar = filter_bar
        root.Children.Add(filter_bar)

        # ── Column headers with sort ──────────────────────────────────────────
        # Store shared column widths so rows and header stay in sync
        self.col_defs = [40, 180, 0, 80, 100, 200]  # 0 = Star
        hdr = Border()
        hdr.Background = SolidColorBrush(Color.FromRgb(238,240,244))
        hdr.BorderBrush = BR_BORDER; hdr.BorderThickness = Thickness(0,0,0,1)
        hdr.SetValue(Grid.RowProperty, 2)
        root.Children.Add(hdr)
        self.hdr_g = Grid()
        self.hdr_col_defs = []
        for w in self.col_widths:
            cd = ColumnDefinition()
            cd.Width = GridLength(w)
            cd.MinWidth = 20
            self.hdr_g.ColumnDefinitions.Add(cd)
            self.hdr_col_defs.append(cd)

        # Column headers - set dynamically but build with defaults now
        # _update_header_cols() will update based on mode
        cols = ["", "Sheet Number", "Sheet Name", "Revision", "Size", "Spool Tag"]
        self.sort_attrs = [None, "SheetNumber", "SheetName", "Revision", "Size", "SpoolTag"]
        self.hdr_labels = []
        for i, txt in enumerate(cols):
            btn = Button()
            btn.Content         = txt
            btn.Background      = Brushes.Transparent
            btn.BorderThickness = Thickness(0)
            btn.FontWeight      = FontWeights.Bold
            btn.FontSize        = 11
            btn.HorizontalContentAlignment = HorizontalAlignment.Left
            btn.Padding         = Thickness(8, 0, 4, 0)
            btn.Height          = 34
            btn.Cursor          = Cursors.Hand if txt else Cursors.Arrow
            btn.Tag             = i
            if txt:
                btn.Click += self._on_sort
            Grid.SetColumn(btn, i)
            self.hdr_g.Children.Add(btn)
            self.hdr_labels.append(btn)
        # Add GridSplitters for resizing
        for i in range(len(self.col_widths) - 1):
            sp = GridSplitter()
            sp.Width               = 4
            sp.Background          = BR_BORDER
            sp.HorizontalAlignment = HorizontalAlignment.Right
            sp.VerticalAlignment   = VerticalAlignment.Stretch
            sp.ResizeDirection     = SWC.GridResizeDirection.Columns
            sp.ResizeBehavior      = SWC.GridResizeBehavior.CurrentAndNext
            sp.Tag                 = i
            Grid.SetColumn(sp, i)
            self.hdr_g.Children.Add(sp)
            sp.DragCompleted += self._on_col_resize
        hdr.Child = self.hdr_g

        # ── ListBox — exact spool manager create_data_grid pattern ──────────────
        from System.Windows.Input import MouseButtonEventHandler as _MBEH
        from System.Windows.Input import MouseEventHandler as _MEH
        import System.Windows.Controls as _SWC

        self.list_box = _SWC.ListBox()
        self.list_box.SetValue(Grid.RowProperty, 3)
        self.list_box.Margin        = Thickness(0)
        self.list_box.SelectionMode = _SWC.SelectionMode.Multiple
        self.list_box.IsEnabled     = True

        from System.Windows import Style, Setter, HorizontalAlignment as _HA
        _LBI = _SWC.ListBoxItem
        item_style = Style(_LBI)
        item_style.Setters.Add(Setter(_LBI.HorizontalContentAlignmentProperty, _HA.Stretch))
        self.list_box.ItemContainerStyle = item_style

        self.list_box.SelectionChanged           += self._on_lb_sel_changed
        self.list_box.PreviewMouseLeftButtonDown += _MBEH(self._on_lb_mouse_down)
        self.list_box.MouseMove                  += _MEH(self._on_lb_mouse_move)
        self.list_box.PreviewMouseLeftButtonUp   += _MBEH(self._on_lb_mouse_up)

        root.Children.Add(self.list_box)
        return root

    def _on_clear_filter(self, s, e):
        self.active_preselection = False
        self.preselected         = None
        self.preselected_views   = None
        import System.Windows as _SWC3
        self.filter_bar.Visibility = _SWC3.Visibility.Collapsed
        self._load_rows(str(self.txt_search.Text) if self.txt_search.Text else "")

    def _on_col_resize(self, sender, e):
        for i, cd in enumerate(self.hdr_col_defs):
            self.col_widths[i] = cd.ActualWidth
        self._save_col_widths()
        self._load_rows(self.txt_search.Text if self.txt_search.Text else "")

    def _on_sort(self, sender, e):
        import System.Windows as _SW7
        col_idx = int(sender.Tag)
        attr = self.sort_attrs[col_idx]
        if not attr:
            return
        if self.sort_col == attr:
            self.sort_asc = not self.sort_asc
        else:
            self.sort_col = attr
            self.sort_asc = True
        # Update header labels to show sort indicator
        for i, btn in enumerate(self.hdr_labels):
            txt = ["", "Sheet Number", "Sheet Name", "Revision", "Size", "Spool Tag"][i]
            if self.sort_attrs[i] == attr:
                btn.Content = txt + (" ▲" if self.sort_asc else " ▼")
            else:
                btn.Content = txt
        # Sort all_sheets and reload
        attr_map = {
            "SheetNumber": lambda s: s.SheetNumber.lower(),
            "SheetName":   lambda s: s.Name.lower(),
            "Revision":    lambda s: (s.LookupParameter("Current Revision").AsString() or "").lower() if s.LookupParameter("Current Revision") and s.LookupParameter("Current Revision").HasValue else "",
            "Size":        lambda s: "",
            "SpoolTag":    lambda s: (s.LookupParameter("SPOOL SHEET HCC IFF TAG").AsString() or "").lower() if s.LookupParameter("SPOOL SHEET HCC IFF TAG") and s.LookupParameter("SPOOL SHEET HCC IFF TAG").HasValue else "",
        }
        self.all_sheets = sorted(self.all_sheets, key=attr_map.get(attr, lambda s: ""), reverse=not self.sort_asc)
        self._load_rows(self.txt_search.Text if self.txt_search.Text else "")

    def _build_row_panel(self, sr, idx):
        import System.Windows.Controls as _SWC6
        import System.Windows.Media as _SWM6
        import System.Windows as _SW6

        row_grid = _SWC6.Grid()
        row_grid.Tag        = sr
        row_grid.Background = _SWM6.SolidColorBrush(_SWM6.Color.FromArgb(1, 255, 255, 255))
        row_grid.MinHeight  = 26

        for w in self.col_widths:
            cd = _SWC6.ColumnDefinition()
            cd.Width = _SW6.GridLength(w)
            row_grid.ColumnDefinitions.Add(cd)

        chk = _SWC6.CheckBox()
        chk.IsChecked           = sr.IsChecked
        chk.VerticalAlignment   = _SW6.VerticalAlignment.Center
        chk.HorizontalAlignment = _SW6.HorizontalAlignment.Center
        chk.IsHitTestVisible    = False
        chk.Tag                 = sr
        sr._chk_ref             = chk
        _SWC6.Grid.SetColumn(chk, 0)
        row_grid.Children.Add(chk)

        for i, val in enumerate([sr.SheetNumber, sr.SheetName,
                                   sr.Revision, sr.Size, sr.SpoolTag]):
            t = _SWC6.TextBlock()
            t.Text              = val or ""
            t.VerticalAlignment = _SW6.VerticalAlignment.Center
            t.Margin            = _SW6.Thickness(8, 0, 0, 0)
            t.FontSize          = 12
            t.IsHitTestVisible  = False
            _SWC6.Grid.SetColumn(t, i + 1)
            row_grid.Children.Add(t)

        return row_grid

    def _on_lb_sel_changed(self, sender, e):
        if self.is_updating_selection:
            return
        self.selected_rows = []
        for item in self.list_box.SelectedItems:
            if hasattr(item, "Tag") and item.Tag:
                item.Tag.IsSelected = True
                self.selected_rows.append(item.Tag)
                self.last_clicked_row = item.Tag

    def _on_lb_mouse_down(self, sender, e):
        from System.Windows.Input import Keyboard, Key
        from System.Windows.Media import VisualTreeHelper
        from System.Windows.Controls import ListBoxItem
        is_ctrl  = Keyboard.IsKeyDown(Key.LeftCtrl)  or Keyboard.IsKeyDown(Key.RightCtrl)
        is_shift = Keyboard.IsKeyDown(Key.LeftShift) or Keyboard.IsKeyDown(Key.RightShift)

        # Get click position relative to listbox
        click_pos = e.GetPosition(self.list_box)

        hit = VisualTreeHelper.HitTest(self.list_box, click_pos)
        clicked_item = None
        clicked_lbi  = None
        if hit:
            el = hit.VisualHit
            while el is not None and not isinstance(el, ListBoxItem):
                el = VisualTreeHelper.GetParent(el)
            if el is not None and isinstance(el, ListBoxItem):
                clicked_lbi  = el
                clicked_item = el.Content
                self.is_mouse_dragging = True
                self.drag_start_item   = clicked_item

        if clicked_item is None:
            return

        # Detect if click is in checkbox column (first 40px of the row)
        if clicked_lbi is not None:
            try:
                pos_in_row = e.GetPosition(clicked_lbi)
                if pos_in_row.X <= 40:
                    sr = clicked_item.Tag if hasattr(clicked_item, "Tag") else None
                    if sr:
                        new_state = not sr.IsChecked
                        # Rebuild selected_rows fresh from ListBox.SelectedItems
                        # because _on_lb_sel_changed may not have run yet
                        current_selected = []
                        for i in range(self.list_box.Items.Count):
                            item = self.list_box.Items[i]
                            if item in self.list_box.SelectedItems:
                                if hasattr(item, "Tag") and item.Tag:
                                    current_selected.append(item.Tag)
                        # Apply to all highlighted, or just this one
                        if sr in current_selected and len(current_selected) > 0:
                            for r in current_selected:
                                r.IsChecked = new_state
                                if hasattr(r, "_chk_ref"):
                                    r._chk_ref.IsChecked = new_state
                        else:
                            sr.IsChecked = new_state
                            if hasattr(sr, "_chk_ref"):
                                sr._chk_ref.IsChecked = new_state
                        self._upd_count()
                    e.Handled = True
                    return
            except: pass

        # Row selection logic
        if is_shift and self.last_clicked_row and clicked_item:
            if hasattr(clicked_item, "Tag") and clicked_item.Tag:
                rows = self.sheet_rows_list
                try:
                    si = rows.index(self.last_clicked_row)
                    ei = rows.index(clicked_item.Tag)
                except ValueError:
                    return
                if si > ei: si, ei = ei, si
                self.is_updating_selection = True
                self.list_box.SelectedItems.Clear()
                for i in range(self.list_box.Items.Count):
                    item = self.list_box.Items[i]
                    if hasattr(item, "Tag") and item.Tag in rows[si:ei+1]:
                        self.list_box.SelectedItems.Add(item)
                self.is_updating_selection = False
                e.Handled = True

        elif not is_ctrl and not is_shift:
            self.is_updating_selection = True
            self.list_box.SelectedItems.Clear()
            self.is_updating_selection = False

    def _on_lb_mouse_up(self, sender, e):
        self.is_mouse_dragging = False
        self.drag_start_item   = None

    def _on_lb_mouse_move(self, sender, e):
        from System.Windows.Media import VisualTreeHelper
        from System.Windows.Controls import ListBoxItem
        if not self.is_mouse_dragging or not self.drag_start_item:
            return
        hit = VisualTreeHelper.HitTest(self.list_box, e.GetPosition(self.list_box))
        if not hit: return
        el = hit.VisualHit
        while el is not None and not isinstance(el, ListBoxItem):
            el = VisualTreeHelper.GetParent(el)
        if el is None: return
        cur = el.Content
        if cur and cur != self.drag_start_item:
            if hasattr(self.drag_start_item,"Tag") and hasattr(cur,"Tag"):
                rows = self.sheet_rows_list
                try:
                    si = rows.index(self.drag_start_item.Tag)
                    ei = rows.index(cur.Tag)
                except ValueError: return
                if si > ei: si, ei = ei, si
                self.is_updating_selection = True
                self.list_box.SelectedItems.Clear()
                for i in range(self.list_box.Items.Count):
                    item = self.list_box.Items[i]
                    if hasattr(item,"Tag") and item.Tag in rows[si:ei+1]:
                        self.list_box.SelectedItems.Add(item)
                self.is_updating_selection = False

    def _is_views_mode(self):
        return hasattr(self, "rad_show_views") and self.rad_show_views.IsChecked

    def _on_show_mode_changed(self, sender, e):
        self._update_header_cols()
        # Update default filename template based on mode (only if not restoring profile)
        if hasattr(self, "txt_template") and not getattr(self, "_applying_profile", False):
            if self._is_views_mode():
                self.txt_template.Text = "{Sheet Name}"
            else:
                self.txt_template.Text = "{Sheet Number}"
        self._load_rows(self.txt_search.Text if self.txt_search.Text else "")

    def _update_header_cols(self):
        if self._is_views_mode():
            labels = ["", "View Name", "View Type", "", "", ""]
        else:
            labels = ["", "Sheet Number", "Sheet Name", "Revision", "Size", "Spool Tag"]
        for i, btn in enumerate(self.hdr_labels):
            # preserve sort indicator if active
            base = labels[i]
            if self.sort_attrs[i] and self.sort_col == self.sort_attrs[i]:
                base += " ▲" if self.sort_asc else " ▼"
            btn.Content = base

    def _load_rows(self, ft=""):
        ft = ft.lower().strip() if ft else ""
        self.list_box.Items.Clear()
        self.sheet_rows      = []
        self.sheet_rows_list = []
        self.selected_rows   = []

        if self._is_views_mode():
            # Views mode
            import Autodesk.Revit.DB as _DB2
            for idx, v in enumerate(self.all_views):
                vtype = str(v.ViewType)
                # Pre-selection filter for views
                if self.active_preselection and self.preselected_views and not self.preselected:
                    if v.UniqueId not in self.preselected_views:
                        continue
                if ft and ft not in v.Name.lower() and ft not in vtype.lower():
                    continue
                try:
                    vr = self._ViewRow.__new__(self._ViewRow)
                    vr.view        = v
                    vr.IsChecked   = False
                    vr.IsSelected  = False
                    vr.SheetNumber = v.Name
                    vr.SheetName   = str(v.ViewType).replace("ViewType.", "")
                    vr.Revision    = ""
                    vr.Size        = ""
                    vr.SpoolTag    = ""
                    vr._chk_ref    = None
                    # Auto-check if pre-selected from project browser
                    if self.preselected_views and hasattr(v, "UniqueId"):
                        if v.UniqueId in self.preselected_views:
                            vr.IsChecked = True
                    self.sheet_rows.append(vr)
                    self.sheet_rows_list.append(vr)
                    self.list_box.Items.Add(self._build_row_panel(vr, idx))
                except: pass
        else:
            # Sheets mode
            for idx, s in enumerate(self.all_sheets):
                # Pre-selection filter
                if self.active_preselection and self.preselected:
                    if s.SheetNumber not in self.preselected:
                        continue
                spool_tag = ""
                p = s.LookupParameter("SPOOL SHEET HCC IFF TAG")
                if p and p.HasValue: spool_tag = (p.AsString() or "").lower()
                if ft and ft not in s.SheetNumber.lower() and ft not in s.Name.lower() and ft not in spool_tag:
                    continue
                try:
                    import Autodesk.Revit.DB as _DB2
                    sr = self._SheetRow.__new__(self._SheetRow)
                    sr.sheet       = s
                    sr.IsChecked   = False
                    sr.IsSelected  = False
                    sr.SheetNumber = s.SheetNumber
                    sr.SheetName   = s.Name
                    p = s.LookupParameter("Current Revision")
                    sr.Revision    = p.AsString() if p and p.HasValue else ""
                    sr.Size        = "Auto"
                    p2 = s.LookupParameter("SPOOL SHEET HCC IFF TAG")
                    sr.SpoolTag    = p2.AsString() if p2 and p2.HasValue else ""
                    sr._chk_ref    = None
                    self.sheet_rows.append(sr)
                    self.sheet_rows_list.append(sr)
                    self.list_box.Items.Add(self._build_row_panel(sr, idx))
                except: pass
        self._upd_count()
        # Update filter bar visibility
        try:
            import System.Windows as _SWV
            has_filter = bool(self.preselected) or bool(self.preselected_views)
            vis = _SWV.Visibility.Visible if has_filter else _SWV.Visibility.Collapsed
            self.filter_bar.Visibility = vis
        except: pass

    def _on_search(self, sender, e):
        self._load_rows(sender.Text)

    def _on_sel_changed(self, s, e):  pass
    def _on_cell_changed(self, s, e): pass
    def _on_mouse_down(self, s, e):   pass
    def _on_mouse_up(self, s, e):     pass
    def _on_mouse_move(self, s, e):   pass

    def _sync_checkboxes(self):
        for sr in self.sheet_rows:
            if hasattr(sr, "_chk_ref"):
                sr._chk_ref.IsChecked = sr.IsChecked

    def _check_all(self, state):
        for sr in self.sheet_rows:
            sr.IsChecked = bool(state)
            if hasattr(sr, "_chk_ref"):
                sr._chk_ref.IsChecked = bool(state)
        self._upd_count()

    def _get_checked(self):
        result = [r for r in self.sheet_rows if r.IsChecked]
        return result

    def _upd_count(self):
        checked = len(self._get_checked())
        self.lbl_count.Text = "{} sheets selected.  Total: {}".format(
            checked, len(self.sheet_rows))


    # ─────────────────────────────────────────────────────────────────────────
    #  PAGE: FORMAT
    # ─────────────────────────────────────────────────────────────────────────
    def _build_page_fmt(self):
        sv = ScrollViewer()
        sv.VerticalScrollBarVisibility = ScrollBarVisibility.Auto

        outer = vsp(margin=Thickness(16, 12, 16, 16))

        # Format toggles
        fmt_bar = Border()
        fmt_bar.Background      = BR_WHITE
        fmt_bar.BorderBrush     = BR_BORDER
        fmt_bar.BorderThickness = Thickness(0, 0, 0, 1)
        fmt_bar.Padding         = Thickness(10, 8, 10, 8)
        fmt_bar.Margin          = Thickness(0, 0, 0, 12)

        fmt_row = StackPanel(); fmt_row.Orientation = Orientation.Horizontal

        self.chk_pdf = CheckBox()
        self.chk_pdf.IsChecked  = True
        pdf_lbl = TextBlock(); pdf_lbl.Text="PDF"; pdf_lbl.FontSize=16
        pdf_lbl.FontWeight=FontWeights.Bold
        pdf_lbl.Foreground=SolidColorBrush(Color.FromRgb(180,130,30))
        pdf_lbl.VerticalAlignment = VerticalAlignment.Center
        self.chk_pdf.Content = pdf_lbl
        self.chk_pdf.Margin  = Thickness(0,0,16,0)
        fmt_row.Children.Add(self.chk_pdf)

        self.chk_dwg = CheckBox()
        self.chk_dwg.IsChecked = False
        dwg_lbl = TextBlock(); dwg_lbl.Text="DWG"; dwg_lbl.FontSize=16
        dwg_lbl.FontWeight=FontWeights.Bold
        dwg_lbl.Foreground=SolidColorBrush(Color.FromRgb(50,100,200))
        dwg_lbl.VerticalAlignment = VerticalAlignment.Center
        self.chk_dwg.Content = dwg_lbl
        self.chk_dwg.Margin = Thickness(0,0,16,0)
        fmt_row.Children.Add(self.chk_dwg)

        self.chk_nwc = CheckBox()
        self.chk_nwc.IsChecked = False
        nwc_lbl = TextBlock(); nwc_lbl.Text="NWC"; nwc_lbl.FontSize=16
        nwc_lbl.FontWeight=FontWeights.Bold
        nwc_lbl.Foreground=SolidColorBrush(Color.FromRgb(22,160,133))
        nwc_lbl.VerticalAlignment = VerticalAlignment.Center
        self.chk_nwc.Content = nwc_lbl
        self.chk_nwc.Click += self._on_nwc_toggled
        self.chk_pdf.Click += self._on_nwc_toggled
        self.chk_dwg.Click += self._on_nwc_toggled
        fmt_row.Children.Add(self.chk_nwc)
        fmt_bar.Child = fmt_row
        outer.Children.Add(fmt_bar)

        # NWC options panel — shown when NWC checked, hidden otherwise
        self.nwc_panel = self._build_nwc_panel()
        self.nwc_panel.Visibility = Visibility.Collapsed
        outer.Children.Add(self.nwc_panel)


        # Row 1 — Paper | Hidden Lines | Options
        row1 = StackPanel(); row1.Orientation=Orientation.Horizontal
        row1.Margin = Thickness(0,0,0,8)

        pp = vsp()
        self.rad_center = make_rad("Center", checked=True, group="paper")
        self.rad_offset = make_rad("Offset from corner",   group="paper")
        xy = hsp()
        xy.Children.Add(make_lbl("X = ", size=11))
        self.txt_ox = make_txt("0.00in", width=70); self.txt_ox.Margin=Thickness(2,0,10,0)
        xy.Children.Add(self.txt_ox)
        xy.Children.Add(make_lbl("Y = ", size=11))
        self.txt_oy = make_txt("0.00in", width=70); self.txt_oy.Margin=Thickness(2,0,0,0)
        xy.Children.Add(self.txt_oy)
        pp.Children.Add(self.rad_center); pp.Children.Add(self.rad_offset)
        pp.Children.Add(xy)
        row1.Children.Add(make_grp("Paper Placement", pp))

        hl = vsp()
        hl.Children.Add(make_lbl("Remove Lines Using", size=11, fg=BR_SUBTEXT))
        self.rad_vector = make_rad("Vector Processing", checked=True, group="hl")
        self.rad_raster = make_rad("Raster Processing",              group="hl")
        hl.Children.Add(self.rad_vector); hl.Children.Add(self.rad_raster)
        row1.Children.Add(make_grp("Hidden Line Views", hl))

        opts_g = Grid()
        opts_g.ColumnDefinitions.Add(ColumnDefinition())
        opts_g.ColumnDefinitions.Add(ColumnDefinition())
        self.chk_view_links   = make_chk("View links in blue (Color prints only)", True)
        self.chk_hide_ref     = make_chk("Hide ref/work planes",                   True)
        self.chk_hide_unref   = make_chk("Hide unreferenced view tags",             True)
        self.chk_hide_scope   = make_chk("Hide scope boxes",                        True)
        self.chk_hide_crop    = make_chk("Hide crop boundaries",                    True)
        self.chk_replace_half = make_chk("Replace halftone with thin lines",        False)
        self.chk_thin_lines = make_chk("Use Thin Lines", False)
        c0 = vsp(); c0.Children.Add(self.chk_view_links); c0.Children.Add(self.chk_hide_ref); c0.Children.Add(self.chk_hide_unref)
        c1 = vsp(); c1.Children.Add(self.chk_hide_scope); c1.Children.Add(self.chk_hide_crop); c1.Children.Add(self.chk_replace_half)
        c2 = vsp(); c2.Children.Add(self.chk_thin_lines)
        opts_g.ColumnDefinitions.Add(ColumnDefinition())
        c0.SetValue(Grid.ColumnProperty, 0)
        c1.SetValue(Grid.ColumnProperty, 1)
        c2.SetValue(Grid.ColumnProperty, 2)
        opts_g.Children.Add(c0); opts_g.Children.Add(c1); opts_g.Children.Add(c2)
        og = make_grp("Options", opts_g)
        og.HorizontalAlignment = HorizontalAlignment.Stretch
        row1.Children.Add(og)
        self._pdf_row1 = row1
        outer.Children.Add(row1)

        # Row 2 — Zoom | Appearance | File
        row2 = StackPanel(); row2.Orientation=Orientation.Horizontal
        row2.Margin = Thickness(0,0,0,8)

        zm = vsp()
        self.rad_fit  = make_rad("Fit to Page", checked=True, group="zoom")
        zm_r = hsp()
        self.rad_zoom = make_rad("Zoom", group="zoom")
        self.txt_zoom = make_txt("100", width=50); self.txt_zoom.Margin=Thickness(4,0,4,0)
        zm_r.Children.Add(self.rad_zoom); zm_r.Children.Add(self.txt_zoom)
        zm_r.Children.Add(make_lbl("% Size", size=11))
        zm.Children.Add(self.rad_fit); zm.Children.Add(zm_r)
        row2.Children.Add(make_grp("Zoom", zm))

        ap = vsp()
        ap.Children.Add(make_lbl("Raster Quality", size=11))
        self.cmb_quality = make_cmb(
            ["Draft (72 DPI)","Low (150 DPI)","Presentation (300 DPI)","High (600 DPI)"],
            selected=2, width=220)
        self.cmb_quality.Margin=Thickness(0,4,0,8)
        ap.Children.Add(self.cmb_quality)
        ap.Children.Add(make_lbl("Colors", size=11))
        self.cmb_colors = make_cmb(["Color","Black & White","Grayscale"], width=220)
        self.cmb_colors.Margin=Thickness(0,4,0,0)
        ap.Children.Add(self.cmb_colors)
        row2.Children.Add(make_grp("Appearance", ap))

        fi = vsp()
        self.rad_separate = make_rad("Create separate files", checked=True, group="file")
        self.rad_combine  = make_rad("Combine into a single file",          group="file")
        fi.Children.Add(self.rad_separate); fi.Children.Add(self.rad_combine)
        fi.Children.Add(make_lbl("Combined file name:", size=11))
        self.txt_combined = make_txt("Combined_Sheets", width=300)
        self.txt_combined.Margin=Thickness(0,4,0,8)
        fi.Children.Add(self.txt_combined)
        self.chk_keep_paper = make_chk("Keep Paper Size & Orientation", False)
        fi.Children.Add(self.chk_keep_paper)
        fg = make_grp("File", fi)
        fg.HorizontalAlignment = HorizontalAlignment.Stretch
        row2.Children.Add(fg)
        self._pdf_row2 = row2
        outer.Children.Add(row2)

        # Filename builder
        fn_sp = vsp()
        tok_r = StackPanel(); tok_r.Orientation=Orientation.Horizontal
        tok_r.Margin=Thickness(0,0,0,8)
        tok_r.Children.Add(make_lbl("Tokens:  ", size=11, fg=BR_SUBTEXT))
        for tok in ["{Sheet Number}","{Sheet Name}","{Current Revision}",
                    "{Date}","{Project Number}","{Project Name}"]:
            b = make_btn(tok, bg=BR_PANEL, fg=BR_TEXT, height=26)
            b.FontSize=11; b.Margin=Thickness(0,0,4,0)
            b.Tag=tok; b.Click += self._insert_token
            tok_r.Children.Add(b)
        fn_sp.Children.Add(tok_r)

        fn_r = StackPanel(); fn_r.Orientation=Orientation.Horizontal
        fn_r.Children.Add(make_lbl("Template:", bold=True, size=11))
        self.txt_template = make_txt("{Sheet Number}", width=300)
        self.txt_template.Margin=Thickness(6,0,12,0)
        fn_r.Children.Add(self.txt_template)
        fn_r.Children.Add(make_lbl("Sep:", size=11))
        self.cmb_sep = make_cmb(["-","_"," ","."], width=60)
        self.cmb_sep.Margin=Thickness(4,0,12,0)
        fn_r.Children.Add(self.cmb_sep)
        fn_r.Children.Add(make_lbl("Preview:", bold=True, size=11))
        self.lbl_preview = TextBlock()
        self.lbl_preview.FontSize=12; self.lbl_preview.FontWeight=FontWeights.Bold
        self.lbl_preview.Foreground=BR_ACCENT
        self.lbl_preview.VerticalAlignment=VerticalAlignment.Center
        self.lbl_preview.Margin=Thickness(6,0,8,0)
        fn_r.Children.Add(self.lbl_preview)
        btn_prev = make_btn("↻", bg=BR_PANEL, fg=BR_TEXT, width=32, height=26)
        btn_prev.Click += self._refresh_preview
        fn_r.Children.Add(btn_prev)
        fn_sp.Children.Add(fn_r)
        self._pdf_fn_grp = make_grp("Custom Filename Builder", fn_sp)
        outer.Children.Add(self._pdf_fn_grp)

        # Report
        rep = vsp()
        self.chk_report = make_chk("Generate CSV report (saved to output folder)")
        rep.Children.Add(self.chk_report)
        self._pdf_rep_grp = make_grp("Export Report", rep)
        outer.Children.Add(self._pdf_rep_grp)

        sv.Content = outer
        return sv

    def _on_nwc_toggled(self, sender, e):
        import System.Windows as _SW
        try:
            nwc_on   = bool(self.chk_nwc.IsChecked)
            pdf_dwg_on = bool(self.chk_pdf.IsChecked) or bool(self.chk_dwg.IsChecked)
            # When NWC checked: uncheck PDF and DWG
            if nwc_on:
                self.chk_pdf.IsChecked = False
                self.chk_dwg.IsChecked = False
            show = _SW.Visibility.Visible
            hide = _SW.Visibility.Collapsed
            # NWC panel: show when NWC on
            self.nwc_panel.Visibility   = show if nwc_on else hide
            # Filename builder: always visible
            self._pdf_fn_grp.Visibility = show
            # PDF-only panels: hide when NWC on
            pdf_vis = hide if nwc_on else show
            self._pdf_row1.Visibility    = pdf_vis
            self._pdf_row2.Visibility    = pdf_vis
            self._pdf_rep_grp.Visibility = pdf_vis
            # Swap template only when user clicks, not when restoring profile
            if not getattr(self, "_applying_profile", False):
                if nwc_on:
                    self.txt_template.Text = "{Sheet Name}"
                else:
                    self.txt_template.Text = "{Sheet Number}"
        except Exception as ex:
            import System.Windows as _SW2
            _SW2.MessageBox.Show("Toggle error: {}".format(str(ex)), "Error")

    def _build_nwc_panel(self):
        outer = vsp(margin=Thickness(0, 8, 0, 0))

        row = StackPanel(); row.Orientation = Orientation.Horizontal
        row.Margin = Thickness(0, 0, 0, 8)

        # Standard options
        std = vsp()
        std.Children.Add(make_lbl("Standard options", bold=True, size=11))
        self.nwc_convert_parts  = make_chk("Convert construction parts",    False)
        self.nwc_convert_ids    = make_chk("Convert element Ids",           True)
        self.nwc_convert_props  = make_chk("Convert element properties",    False)
        self.nwc_convert_links  = make_chk("Export linked Revit models",    False)
        self.nwc_convert_room   = make_chk("Convert room as attribute",     True)
        self.nwc_convert_urls   = make_chk("Convert URLs",                  True)
        self.nwc_divide_levels  = make_chk("Divide file into levels",       False)
        self.nwc_export_rooms   = make_chk("Export room geometry",          True)
        self.nwc_find_materials = make_chk("Try and find missing materials", True)

        for chk in [self.nwc_convert_parts, self.nwc_convert_ids,
                    self.nwc_convert_props, self.nwc_convert_links,
                    self.nwc_convert_room,  self.nwc_convert_urls,
                    self.nwc_divide_levels, self.nwc_export_rooms,
                    self.nwc_find_materials]:
            std.Children.Add(chk)

        # Convert element parameters dropdown
        param_row = hsp()
        param_row.Children.Add(make_lbl("Convert element parameters", size=11))
        self.nwc_params_cmb = make_cmb(["All", "Elements", "None"], width=80)
        self.nwc_params_cmb.Margin = Thickness(8, 0, 0, 0)
        param_row.Children.Add(self.nwc_params_cmb)
        std.Children.Add(param_row)

        # Coordinates dropdown
        coord_row = hsp()
        coord_row.Children.Add(make_lbl("Coordinates", size=11))
        self.nwc_coords_cmb = make_cmb(["Shared", "Internal", "Project"], selected=0, width=100)
        self.nwc_coords_cmb.Margin = Thickness(8, 0, 0, 0)
        coord_row.Children.Add(self.nwc_coords_cmb)
        std.Children.Add(coord_row)

        row.Children.Add(make_grp("Standard options", std))

        # Revit 2020+ options
        rev = vsp()
        self.nwc_convert_cad    = make_chk("Convert linked CAD format", True)
        self.nwc_convert_lights = make_chk("Convert lights",            False)
        rev.Children.Add(self.nwc_convert_cad)
        rev.Children.Add(self.nwc_convert_lights)

        facet_row = hsp()
        facet_row.Children.Add(make_lbl("Faceting factor", size=11))
        self.nwc_faceting = make_txt("1", width=60)
        self.nwc_faceting.Margin = Thickness(8, 0, 0, 0)
        facet_row.Children.Add(self.nwc_faceting)
        rev.Children.Add(facet_row)

        row.Children.Add(make_grp("Revit 2020 & above", rev))
        outer.Children.Add(row)
        return outer

    # ─────────────────────────────────────────────────────────────────────────
    #  PAGE: CREATE
    # ─────────────────────────────────────────────────────────────────────────
    def _build_page_cre(self):
        import System.Windows as _SW5
        import System.Windows.Controls as _SWC5
        import System.Windows.Media as _SWM5

        root = Grid()
        root.RowDefinitions.Add(RowDefinition(Height=GridLength(1, GridUnitType.Auto)))
        root.RowDefinitions.Add(RowDefinition(Height=GridLength(1, GridUnitType.Star)))
        root.Margin = Thickness(16, 12, 16, 12)

        # Top section in ScrollViewer
        top_sv = ScrollViewer()
        top_sv.SetValue(Grid.RowProperty, 0)
        top_sv.VerticalScrollBarVisibility = _SWC5.ScrollBarVisibility.Disabled
        top_sv.MaxHeight = 500
        top = vsp()

        top.Children.Add(make_lbl("Output Folder", bold=True))
        folder_dp = DockPanel(); folder_dp.LastChildFill = True
        folder_dp.Margin = Thickness(0,6,0,8)
        btn_br = make_btn("Browse...", bg=BR_PANEL, fg=BR_TEXT, width=110, height=28)
        btn_br.Click += self._browse
        DockPanel.SetDock(btn_br, Dock.Right)
        folder_dp.Children.Add(btn_br)
        self.txt_folder = make_txt("", readonly=True)
        folder_dp.Children.Add(self.txt_folder)
        top.Children.Add(folder_dp)

        self.chk_date_folder = make_chk("Create dated subfolder  (e.g. 2024-01-15\\)")
        self.chk_fmt_folder  = make_chk("Create format subfolder (e.g. PDF\\ or DWG\\)")
        self.chk_date_folder.Margin = Thickness(0,2,0,2)
        self.chk_fmt_folder.Margin  = Thickness(0,2,0,4)
        top.Children.Add(self.chk_date_folder)
        top.Children.Add(self.chk_fmt_folder)



        # Scheduler — inside top, inside ScrollViewer
        sched_sp = vsp()

        days_row = hsp(margin=Thickness(0,0,0,6))
        days_row.Children.Add(make_lbl("Days:", size=11, bold=True))
        days_row.Margin = Thickness(0,0,0,4)
        self.sched_days = {}
        for day in ["Mon","Tue","Wed","Thu","Fri","Sat","Sun"]:
            chk = CheckBox()
            chk.Content = day; chk.FontSize = 11
            chk.Margin = Thickness(8,0,0,0)
            chk.IsChecked = day in ["Mon","Tue","Wed","Thu","Fri"]
            days_row.Children.Add(chk)
            self.sched_days[day] = chk
        sched_sp.Children.Add(days_row)

        time_row = hsp(margin=Thickness(0,0,0,6))
        time_row.Children.Add(make_lbl("Time (HH:MM):", size=11, bold=True))
        self.txt_sched_time = make_txt("08:00", width=70)
        self.txt_sched_time.Margin = Thickness(8,0,16,0)
        time_row.Children.Add(self.txt_sched_time)

        self.btn_schedule = make_btn("Start Schedule", bg=BR_ACCENT, fg=BR_WHITE, width=110, height=28)
        self.btn_cancel_sched = make_btn("Stop", bg=BR_RED, fg=BR_WHITE, width=60, height=28)
        self.btn_cancel_sched.Visibility = Visibility.Collapsed
        self.btn_cancel_sched.Margin = Thickness(6,0,0,0)
        self.btn_schedule.Click += self._on_schedule
        self.btn_cancel_sched.Click += self._on_cancel_schedule
        time_row.Children.Add(self.btn_schedule)
        time_row.Children.Add(self.btn_cancel_sched)
        sched_sp.Children.Add(time_row)

        sel_row = hsp(margin=Thickness(0,0,0,4))
        self.btn_sched_configure = make_btn("Configure Scheduled Export...", bg=BR_ACCENT, fg=BR_WHITE, width=200, height=28)
        self.btn_sched_configure.Click += self._on_sched_configure
        self.lbl_sched_sel_count = make_lbl("(not configured)", size=11, fg=BR_SUBTEXT)
        self.lbl_sched_sel_count.Margin = Thickness(10,0,0,0)
        sel_row.Children.Add(self.btn_sched_configure)
        sel_row.Children.Add(self.lbl_sched_sel_count)
        sched_sp.Children.Add(sel_row)

        self.lbl_sched_status = make_lbl("", size=11, fg=BR_ACCENT)
        sched_sp.Children.Add(self.lbl_sched_status)

        top.Children.Add(make_grp("Scheduling Assistant", sched_sp))
        top_sv.Content = top
        root.Children.Add(top_sv)

        # Selected Sheets grid fills remaining space
        sel_root = Grid()
        sel_root.SetValue(Grid.RowProperty, 1)
        sel_root.Margin = Thickness(0, 8, 0, 0)
        sel_root.RowDefinitions.Add(RowDefinition(Height=GridLength(30)))   # header
        sel_root.RowDefinitions.Add(RowDefinition(Height=GridLength(1, GridUnitType.Auto)))  # progress bar
        sel_root.RowDefinitions.Add(RowDefinition(Height=GridLength(1, GridUnitType.Star)))  # list

        hdr_g = Grid()
        hdr_g.Background = _SWM5.SolidColorBrush(_SWM5.Color.FromRgb(238,240,244))
        hdr_g.SetValue(Grid.RowProperty, 0)
        for w in [200, 0, 80]:
            cd = ColumnDefinition()
            cd.Width = GridLength(w) if w else GridLength(1, GridUnitType.Star)
            hdr_g.ColumnDefinitions.Add(cd)
        for i, txt in enumerate(["Sheet Number", "Sheet Name", "Revision"]):
            t = TextBlock(); t.Text = txt
            t.FontWeight = FontWeights.Bold; t.FontSize = 11
            t.VerticalAlignment = VerticalAlignment.Center
            t.Margin = Thickness(8,0,0,0)
            Grid.SetColumn(t, i); hdr_g.Children.Add(t)
        sel_root.Children.Add(hdr_g)

        # Progress bar (hidden until export starts)
        prog_border = Border()
        prog_border.SetValue(Grid.RowProperty, 1)
        prog_border.Background = SolidColorBrush(Color.FromRgb(245,246,248))
        prog_border.BorderBrush = BR_BORDER
        prog_border.BorderThickness = Thickness(0,0,0,1)
        prog_border.Padding = Thickness(10,6,10,6)
        import System.Windows.Controls as _SWCpb
        prog_inner = vsp()
        prog_top = hsp()
        self.lbl_export_status = make_lbl("Exporting...", bold=True, size=11)
        self.lbl_export_count  = make_lbl("", size=11, fg=BR_SUBTEXT)
        self.lbl_export_count.Margin = Thickness(8,0,0,0)
        prog_top.Children.Add(self.lbl_export_status)
        prog_top.Children.Add(self.lbl_export_count)
        prog_inner.Children.Add(prog_top)
        self.export_progress = _SWCpb.ProgressBar()
        self.export_progress.Height  = 14
        self.export_progress.Minimum = 0
        self.export_progress.Maximum = 100
        self.export_progress.Value   = 0
        self.export_progress.Margin  = Thickness(0,4,0,0)
        self.export_progress.Foreground = SolidColorBrush(Color.FromRgb(22,160,133))
        prog_inner.Children.Add(self.export_progress)
        prog_border.Child = prog_inner
        import System.Windows as _SWpb
        prog_border.Visibility = _SWpb.Visibility.Collapsed
        self.export_prog_border = prog_border
        sel_root.Children.Add(prog_border)

        sv = ScrollViewer()
        sv.SetValue(Grid.RowProperty, 2)
        sv.VerticalScrollBarVisibility = ScrollBarVisibility.Auto
        self.sel_rows_panel = SWC.StackPanel()
        self.sel_rows_panel.Background = Brushes.White
        sv.Content = self.sel_rows_panel
        sel_root.Children.Add(sv)

        sg = make_grp("Selected Sheets", sel_root)
        sg.SetValue(Grid.RowProperty, 1)
        sg.HorizontalAlignment = HorizontalAlignment.Stretch
        sg.VerticalAlignment   = VerticalAlignment.Stretch
        root.Children.Add(sg)

        return root

    def _go(self, idx):
        self.cur_page = idx
        pages = [self.page_sel, self.page_fmt, self.page_cre]
        self.content_border.Child = pages[idx]

        # Tab button styles
        import System.Windows.Media as _SWM2
        import System.Windows as _SW2
        accent  = _SWM2.SolidColorBrush(_SWM2.Color.FromRgb(22, 160, 133))
        subtext = _SWM2.SolidColorBrush(_SWM2.Color.FromRgb(120, 120, 135))
        for i, b in enumerate(self.tab_btns):
            b.Foreground = accent if i == idx else subtext
            b.FontWeight = _SW2.FontWeights.Bold if i == idx else _SW2.FontWeights.Normal

        import System.Windows as _SW3
        self.btn_back.Visibility   = _SW3.Visibility.Visible  if idx > 0             else _SW3.Visibility.Collapsed
        self.btn_next.Visibility   = _SW3.Visibility.Visible  if idx < self.PAGE_CRE else _SW3.Visibility.Collapsed
        self.btn_export.Visibility = _SW3.Visibility.Visible  if idx == self.PAGE_CRE else _SW3.Visibility.Collapsed
        try:
            self._update_summary()
        except: pass



    def _on_tab(self, sender, e):
        self._go(int(sender.Tag))
    def _on_next(self, s, e):
        if self.cur_page < self.PAGE_CRE:
            self._go(self.cur_page + 1)
    def _on_back(self, s, e):
        if self.cur_page > self.PAGE_SEL: self._go(self.cur_page - 1)

    def _update_summary(self):
        if not hasattr(self, "sel_rows_panel"): return
        import System.Windows.Controls as _SWC4
        import System.Windows.Media as _SWM4
        import System.Windows as _SW4
        checked = self._get_checked()
        self.sel_rows_panel.Children.Clear()
        self._summary_row_panels = {}
        for i, sr in enumerate(checked):
            outer = _SWC4.Border()
            outer.BorderThickness = _SW4.Thickness(0,0,0,1)
            outer.BorderBrush = _SWM4.SolidColorBrush(_SWM4.Color.FromRgb(235,237,240))
            outer.Background = _SWM4.SolidColorBrush(
                _SWM4.Color.FromRgb(248,249,251) if i%2 else _SWM4.Color.FromRgb(255,255,255))
            dp = _SWC4.DockPanel(); dp.LastChildFill = True; dp.MinHeight = 28
            icon = _SWC4.TextBlock(); icon.Text = ""; icon.Width = 22
            icon.FontSize = 13; icon.VerticalAlignment = _SW4.VerticalAlignment.Center
            icon.Margin = _SW4.Thickness(6,0,0,0)
            _SWC4.DockPanel.SetDock(icon, _SWC4.Dock.Left); dp.Children.Add(icon)
            num = _SWC4.TextBlock(); num.Text = sr.SheetNumber or ""
            num.Width = 200; num.FontSize = 12
            num.VerticalAlignment = _SW4.VerticalAlignment.Center
            num.Margin = _SW4.Thickness(4,2,0,2)
            _SWC4.DockPanel.SetDock(num, _SWC4.Dock.Left); dp.Children.Add(num)
            rev = _SWC4.TextBlock(); rev.Text = sr.Revision or ""
            rev.Width = 80; rev.FontSize = 12
            rev.VerticalAlignment = _SW4.VerticalAlignment.Center
            rev.Margin = _SW4.Thickness(8,2,8,2)
            _SWC4.DockPanel.SetDock(rev, _SWC4.Dock.Right); dp.Children.Add(rev)
            name = _SWC4.TextBlock(); name.Text = sr.SheetName or ""
            name.FontSize = 12; name.VerticalAlignment = _SW4.VerticalAlignment.Center
            name.Margin = _SW4.Thickness(4,2,0,2); dp.Children.Add(name)
            outer.Child = dp
            self.sel_rows_panel.Children.Add(outer)
            self._summary_row_panels[sr.SheetNumber] = (outer, icon)

    def _mark_row_done(self, sheet_number, ok=True):
        import System.Windows.Media as _SWM5
        import System.Windows as _SW5
        try:
            pair = getattr(self, "_summary_row_panels", {}).get(sheet_number)
            if pair:
                outer, icon = pair
                if ok:
                    outer.Background = _SWM5.SolidColorBrush(_SWM5.Color.FromRgb(220,245,220))
                    icon.Text = "✓"
                    icon.Foreground = _SWM5.SolidColorBrush(_SWM5.Color.FromRgb(40,140,60))
                else:
                    outer.Background = _SWM5.SolidColorBrush(_SWM5.Color.FromRgb(255,220,220))
                    icon.Text = "✗"
                    icon.Foreground = _SWM5.SolidColorBrush(_SWM5.Color.FromRgb(190,50,50))
        except: pass

    def _mark_row_active(self, sheet_number):
        import System.Windows.Media as _SWM6
        try:
            pair = getattr(self, "_summary_row_panels", {}).get(sheet_number)
            if pair:
                outer, icon = pair
                outer.Background = _SWM6.SolidColorBrush(_SWM6.Color.FromRgb(255,250,220))
                icon.Text = "⟳"
                icon.Foreground = _SWM6.SolidColorBrush(_SWM6.Color.FromRgb(180,130,20))
        except: pass

    def _insert_token(self, sender, e):
        self.txt_template.Text += str(sender.Tag)

    def _refresh_preview(self, s, e):
        checked = self._get_checked()
        sheet = checked[0].sheet if checked else (self.all_sheets[0] if self.all_sheets else None)
        if not sheet: self.lbl_preview.Text = "(no sheets)"; return
        sep  = str(self.cmb_sep.SelectedItem or "-")
        self.lbl_preview.Text = build_filename(sheet, self.txt_template.Text, sep) + ".pdf"

    def _browse(self, s, e):
        import clr as _clr
        _clr.AddReference("System.Windows.Forms")
        from System.Windows.Forms import FolderBrowserDialog, DialogResult
        dlg = FolderBrowserDialog()
        dlg.Description = "Select output folder"
        if dlg.ShowDialog() == DialogResult.OK:
            self.txt_folder.Text = dlg.SelectedPath

    # ─────────────────────────────────────────────────────────────────────────
    #  PROFILES
    # ─────────────────────────────────────────────────────────────────────────
    def _collect(self):
        return {
            "template"        : self.txt_template.Text,
            "separator"       : self.cmb_sep.SelectedIndex,
            "export_pdf"      : bool(self.chk_pdf.IsChecked),
            "export_dwg"      : bool(self.chk_dwg.IsChecked),
            "combine"         : bool(self.rad_combine.IsChecked),
            "combined_name"   : self.txt_combined.Text,
            "quality_idx"     : self.cmb_quality.SelectedIndex,
            "color_idx"       : self.cmb_colors.SelectedIndex,
            "zoom_fit"        : bool(self.rad_fit.IsChecked),
            "zoom_pct"        : int(self.txt_zoom.Text) if self.txt_zoom.Text.isdigit() else 100,
            "raster_proc"     : bool(self.rad_raster.IsChecked),
            "hide_ref"        : bool(self.chk_hide_ref.IsChecked),
            "hide_unref"      : bool(self.chk_hide_unref.IsChecked),
            "hide_scope"      : bool(self.chk_hide_scope.IsChecked),
            "hide_crop"       : bool(self.chk_hide_crop.IsChecked),
            "replace_halftone": bool(self.chk_replace_half.IsChecked),
            "thin_lines"      : bool(self.chk_thin_lines.IsChecked),
            "report"          : bool(self.chk_report.IsChecked),
            "date_folder"     : bool(self.chk_date_folder.IsChecked),
            "fmt_folder"      : bool(self.chk_fmt_folder.IsChecked),
            "output_folder"   : self.txt_folder.Text,
            "export_nwc"      : bool(self.chk_nwc.IsChecked),
            "nwc_convert_parts"  : bool(self.nwc_convert_parts.IsChecked),
            "nwc_convert_ids"    : bool(self.nwc_convert_ids.IsChecked),
            "nwc_convert_props"  : bool(self.nwc_convert_props.IsChecked),
            "nwc_convert_links"  : bool(self.nwc_convert_links.IsChecked),
            "nwc_convert_room"   : bool(self.nwc_convert_room.IsChecked),
            "nwc_convert_urls"   : bool(self.nwc_convert_urls.IsChecked),
            "nwc_divide_levels"  : bool(self.nwc_divide_levels.IsChecked),
            "nwc_export_rooms"   : bool(self.nwc_export_rooms.IsChecked),
            "nwc_find_materials" : bool(self.nwc_find_materials.IsChecked),
            "nwc_params_idx"     : self.nwc_params_cmb.SelectedIndex,
            "nwc_coords_idx"     : self.nwc_coords_cmb.SelectedIndex,
            "nwc_convert_cad"    : bool(self.nwc_convert_cad.IsChecked),
            "nwc_convert_lights" : bool(self.nwc_convert_lights.IsChecked),
            "nwc_faceting"       : self.nwc_faceting.Text,
        }

    def _apply(self, d):
        self._applying_profile = True  # prevent _on_nwc_toggled from overwriting template
        self.txt_template.Text          = d.get("template", "{Sheet Number}")
        self.cmb_sep.SelectedIndex      = d.get("separator", 0)
        self.chk_pdf.IsChecked          = d.get("export_pdf", True)
        self.chk_dwg.IsChecked          = d.get("export_dwg", False)
        self.rad_combine.IsChecked      = d.get("combine", False)
        self.rad_separate.IsChecked     = not d.get("combine", False)
        self.txt_combined.Text          = d.get("combined_name", "Combined_Sheets")
        self.cmb_quality.SelectedIndex  = d.get("quality_idx", 2)
        self.cmb_colors.SelectedIndex   = d.get("color_idx", 0)
        self.rad_fit.IsChecked          = d.get("zoom_fit", True)
        self.rad_zoom.IsChecked         = not d.get("zoom_fit", True)
        self.txt_zoom.Text              = str(d.get("zoom_pct", 100))
        self.rad_raster.IsChecked       = d.get("raster_proc", False)
        self.rad_vector.IsChecked       = not d.get("raster_proc", False)
        self.chk_hide_ref.IsChecked     = d.get("hide_ref", True)
        self.chk_hide_unref.IsChecked   = d.get("hide_unref", True)
        self.chk_hide_scope.IsChecked   = d.get("hide_scope", True)
        self.chk_hide_crop.IsChecked    = d.get("hide_crop", True)
        self.chk_replace_half.IsChecked = d.get("replace_halftone", False)
        self.chk_thin_lines.IsChecked   = d.get("thin_lines", False)
        self.chk_report.IsChecked       = d.get("report", False)
        self.chk_date_folder.IsChecked  = d.get("date_folder", False)
        self.chk_fmt_folder.IsChecked   = d.get("fmt_folder", False)
        if d.get("output_folder"):
            self.txt_folder.Text           = d.get("output_folder", "")
        self.chk_nwc.IsChecked             = d.get("export_nwc", False)
        self.nwc_convert_parts.IsChecked   = d.get("nwc_convert_parts", False)
        self.nwc_convert_ids.IsChecked     = d.get("nwc_convert_ids", True)
        self.nwc_convert_props.IsChecked   = d.get("nwc_convert_props", False)
        self.nwc_convert_links.IsChecked   = d.get("nwc_convert_links", False)
        self.nwc_convert_room.IsChecked    = d.get("nwc_convert_room", True)
        self.nwc_convert_urls.IsChecked    = d.get("nwc_convert_urls", True)
        self.nwc_divide_levels.IsChecked   = d.get("nwc_divide_levels", False)
        self.nwc_export_rooms.IsChecked    = d.get("nwc_export_rooms", True)
        self.nwc_find_materials.IsChecked  = d.get("nwc_find_materials", True)
        self.nwc_params_cmb.SelectedIndex  = d.get("nwc_params_idx", 0)
        self.nwc_coords_cmb.SelectedIndex  = d.get("nwc_coords_idx", 0)
        self.nwc_convert_cad.IsChecked     = d.get("nwc_convert_cad", True)
        self.nwc_convert_lights.IsChecked  = d.get("nwc_convert_lights", False)
        self.nwc_faceting.Text             = d.get("nwc_faceting", "1")
        import System.Windows as _SWA
        self.nwc_panel.Visibility = _SWA.Visibility.Visible if d.get("export_nwc") else _SWA.Visibility.Collapsed
        # Re-apply full panel visibility (hides PDF panels if NWC is on)
        try:
            self._on_nwc_toggled(None, None)
        except: pass
        self._applying_profile = False

    def _prof_dir(self, project=False):
        import os as _o, re as _r
        base = _o.path.join(_o.environ.get("APPDATA", _o.path.expanduser("~")), "pyRevit_SheetsExporter")
        if not _o.path.exists(base):
            _o.makedirs(base)
        if not project:
            return base
        # Project-specific subfolder keyed by document title
        try:
            import Autodesk.Revit.DB as _DBP
            _doc = __revit__.ActiveUIDocument.Document
            title = _doc.Title or "Unknown"
            safe  = _r.sub(r'[\\/*?:"<>|]', "-", title).strip()
            proj_dir = _o.path.join(base, "projects", safe)
            if not _o.path.exists(proj_dir):
                _o.makedirs(proj_dir)
            return proj_dir
        except:
            return base

    def _save_col_widths(self):
        import os as _o, json as _j
        p = _o.path.join(self._prof_dir(), "_col_widths.json")
        with open(p, "w") as f:
            _j.dump({"widths": list(self.col_widths)}, f)

    def _load_col_widths(self):
        import os as _o, json as _j
        p = _o.path.join(self._prof_dir(), "_col_widths.json")
        if _o.path.exists(p):
            try:
                with open(p) as f:
                    return _j.load(f).get("widths", None)
            except: pass
        return None

    def _save_scheduler_state(self):
        import os as _o, json as _j
        p = _o.path.join(self._prof_dir(project=True), "_scheduler.json")
        days = {d: bool(chk.IsChecked) for d, chk in self.sched_days.items()}
        # Save enough info to re-resolve sheets on reopen
        saved = []
        for sr in (self._sched_sheets if hasattr(self, "_sched_sheets") else []):
            if hasattr(sr, "view"):
                try:
                    uid = sr.view.UniqueId
                except:
                    uid = ""
                saved.append({"type": "view", "name": sr.SheetNumber, "uid": uid})
            else:
                saved.append({"type": "sheet", "number": sr.SheetNumber, "name": sr.SheetName})
        state = {
            "time"        : self.txt_sched_time.Text,
            "days"        : days,
            "running"     : self._sched_running,
            "sheet_ids"   : saved,
            "out_folder"  : getattr(self, "_sched_out_folder", ""),
            "fmt_settings": getattr(self, "_sched_fmt_settings", None),
        }
        with open(p, "w") as f: _j.dump(state, f)

    def _load_scheduler_state(self):
        import os as _o, json as _j
        p = _o.path.join(self._prof_dir(project=True), "_scheduler.json")
        if _o.path.exists(p):
            try:
                with open(p) as f: return _j.load(f)
            except: pass
        return None

    def _save_last_profile(self, name):
        import os as _o, json as _j
        p = _o.path.join(self._prof_dir(), "_last_profile.json")
        with open(p, "w") as f: _j.dump({"last": name}, f)

    def _load_last_profile(self):
        import os as _o, json as _j
        p = _o.path.join(self._prof_dir(), "_last_profile.json")
        if _o.path.exists(p):
            try:
                with open(p) as f: return _j.load(f).get("last", "")
            except: pass
        return ""

    def _refresh_prof_cmb(self, reselect=True):
        import os as _o
        d = self._prof_dir()
        profiles = [f[:-5] for f in _o.listdir(d) if f.endswith(".json") and not f.startswith("_")]
        sel = self.cmb_profile.SelectedItem
        # Block _on_profile_changed while repopulating the dropdown
        self._applying_profile = True
        self.cmb_profile.Items.Clear()
        self.cmb_profile.Items.Add("(none)")
        for p in profiles: self.cmb_profile.Items.Add(p)
        if reselect:
            try: self.cmb_profile.SelectedItem = sel
            except: self.cmb_profile.SelectedIndex = 0
        self._applying_profile = False

    def _on_profile_changed(self, sender, e):
        if not hasattr(self, "chk_pdf"): return
        if getattr(self, "_applying_profile", False): return
        import os as _o, json as _j
        name = str(self.cmb_profile.SelectedItem or "")
        if not name or name == "(none)": return
        try:
            import re as _r
            fname = _r.sub(r'[\/*?:"<>|]', "-", name).strip()
            p = _o.path.join(self._prof_dir(), fname + ".json")
            if _o.path.exists(p):
                with open(p) as f: data = _j.load(f)
                self._apply(data)
                self._save_last_profile(name)
        except: pass

    def _prof_load_or_new(self, s, e):
        import System.Windows as _SW
        import System.Windows.Controls as _SWC
        import System.Windows.Controls.Primitives as _SWCP

        # Build a small popup menu: Browse / New Profile
        menu = _SWC.ContextMenu()

        item_browse = _SWC.MenuItem()
        item_browse.Header = "Browse (load existing profile)"
        item_browse.Click += self._prof_load

        item_new = _SWC.MenuItem()
        item_new.Header = "New Profile (save current settings)"
        item_new.Click += self._prof_new

        menu.Items.Add(item_browse)
        menu.Items.Add(item_new)
        menu.IsOpen = True

        # Position under the Load button
        from System.Windows.Controls import Button as _Btn
        menu.PlacementTarget = s
        menu.Placement = _SWCP.PlacementMode.Bottom
        menu.IsOpen = True

    def _prof_new(self, s, e):
        import System.Windows as _SW
        import System.Windows.Controls as _SWC

        # Simple input dialog
        dlg = _SW.Window()
        dlg.Title = "New Profile"
        dlg.Width = 420; dlg.Height = 180
        dlg.WindowStartupLocation = _SW.WindowStartupLocation.CenterScreen
        dlg.ResizeMode = _SW.ResizeMode.NoResize

        try:
            import System.Windows.Interop as _Interop
            import System.Diagnostics as _Diag
            helper = _Interop.WindowInteropHelper(dlg)
            helper.Owner = _Diag.Process.GetCurrentProcess().MainWindowHandle
        except: pass

        import System.Windows.Controls as _SWC2
        import System.Windows.Media as _SWM2
        import System
        sp = _SWC2.StackPanel(); sp.Margin = _SW.Thickness(12)
        lbl = _SWC2.TextBlock(); lbl.Text = "Enter a name for this profile:"
        lbl.FontSize = 12; lbl.Margin = _SW.Thickness(0,0,0,8)
        sp.Children.Add(lbl)
        txt = _SWC2.TextBox(); txt.Height = 28; txt.FontSize = 12
        txt.Padding = _SW.Thickness(6,2,6,2); txt.Margin = _SW.Thickness(0,0,0,10)
        sp.Children.Add(txt)
        btn_row = _SWC2.StackPanel(); btn_row.Orientation = _SWC2.Orientation.Horizontal
        btn_row.HorizontalAlignment = _SW.HorizontalAlignment.Right
        btn_ok = _SWC2.Button(); btn_ok.Content = "Save"
        btn_ok.Width = 70; btn_ok.Height = 28
        btn_ok.Margin = _SW.Thickness(0,0,6,0)
        btn_cancel = _SWC2.Button(); btn_cancel.Content = "Cancel"
        btn_cancel.Width = 70; btn_cancel.Height = 28
        btn_row.Children.Add(btn_ok); btn_row.Children.Add(btn_cancel)
        sp.Children.Add(btn_row)
        dlg.Content = sp

        result = [False]
        def on_ok(s2, e2):
            result[0] = True
            dlg.Close()
        def on_cancel(s2, e2):
            dlg.Close()
        btn_ok.Click += on_ok
        btn_cancel.Click += on_cancel
        txt.KeyDown += lambda s2,e2: on_ok(s2,e2) if str(e2.Key) == "Return" else None

        dlg.ShowDialog()

        if result[0]:
            name = txt.Text.strip()
            if not name:
                _SW.MessageBox.Show("Please enter a profile name.", "No Name")
                return
            try:
                import os as _o, json as _j, re as _r
                d = _o.path.join(_o.environ.get("APPDATA", _o.path.expanduser("~")), "pyRevit_SheetsExporter")
                if not _o.path.exists(d): _o.makedirs(d)
                fname = _r.sub(r'[\/*?:"<>|]', "-", name).strip()
                self._applying_profile = True
                with open(_o.path.join(d, fname + ".json"), "w") as f:
                    _j.dump(self._collect(), f, indent=2)
                self._refresh_prof_cmb()
                self.cmb_profile.SelectedItem = name
                self._save_last_profile(name)
                self._applying_profile = False
                _SW.MessageBox.Show("Profile '{}' saved.".format(name), "Saved")
            except Exception as ex:
                _SW.MessageBox.Show("Save error: {}".format(str(ex)), "Error")

    def _prof_load(self, s, e):
        import os as _o, json as _j
        name = str(self.cmb_profile.SelectedItem or "")
        if not name or name == "(none)": return
        try:
            d_path = _o.path.join(_o.environ.get("APPDATA", _o.path.expanduser("~")), "pyRevit_SheetsExporter")
            import re as _r
            fname = _r.sub(r'[\/*?:"<>|]', "-", name).strip()
            p = _o.path.join(d_path, fname + ".json")
            if _o.path.exists(p):
                with open(p) as f:
                    data = _j.load(f)
                self._apply(data)
        except Exception as ex:
            import System.Windows as _SW
            _SW.MessageBox.Show("Load error: {}".format(str(ex)))

    def _prof_save(self, s, e):
        import System.Windows as _SW9
        import os as _o, json as _j, re as _r
        name = str(self.cmb_profile.SelectedItem or "").strip()
        if not name or name == "(none)":
            _SW9.MessageBox.Show("Select a profile from the dropdown first, or use Load > New Profile to create one.", "No Profile Selected")
            return
        try:
            d = _o.path.join(_o.environ.get("APPDATA", _o.path.expanduser("~")), "pyRevit_SheetsExporter")
            if not _o.path.exists(d): _o.makedirs(d)
            fname = _r.sub(r'[\/*?:"<>|]', "-", name).strip()
            self._applying_profile = True
            data = self._collect()
            with open(_o.path.join(d, fname + ".json"), "w") as f:
                _j.dump(data, f, indent=2)
            self._save_last_profile(name)
            self._applying_profile = False
            _SW9.MessageBox.Show("Profile '{}' saved.".format(name), "Saved")
        except Exception as ex:
            _SW9.MessageBox.Show("Save error: {}".format(str(ex)), "Error")

    def _prof_del(self, s, e):
        import os as _o, re as _r
        name = str(self.cmb_profile.SelectedItem or "")
        if not name or name == "(none)": return
        try:
            d_path = _o.path.join(_o.environ.get("APPDATA", _o.path.expanduser("~")), "pyRevit_SheetsExporter")
            fname = _r.sub(r'[\/*?:"<>|]', "-", name).strip()
            p = _o.path.join(d_path, fname + ".json")
            if _o.path.exists(p): _o.remove(p)
            self._refresh_prof_cmb()
        except: pass

    def _prof_save_inline(self, s, e):
        import System.Windows as _SW8
        import json as _json, os as _os8
        def _save_profile(name, data):
            import os as _o, json as _j, re as _r
            d = _o.path.join(_o.environ.get("APPDATA", _o.path.expanduser("~")), "pyRevit_SheetsExporter")
            if not _o.path.exists(d): _o.makedirs(d)
            fname = _r.sub(r'[\/*?:"<>|]', "-", name or "").strip()
            with open(_o.path.join(d, fname + ".json"), "w") as f:
                _j.dump(data, f, indent=2)
        name = self.txt_prof_name.Text.strip()
        if not name:
            _SW8.MessageBox.Show("Enter a profile name."); return
        try:
            _save_profile(name, self._collect())
            self._refresh_prof_cmb()
            self._save_last_profile(name)
            _SW8.MessageBox.Show("Profile '{}' saved.".format(name))
        except Exception as ex:
            _SW8.MessageBox.Show("Save error: {}".format(str(ex)))

    # ─────────────────────────────────────────────────────────────────────────
    #  SCHEDULER
    # ─────────────────────────────────────────────────────────────────────────
    def _on_sched_configure(self, sender, e):
        import System.Windows as _SWC2
        try:
            cfg_win = self._SchedPickerWindow(self)
            # Restore previous format settings into picker
            if self._sched_fmt_settings:
                try:
                    prof_name = self._sched_fmt_settings.get("_profile_name", "")
                    if prof_name:
                        try: cfg_win.sched_cmb_profile.SelectedItem = prof_name
                        except: pass
                except: pass
            try:
                import System.Windows.Interop as _Interop
                import System.Diagnostics as _Diag
                proc   = _Diag.Process.GetCurrentProcess()
                helper = _Interop.WindowInteropHelper(cfg_win)
                helper.Owner = proc.MainWindowHandle
            except: pass
            result = cfg_win.ShowDialog()
            if result:
                self._sched_sheets        = cfg_win.selected_rows
                self._sched_out_folder    = cfg_win.output_folder
                self._sched_fmt_settings  = cfg_win.format_settings
                count = len(self._sched_sheets)
                fmt_label = "(main format)" if not self._sched_fmt_settings else "(custom format)"
                self.lbl_sched_sel_count.Text = "{} item(s) configured  |  {}  {}".format(
                    count, self._sched_out_folder or "(no folder)", fmt_label)
                self._save_scheduler_state()
        except Exception as ex:
            import traceback
            _SWC2.MessageBox.Show("Configure error: {}\n{}".format(str(ex), traceback.format_exc()), "Error")

    def _restore_scheduler_state(self):
        import System.Windows as _SW
        import traceback as _trb
        # Show what path we're loading from
        try:
            state = self._load_scheduler_state()
        except:
            return
        if not state:
            return
        try:
            import Autodesk.Revit.DB as _DB3
            self.txt_sched_time.Text = state.get("time", "08:00")
            days = state.get("days", {})
            for d, chk in self.sched_days.items():
                chk.IsChecked = days.get(d, d in ["Mon","Tue","Wed","Thu","Fri"])
            self._sched_fmt_settings = state.get("fmt_settings", None)
            self._sched_out_folder   = state.get("out_folder", "")
            saved_ids = state.get("sheet_ids", [])
            if saved_ids:
                sheet_map = {s.SheetNumber: s for s in self.all_sheets}
                view_map  = {v.Name: v for v in self.all_views}
                restored  = []
                for item in saved_ids:
                    if item.get("type") == "view":
                        uid   = item.get("uid", "")
                        vname = item.get("name", "")
                        v = None
                        if uid:
                            try:
                                el = __revit__.ActiveUIDocument.Document.GetElement(uid)
                                if el and hasattr(el, "ViewType"): v = el
                            except: pass
                        if not v: v = view_map.get(vname)
                        if not v:
                            vname_lower = vname.lower()
                            v = next((vv for vv in self.all_views if vv.Name.lower() == vname_lower), None)
                        if v:
                            vr = self._ViewRow.__new__(self._ViewRow)
                            vr.view = v; vr.IsChecked = True
                            vr.SheetNumber = v.Name
                            vr.SheetName   = str(v.ViewType).replace("ViewType.", "")
                            vr.Revision = ""; vr.Size = ""; vr.SpoolTag = ""
                            vr._chk_ref = None
                            restored.append(vr)
                    else:
                        s = sheet_map.get(item.get("number", ""))
                        if s:
                            sr = self._SheetRow.__new__(self._SheetRow)
                            sr.sheet = s; sr.IsChecked = True
                            sr.SheetNumber = s.SheetNumber; sr.SheetName = s.Name
                            p = s.LookupParameter("Current Revision")
                            sr.Revision = p.AsString() if p and p.HasValue else ""
                            sr.Size = "Auto"
                            p2 = s.LookupParameter("SPOOL SHEET HCC IFF TAG")
                            sr.SpoolTag = p2.AsString() if p2 and p2.HasValue else ""
                            sr._chk_ref = None
                            restored.append(sr)
                self._sched_sheets = restored
                count = len(restored)
                fmt_label = "(main format)" if not self._sched_fmt_settings else "(custom format)"
                folder = self._sched_out_folder or "(no folder)"
                if count > 0:
                    self.lbl_sched_sel_count.Text = "{} item(s) configured  |  {}  {}".format(count, folder, fmt_label)
                else:
                    self.lbl_sched_sel_count.Text = "(saved items not found in current project)"
            if state.get("running", False):
                import threading as _th, datetime as _dtr
                time_str = self.txt_sched_time.Text.strip()
                parts = time_str.split(":")
                h, m = int(parts[0]), int(parts[1])
                selected_days = [d for d, chk in self.sched_days.items() if chk.IsChecked]
                if selected_days:
                    self._sched_h = h; self._sched_m = m; self._sched_running = True
                    # Always schedule NEXT occurrence - never fire missed exports on reopen
                    # This prevents the timer firing immediately if we reopen after scheduled time
                    now = _dtr.datetime.now()
                    target = now.replace(hour=h, minute=m, second=0, microsecond=0)
                    if target <= now:
                        # Time already passed today - next occurrence is tomorrow or later
                        # Force schedule_next to find the next valid day
                        pass
                    self._schedule_next(h, m, selected_days)
                    self.lbl_sched_status.Text = "Active - runs {} at {:02d}:{:02d}".format(
                        ", ".join(selected_days), h, m)
                    self.btn_cancel_sched.Visibility = _SW.Visibility.Visible
                    self.btn_schedule.Visibility     = _SW.Visibility.Collapsed
        except: pass

    def _on_schedule(self, sender, e):
        import System.Windows as _SW
        import datetime as _dt, threading as _th, traceback as _tb
        try:
            time_str = self.txt_sched_time.Text.strip()
            try:
                parts = time_str.split(":")
                h, m = int(parts[0]), int(parts[1])
                assert 0 <= h <= 23 and 0 <= m <= 59
            except:
                _SW.MessageBox.Show("Invalid time. Use HH:MM (e.g. 08:00)", "Invalid Time")
                return
            selected_days = [d for d, chk in self.sched_days.items() if chk.IsChecked]
            if not selected_days:
                _SW.MessageBox.Show("Select at least one day.", "No Days")
                return
            self._cancel_existing_timer()
            self._sched_h = h
            self._sched_m = m
            self._sched_running = True
            self._schedule_next(h, m, selected_days)
            self.lbl_sched_status.Text = "Active - runs {} at {:02d}:{:02d}".format(
                ", ".join(selected_days), h, m)
            self.btn_cancel_sched.Visibility = _SW.Visibility.Visible
            self.btn_schedule.Visibility     = _SW.Visibility.Collapsed
            self._save_scheduler_state()
        except Exception as _sex:
            _SW.MessageBox.Show("Schedule error: {}\n{}".format(str(_sex), _tb.format_exc()), "Error")



    def _schedule_next(self, h, m, days):
        import datetime as _dt, threading as _th
        day_map = {"Mon":0,"Tue":1,"Wed":2,"Thu":3,"Fri":4,"Sat":5,"Sun":6}
        now = _dt.datetime.now()
        # Find next future occurrence (must be at least 10 seconds away)
        best = None
        for day in days:
            d = day_map[day]
            days_ahead = (d - now.weekday()) % 7
            candidate = now.replace(hour=h, minute=m, second=0, microsecond=0)
            candidate += _dt.timedelta(days=days_ahead)
            if candidate <= now + _dt.timedelta(seconds=10):
                candidate += _dt.timedelta(days=7)
            if best is None or candidate < best:
                best = candidate

        delay = (best - now).total_seconds()
        # Update status to show exact next fire time
        try:
            import System.Windows.Threading as _SWT_s
            import System as _Sys_s
            _best_str = best.strftime("%a %H:%M")
            def _upd_status():
                self.lbl_sched_status.Text = "Active - next run: {}  (in {:.0f} min)".format(
                    _best_str, delay / 60)
            self.Dispatcher.BeginInvoke(
                _SWT_s.DispatcherPriority.Normal,
                _Sys_s.Action(_upd_status))
        except: pass

        def fire():
            if not getattr(self, "_sched_running", False):
                return
            # Must raise ExternalEvent from UI thread — use Dispatcher
            try:
                import System.Windows.Threading as _SWT2
                import System
                def do_export():
                    try:
                        if not self._sched_sheets:
                            return  # Nothing configured - skip
                        import re as _re3
                        def _safe3(name):
                            return _re3.sub(r'[\/*?:"<>|]', "-", name or "").strip()
                        checked  = self._sched_sheets
                        folder   = getattr(self, "_sched_out_folder", "") or self.txt_folder.Text.strip()
                        # Merge: start with main settings, override with scheduled format settings
                        # Use profile settings if configured, else main settings
                        sched_fmt = getattr(self, "_sched_fmt_settings", None) or {}
                        if sched_fmt:
                            cfg = self._collect()  # get base settings
                            # Override with all profile settings
                            for _k, _v in sched_fmt.items():
                                if not _k.startswith("_"):
                                    cfg[_k] = _v
                        else:
                            cfg = self._collect()
                        # Build separator string
                        sep_chars = ["-","_"," ","."]
                        sep = sep_chars[cfg["separator"]] if cfg["separator"] < len(sep_chars) else "-"
                        is_views = any(hasattr(r, "view") for r in checked)
                        if is_views:
                            items = [r.view for r in checked]
                            sheets = items
                            pairs  = [(v.Id, _safe3(v.Name)) for v in items]
                        else:
                            sheets = [r.sheet for r in checked]
                            pairs  = [(s.sheet.Id, _safe3(s.SheetNumber)) for s in checked]
                        nwc_pairs = pairs
                        def cb(msg, ok):
                            import System.Windows as _SWcb
                            _SWcb.MessageBox.Show(msg, "Scheduled Export Complete")
                        self.export_handler.data = {
                            "sheets": sheets, "pairs": pairs,
                            "nwc_pairs": nwc_pairs, "settings": cfg,
                            "out_folder": folder,
                            "do_pdf": bool(cfg["export_pdf"]) and not is_views,
                            "do_dwg": bool(cfg["export_dwg"]) and not is_views,
                            "do_nwc": bool(cfg["export_nwc"]),
                            "do_report": bool(cfg.get("report", False)),
                            "date_sub":  bool(cfg.get("date_folder", False)),
                            "fmt_sub":   bool(cfg.get("fmt_folder", False)),
                            "callback":  cb,
                        }
                        # Schedule next run inside callback - AFTER export completes
                        def cb(msg, ok):
                            import System.Windows as _SWcb
                            import datetime as _dt2
                            _SWcb.MessageBox.Show(msg, "Scheduled Export Complete")
                            # Only schedule next after this one fully finishes
                            if getattr(self, "_sched_running", False):
                                days2 = [d for d, chk in self.sched_days.items() if chk.IsChecked]
                                self._schedule_next(self._sched_h, self._sched_m, days2)
                                self.lbl_sched_status.Text = "Last ran {}. Next: see status.".format(
                                    _dt2.datetime.now().strftime("%H:%M"))
                        self.export_handler.data["callback"] = cb
                        self.export_event.Raise()
                    except Exception as _ex:
                        import System.Windows as _SWex
                        _SWex.MessageBox.Show("Scheduled export error: {}".format(str(_ex)), "Scheduler Error")
                        # Still schedule next even if export failed
                        if getattr(self, "_sched_running", False):
                            days2 = [d for d, chk in self.sched_days.items() if chk.IsChecked]
                            self._schedule_next(self._sched_h, self._sched_m, days2)
                self.Dispatcher.BeginInvoke(
                    _SWT2.DispatcherPriority.Normal,
                    System.Action(do_export)
                )
            except Exception as _fe:
                import System.Windows as _SWfe
                _SWfe.MessageBox.Show("Fire error: {}".format(str(_fe)), "Scheduler")

        self._sched_timer = _th.Timer(delay, fire)
        self._sched_timer.daemon = True
        self._sched_timer.start()

    def _on_cancel_schedule(self, sender, e):
        import System.Windows as _SW
        import os as _os2
        self._sched_running = False
        self._cancel_existing_timer()
        self.lbl_sched_status.Text = "Schedule stopped."
        self.btn_cancel_sched.Visibility = _SW.Visibility.Collapsed
        self.btn_schedule.Visibility     = _SW.Visibility.Visible
        # Delete the scheduler JSON so it won't restart on reopen
        try:
            _p = _os2.path.join(self._prof_dir(project=True), "_scheduler.json")
            if _os2.path.exists(_p): _os2.remove(_p)
        except: pass

    def _cancel_existing_timer(self):
        if hasattr(self, "_sched_timer") and self._sched_timer:
            try: self._sched_timer.cancel()
            except: pass
            self._sched_timer = None

    # ─────────────────────────────────────────────────────────────────────────
    #  EXPORT
    # ─────────────────────────────────────────────────────────────────────────
    def _on_export(self, s, e):
        try:
            import System.Windows as _SWX
            _os = self._os
            _re = self._re
            _dt = self._dt

            def _safe_fn(name):
                return _re.sub(r'[\\/*?:"<>|]', "-", name or "").strip()

            def _build_filename(sheet, template, sep):
                tokens = {
                    "{Sheet Number}"     : sheet.SheetNumber,
                    "{Sheet Name}"       : sheet.Name,
                    "{Current Revision}" : sheet.LookupParameter("Current Revision").AsString() if sheet.LookupParameter("Current Revision") and sheet.LookupParameter("Current Revision").HasValue else "",
                    "{Date}"             : _dt.datetime.now().strftime("%Y%m%d"),
                }
                r = template
                for t, v in tokens.items():
                    r = r.replace(t, _safe_fn(v or ""))
                while sep * 2 in r:
                    r = r.replace(sep * 2, sep)
                return r.strip(sep) or "Sheet"

            checked = self._get_checked()
            if not checked:
                _SWX.MessageBox.Show("No sheets selected. Go to the Selection tab first.",
                                   "Nothing to export"); return
            folder = self.txt_folder.Text.strip()
            if not folder or not _os.path.isdir(folder):
                _SWX.MessageBox.Show("Please select a valid output folder.", "No Folder"); return
            cfg = self._collect()
            if not cfg["export_pdf"] and not cfg["export_dwg"] and not cfg["export_nwc"]:
                _SWX.MessageBox.Show("Select at least one format on the Format tab.", "No Format"); return
            sep    = str(self.cmb_sep.SelectedItem or "-")
            is_views = self._is_views_mode()
            if is_views:
                # Views mode — use view objects
                view_items = [r.view for r in checked]
                sheets = view_items
                pairs  = [(v.Id, _safe_fn(v.Name)) for v in view_items]
                nwc_pairs = pairs
            else:
                sheets = [r.sheet for r in checked]
                pairs  = [(r.Id, _build_filename(r, cfg["template"], sep)) for r in sheets]
                nwc_pairs = pairs


            def callback(msg, ok):
                import System.Windows.Threading as _SWTP6
                import System as _Sys6
                import System.Windows as _SW6
                def _finish():
                    self.btn_export.IsEnabled = True

                    _SW6.MessageBox.Show(msg, "Export Complete")
                    self.is_updating_selection = True
                    self.list_box.SelectedItems.Clear()
                    self.is_updating_selection = False
                    self.selected_rows = []
                    self._update_summary()
                self.Dispatcher.BeginInvoke(
                    _SWTP6.DispatcherPriority.Normal,
                    _Sys6.Action(_finish))

            total = len(pairs)
            # Show progress bar and reset row statuses immediately (we're on UI thread)
            try:
                import System.Windows as _SWprog
                import System.Windows.Threading as _SWTshow
                import System as _Sysshow
                self._update_summary()
                self.export_progress.Maximum = max(total, 1)
                self.export_progress.Value   = 0
                self.lbl_export_status.Text  = "Exporting..."
                self.lbl_export_count.Text   = "0 / {}".format(total)
                self.export_prog_border.Visibility = _SWprog.Visibility.Visible
                self.btn_export.IsEnabled = False
                # Force immediate repaint before export starts
                _SWTshow.Dispatcher.CurrentDispatcher.Invoke(
                    _Sysshow.Action(lambda: None),
                    _SWTshow.DispatcherPriority.Background)
            except: pass

            def _flush_ui():
                # Force WPF to process all pending render messages
                try:
                    import System.Windows.Threading as _SWTf
                    import System as _Sysf
                    _SWTf.Dispatcher.CurrentDispatcher.Invoke(
                        _Sysf.Action(lambda: None),
                        _SWTf.DispatcherPriority.Background)
                except: pass

            def _row_started(fname):
                try:
                    import System.Windows.Threading as _SWTp
                    import System as _Sysp
                    def _do(): self._mark_row_active(fname)
                    self.Dispatcher.Invoke(_SWTp.DispatcherPriority.Render, _Sysp.Action(_do))
                    _flush_ui()
                except: pass

            def _row_done(fname, ok, done_count):
                try:
                    import System.Windows.Threading as _SWTd
                    import System as _Sysd
                    def _do():
                        self._mark_row_done(fname, ok)
                        self.export_progress.Value = done_count
                        self.lbl_export_count.Text = "{} / {}".format(done_count, total)
                        self.lbl_export_status.Text = fname
                    self.Dispatcher.Invoke(_SWTd.DispatcherPriority.Render, _Sysd.Action(_do))
                    _flush_ui()
                except: pass

            self.export_handler.data = {
                "sheets"    : sheets,
                "pairs"     : pairs,
                "nwc_pairs" : nwc_pairs,
                "settings"  : cfg,
                "out_folder": folder,
                "do_pdf"    : bool(cfg.get("export_pdf", False)) and not is_views,
                "do_dwg"    : bool(cfg.get("export_dwg", False)) and not is_views,
                "do_nwc"    : bool(cfg.get("export_nwc", False)),
                "do_report" : bool(cfg.get("report", False)),
                "date_sub"  : bool(cfg.get("date_folder", False)),
                "fmt_sub"   : bool(cfg.get("fmt_folder", False)),
                "callback"  : callback,
                "is_scheduled": False,
                "row_started": _row_started,
                "row_done"   : _row_done,
            }
            self.export_event.Raise()
        except Exception as ex:
            import traceback as _tb2
            import System.Windows as _SWX2
            _SWX2.MessageBox.Show("Export error: {}\n{}".format(str(ex), _tb2.format_exc()), "Error")


# ─────────────────────────────────────────────────────────────────────────────
#  ENTRY POINT
# ─────────────────────────────────────────────────────────────────────────────
sheets = get_sheets()
if not sheets:
    SW.MessageBox.Show("No sheets found in this project.", "Sheets Exporter")
else:
    preselected, preselected_views = get_preselected_sheet_numbers()
    handler = ExportSheetsHandler()
    event   = ExternalEvent.Create(handler)
    window  = ProSheetsWindow(handler, event, preselected=preselected, preselected_views=preselected_views)
    # Now both classes exist - wire SchedPickerWindow reference

    # Parent to Revit (same as Spool Manager)
    try:
        proc   = Diagnostics.Process.GetCurrentProcess()
        helper = Interop.WindowInteropHelper(window)
        helper.Owner = proc.MainWindowHandle
    except: pass

    # Prevent GC (same as Spool Manager)
    if not hasattr(__main__, "_sheets_exporter_ref"):
        __main__._sheets_exporter_ref = None
    __main__._sheets_exporter_ref = window

    window.Show()