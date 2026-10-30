# -*- coding: utf-8 -*-
"""Hidden Element Manager.

Lists model elements that exist in the project but are not displayed in the
active view, with the reason they are hidden. Select, zoom, unhide, fix,
reveal, and export to CSV.

Target: pyRevit, IronPython 2.7 engine, Revit 2022 and later.
"""
__title__ = "Hidden\nElements"
__persistentengine__ = True  # keeps the modeless window and event handlers alive

import clr

def _main(uiapp_in):
    """Everything lives inside this function on purpose. pyRevit clears
    module-level names once the script ends, but closures keep working."""
    import io
    import re
    import time

    clr.AddReference("PresentationFramework")

    from System.Collections.Generic import List
    from Autodesk.Revit import DB
    from Autodesk.Revit.UI import (IExternalEventHandler, ExternalEvent,
                                   TaskDialog, TaskDialogCommonButtons,
                                   TaskDialogResult)
    from pyrevit import forms

    ALL_REASONS = "(all reasons)"
    ALL_CATS = "(all categories)"

    KIND_LABELS = {
        "element": "Hidden in view (element)",
        "category": "Category hidden",
        "filter": "Hidden by view filter",
        "workset": "Workset hidden or closed",
        "phase": "Phase filter",
        "section": "Outside section box",
        "temp": "Temporary hide/isolate",
        "unexplained": "Unexplained (heuristic)",
    }
    UNEXPLAINED_TEXT = ("Not displayed by Revit, but no specific cause was "
                        "identified (heuristic)")

    XAML = u"""
    <Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
            xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
            Title="Hidden Element Manager v5" Width="920" Height="660"
            MinWidth="720" MinHeight="440"
            WindowStartupLocation="CenterScreen" ShowInTaskbar="False"
            Closed="window_closed">
      <Grid Margin="10">
        <Grid.RowDefinitions>
          <RowDefinition Height="Auto"/>
          <RowDefinition Height="Auto"/>
          <RowDefinition Height="Auto"/>
          <RowDefinition Height="*"/>
          <RowDefinition Height="Auto"/>
          <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <TextBlock Grid.Row="0" Margin="0,0,0,8">
          <Run FontWeight="Bold" Text="Active view: "/>
          <Run x:Name="lbl_view" Text=""/>
        </TextBlock>

        <WrapPanel Grid.Row="1" Margin="0,0,0,8">
          <Button x:Name="btn_refresh" Content="Refresh" Width="80" Height="28" Margin="0,0,6,0" Click="refresh_click"/>
          <Button x:Name="btn_select" Content="Select" Width="80" Height="28" Margin="0,0,6,0" Click="select_click"/>
          <Button x:Name="btn_zoom" Content="Zoom to" Width="80" Height="28" Margin="0,0,6,0" Click="zoom_click"/>
          <Button x:Name="btn_unhide" Content="Unhide element(s)" Width="120" Height="28" Margin="0,0,6,0" Click="unhide_click"/>
          <Button x:Name="btn_fix" Content="Fix all reasons..." Width="120" Height="28" Margin="0,0,6,0" Click="fix_click"/>
          <Button x:Name="btn_reveal" Content="Reveal hidden" Width="110" Height="28" Margin="0,0,6,0" Click="reveal_click"/>
          <Button x:Name="btn_csv" Content="Export CSV..." Width="100" Height="28" Click="csv_click"/>
        </WrapPanel>

        <Grid Grid.Row="2" Margin="0,0,0,8">
          <Grid.ColumnDefinitions>
            <ColumnDefinition Width="Auto"/>
            <ColumnDefinition Width="*"/>
            <ColumnDefinition Width="Auto"/>
            <ColumnDefinition Width="190"/>
            <ColumnDefinition Width="Auto"/>
            <ColumnDefinition Width="190"/>
          </Grid.ColumnDefinitions>
          <TextBlock Grid.Column="0" Text="Search:" VerticalAlignment="Center" Margin="0,0,6,0"/>
          <TextBox x:Name="txt_search" Grid.Column="1" Height="24" VerticalContentAlignment="Center" TextChanged="filter_changed"/>
          <TextBlock Grid.Column="2" Text="Reason:" VerticalAlignment="Center" Margin="12,0,6,0"/>
          <ComboBox x:Name="cmb_reason" Grid.Column="3" Height="24" SelectionChanged="filter_changed"/>
          <TextBlock Grid.Column="4" Text="Category:" VerticalAlignment="Center" Margin="12,0,6,0"/>
          <ComboBox x:Name="cmb_cat" Grid.Column="5" Height="24" SelectionChanged="filter_changed"/>
        </Grid>

        <ListView x:Name="lv" Grid.Row="3" SelectionMode="Extended"
                  ScrollViewer.HorizontalScrollBarVisibility="Disabled">
          <ListView.View>
            <GridView>
              <GridViewColumn Header="Element" Width="230">
                <GridViewColumn.CellTemplate>
                  <DataTemplate>
                    <StackPanel>
                      <TextBlock Text="{Binding Name}" FontWeight="SemiBold"/>
                      <TextBlock Text="{Binding CategoryName}" Foreground="Gray"/>
                    </StackPanel>
                  </DataTemplate>
                </GridViewColumn.CellTemplate>
              </GridViewColumn>
              <GridViewColumn Header="Reason Hidden" Width="520">
                <GridViewColumn.CellTemplate>
                  <DataTemplate>
                    <TextBlock Text="{Binding ReasonText}" TextWrapping="Wrap" Width="500"/>
                  </DataTemplate>
                </GridViewColumn.CellTemplate>
              </GridViewColumn>
              <GridViewColumn Header="Element Id" Width="110" DisplayMemberBinding="{Binding IdText}"/>
            </GridView>
          </ListView.View>
        </ListView>

        <Expander Grid.Row="4" Header="View details" Margin="0,8,0,0">
          <TextBlock x:Name="txt_details" Margin="18,4,0,0" TextWrapping="Wrap"/>
        </Expander>

        <DockPanel Grid.Row="5" Margin="0,8,0,0" LastChildFill="False">
          <TextBlock x:Name="txt_status" DockPanel.Dock="Left" VerticalAlignment="Center"/>
          <StackPanel DockPanel.Dock="Right" Orientation="Horizontal">
            <CheckBox x:Name="chk_deep" Content="Deep scan" IsChecked="True" Margin="0,0,14,0" Click="deep_click"/>
            <CheckBox x:Name="chk_unexplained" Content="Show unexplained" IsChecked="True" Margin="0,0,14,0" Click="filter_changed"/>
            <CheckBox x:Name="chk_auto" Content="Auto-refresh on view change" IsChecked="True"/>
          </StackPanel>
        </DockPanel>
      </Grid>
    </Window>
    """


    # ---------------------------------------------------------------- helpers
    def iid(eid):
        """ElementId to int (works for 2024 .Value and older .IntegerValue)."""
        try:
            return eid.Value
        except AttributeError:
            return eid.IntegerValue


    def is_graphical(view):
        try:
            if view.IsTemplate:
                return False
            if isinstance(view, DB.ViewSchedule):
                return False
            bad = (DB.ViewType.ProjectBrowser, DB.ViewType.SystemBrowser,
                   DB.ViewType.Schedule, DB.ViewType.Report)
            return view.ViewType not in bad
        except Exception:
            return False


    def family_param(el):
        try:
            p = el.get_Parameter(DB.BuiltInParameter.ELEM_FAMILY_PARAM)
            if p is not None:
                return p.AsValueString() or p.AsString() or u""
        except Exception:
            pass
        try:
            p = el.LookupParameter("Family")
            if p is not None:
                return p.AsValueString() or p.AsString() or u""
        except Exception:
            pass
        return u""


    def display_name(el):
        """Fabrication parts: family name (Name just returns 'Default').
        Everything else: Name, falling back to family name."""
        n = u""
        fab = getattr(DB, "FabricationPart", None)
        if fab is not None and isinstance(el, fab):
            n = family_param(el)
        if not n:
            try:
                n = el.Name
            except Exception:
                n = u""
        if not n or n.lower() == u"default":
            n = family_param(el) or n
        return n or u"(unnamed)"


    class Item(object):
        """One row. Public attributes are bound by the ListView."""

        def __init__(self, el, reasons):
            self.ElementId = el.Id
            self.IdText = str(iid(el.Id))
            self.CategoryName = el.Category.Name
            self.Name = display_name(el)
            self.Reasons = reasons  # list of (kind, payload, text)
            self.ReasonText = u"\n".join(u"\u2022 " + r[2] for r in reasons)
            self.Kinds = set(r[0] for r in reasons)


    class Ctx(object):
        """Per-scan cache of view settings."""

        def __init__(self, doc, view, deep):
            self.doc = doc
            self.view = view
            self.deep = deep
            self.cat_cache = {}
            self.ws_cache = {}
            self.hiding_filters = []
            self.temp = False
            self.box = None
            self.tinv = None
            self.phase_id = None
            self.phase_filter = None
            self.ws_defaults = None
            self.workshared = doc.IsWorkshared

            try:
                self.temp = view.IsTemporaryHideIsolateActive()
            except Exception:
                pass

            if not deep:
                return

            try:
                for fid in view.GetFilters():
                    if view.GetFilterVisibility(fid):
                        continue
                    f = doc.GetElement(fid)
                    cats, selids = None, None
                    if isinstance(f, DB.ParameterFilterElement):
                        cats = set(iid(c) for c in f.GetCategories())
                    elif isinstance(f, DB.SelectionFilterElement):
                        selids = set(iid(i) for i in f.GetElementIds())
                    self.hiding_filters.append((fid, f, cats, selids))
            except Exception:
                pass

            if self.workshared:
                try:
                    self.ws_defaults = (DB.WorksetDefaultVisibilitySettings
                                        .GetWorksetDefaultVisibilitySettings(doc))
                except Exception:
                    pass

            try:
                p = view.get_Parameter(DB.BuiltInParameter.VIEW_PHASE).AsElementId()
                pf = view.get_Parameter(DB.BuiltInParameter.VIEW_PHASE_FILTER).AsElementId()
                if iid(p) > 0 and iid(pf) > 0:
                    self.phase_id = p
                    self.phase_filter = doc.GetElement(pf)
            except Exception:
                pass

            try:
                if isinstance(view, DB.View3D) and view.IsSectionBoxActive:
                    self.box = view.GetSectionBox()
                    self.tinv = self.box.Transform.Inverse
            except Exception:
                pass

        def category_hidden(self, cat):
            k = iid(cat.Id)
            if k not in self.cat_cache:
                h = False
                try:
                    if self.view.CanCategoryBeHidden(cat.Id):
                        h = self.view.GetCategoryHidden(cat.Id)
                except Exception:
                    pass
                self.cat_cache[k] = h
            return self.cat_cache[k]

        def workset_state(self, wsid):
            """Return None if visible, else a text reason."""
            k = iid(wsid)
            if k in self.ws_cache:
                return self.ws_cache[k]
            res = None
            try:
                ws = self.doc.GetWorksetTable().GetWorkset(wsid)
                vis = self.view.GetWorksetVisibility(wsid)
                if vis == DB.WorksetVisibility.Hidden:
                    res = u"Workset hidden in view: " + ws.Name
                elif vis == DB.WorksetVisibility.UseGlobalSetting:
                    if (self.ws_defaults is not None and
                            not self.ws_defaults.IsWorksetVisible(wsid)):
                        res = u"Workset hidden by global setting: " + ws.Name
                if res is None and not ws.IsOpen:
                    res = u"Workset is closed: " + ws.Name
            except Exception:
                pass
            self.ws_cache[k] = res
            return res

        def outside_section_box(self, el):
            if self.box is None:
                return False
            bb = el.get_BoundingBox(None)
            if bb is None:
                return False
            xs, ys, zs = [], [], []
            for x in (bb.Min.X, bb.Max.X):
                for y in (bb.Min.Y, bb.Max.Y):
                    for z in (bb.Min.Z, bb.Max.Z):
                        p = self.tinv.OfPoint(bb.Transform.OfPoint(DB.XYZ(x, y, z)))
                        xs.append(p.X)
                        ys.append(p.Y)
                        zs.append(p.Z)
            b = self.box
            return (max(xs) < b.Min.X or min(xs) > b.Max.X or
                    max(ys) < b.Min.Y or min(ys) > b.Max.Y or
                    max(zs) < b.Min.Z or min(zs) > b.Max.Z)


    def explain(el, ctx):
        """Return list of (kind, payload, text). Empty list means not explained."""
        reasons = []
        view = ctx.view
        cat = el.Category

        try:
            if el.IsHidden(view):
                reasons.append(("element", el.Id, u"Hidden in view (Hide in View > Elements)"))
        except Exception:
            pass

        if ctx.category_hidden(cat):
            reasons.append(("category", cat.Id,
                            u"Category hidden in Visibility/Graphics: " + cat.Name))

        if ctx.deep:
            for fid, f, cats, selids in ctx.hiding_filters:
                try:
                    hit = False
                    if cats is not None:
                        if iid(cat.Id) in cats:
                            try:
                                hit = f.PassesFilter(ctx.doc, el.Id)
                            except Exception:
                                hit = f.PassesFilter(el)
                    elif selids is not None:
                        hit = iid(el.Id) in selids
                    if hit:
                        reasons.append(("filter", fid,
                                        u"Hidden by view filter: " + f.Name))
                except Exception:
                    pass

            if ctx.workshared:
                try:
                    txt = ctx.workset_state(el.WorksetId)
                    if txt:
                        reasons.append(("workset", el.WorksetId, txt))
                except Exception:
                    pass

            if ctx.phase_filter is not None:
                try:
                    st = el.GetPhaseStatus(ctx.phase_id)
                    pres = ctx.phase_filter.GetPhaseStatusPresentation(st)
                    if pres == DB.PhaseStatusPresentation.DontShow:
                        reasons.append(("phase", None,
                                        u"Not shown by phase filter: " + ctx.phase_filter.Name))
                except Exception:
                    pass

            try:
                if ctx.outside_section_box(el):
                    reasons.append(("section", None, u"Outside the active section box"))
            except Exception:
                pass

        if not reasons and ctx.temp:
            reasons.append(("temp", None, u"Temporary Hide/Isolate is active in this view"))

        if not reasons:
            try:
                has_geom = el.get_BoundingBox(None) is not None
            except Exception:
                has_geom = False
            if has_geom:
                reasons.append(("unexplained", None, UNEXPLAINED_TEXT))

        return reasons


    def scan(doc, view, deep):
        t0 = time.time()
        ctx = Ctx(doc, view, deep)

        visible = set(iid(i) for i in
                      DB.FilteredElementCollector(doc, view.Id)
                      .WhereElementIsNotElementType().ToElementIds())

        all_ids = (DB.FilteredElementCollector(doc)
                   .WhereElementIsNotElementType()
                   .WhereElementIsViewIndependent().ToElementIds())

        items = []
        cand = 0
        for eid in all_ids:
            el = doc.GetElement(eid)
            if el is None:
                continue
            cat = el.Category
            if cat is None or cat.CategoryType != DB.CategoryType.Model:
                continue
            cand += 1
            if iid(eid) in visible:
                continue
            reasons = explain(el, ctx)
            if reasons:
                items.append(Item(el, reasons))

        items.sort(key=lambda i: (i.CategoryName, i.Name, int(i.IdText)))
        stats = {"candidates": cand, "visible": cand - len(items),
                 "seconds": time.time() - t0}
        return items, stats


    def view_details(doc, view):
        def pstr(bip):
            try:
                return view.get_Parameter(bip).AsValueString() or view.get_Parameter(bip).AsString() or ""
            except Exception:
                return ""
        lines = []
        lines.append(u"View: %s  (%s)" % (view.Name, view.ViewType))
        try:
            dl = view.DetailLevel
        except Exception:
            dl = "n/a"
        try:
            sc = "1:%d" % view.Scale
        except Exception:
            sc = "n/a"
        lines.append(u"Discipline: %s   Detail level: %s   Scale: %s" %
                     (pstr(DB.BuiltInParameter.VIEW_DISCIPLINE), dl, sc))
        lines.append(u"Phase: %s   Phase filter: %s" %
                     (pstr(DB.BuiltInParameter.VIEW_PHASE),
                      pstr(DB.BuiltInParameter.VIEW_PHASE_FILTER)))
        try:
            if isinstance(view, DB.View3D):
                lines.append(u"Section box is %s." %
                             ("active" if view.IsSectionBoxActive else "not active"))
        except Exception:
            pass
        try:
            if iid(view.ViewTemplateId) > 0:
                lines.append(u"View template applied: %s" % doc.GetElement(view.ViewTemplateId).Name)
        except Exception:
            pass
        return u"\n".join(lines)


    # ------------------------------------------------------- external event
    def guarded(fn):
        """Catch everything in a WPF handler. An unhandled exception there
        would crash Revit, so report it in the title bar instead."""
        def wrapper(self, *args):
            try:
                return fn(self, *args)
            except Exception as ex:
                try:
                    self.Title = u"Hidden Element Manager v5  [ERROR in %s: %s]" % (fn.__name__, ex)
                except Exception:
                    pass
        wrapper.__name__ = fn.__name__
        return wrapper


    class _Handler(IExternalEventHandler):
        def __init__(self):
            self.win = None

        def Execute(self, uiapp):
            if self.win is not None:
                self.win.pump(uiapp, "ExternalEvent")

        def GetName(self):
            return "Hidden Element Manager"


    # ----------------------------------------------------------------- window
    class HiddenWindow(forms.WPFWindow):
        def __init__(self, uiapp, handler, ext):
            forms.WPFWindow.__init__(self, XAML, literal_string=True)
            self.uiapp = uiapp
            self.handler = handler
            self.ext = ext
            self.items = []
            self.stats = {"candidates": 0, "visible": 0, "seconds": 0.0}
            self.view_name = ""
            self.scan_view_id = None
            self._loading = False
            self.pending = None
            self.handler.win = self
            self.uiapp.ViewActivated += self.on_view_activated
            self.uiapp.Idling += self.on_idling
            self.Title = u"Hidden Element Manager v5"

        # ---- plumbing
        def run(self, fn, label=None):
            """Queue fn(uiapp) to run in a valid Revit API context.
            Two triggers: ExternalEvent and the Idling event. Whichever fires
            first runs it; the other finds nothing pending."""
            label = label or getattr(fn, "__name__", "action")
            self.pending = (fn, label)
            self.Title = u"Hidden Element Manager v5  [queued: %s]" % label
            try:
                self.ext.Raise()
            except Exception as ex:
                self.Title = u"Hidden Element Manager v5  [Raise failed: %s]" % ex

        def pump(self, uiapp, source):
            p = self.pending
            if p is None:
                return
            self.pending = None
            fn, label = p
            try:
                fn(uiapp)
                self.Title = u"Hidden Element Manager v5  [done: %s via %s]" % (label, source)
            except Exception as ex:
                self.Title = u"Hidden Element Manager v5  [ERROR in %s: %s]" % (label, ex)
                TaskDialog.Show("Hidden Element Manager", "%s failed:\n%s" % (label, ex))

        def on_idling(self, sender, args):
            if self.pending is not None:
                self.pump(sender, "Idling")

        def window_closed(self, sender, args):
            for ev, h in (("ViewActivated", self.on_view_activated),
                          ("Idling", self.on_idling)):
                try:
                    if ev == "ViewActivated":
                        self.uiapp.ViewActivated -= h
                    else:
                        self.uiapp.Idling -= h
                except Exception:
                    pass

        def on_view_activated(self, sender, args):
            if self.chk_auto.IsChecked:
                try:
                    self.refresh(self.uiapp)
                except Exception:
                    pass

        def selected_items(self):
            return [i for i in self.lv.SelectedItems]

        def selected_ids(self):
            ids = List[DB.ElementId]()
            for it in self.selected_items():
                ids.Add(it.ElementId)
            return ids

        # ---- scan and display
        def refresh(self, uiapp):
            uidoc = uiapp.ActiveUIDocument
            if uidoc is None:
                return
            doc = uidoc.Document
            view = uidoc.ActiveView
            self.lbl_view.Text = u"%s  (%s)" % (view.Name, view.ViewType)
            if not is_graphical(view):
                self.items = []
                self.scan_view_id = None
                self.txt_details.Text = u"The active view is not a graphical view."
                self.txt_status.Text = u"Open a plan, section, elevation, or 3D view."
                self.lv.ItemsSource = List[object]()
                return
            self.items, self.stats = scan(doc, view, bool(self.chk_deep.IsChecked))
            self.scan_view_id = view.Id
            self.view_name = view.Name
            self.txt_details.Text = view_details(doc, view)
            self.fill_combos()
            self.apply_filter()

        def fill_combos(self):
            self._loading = True
            try:
                cur_r = self.cmb_reason.SelectedItem
                cur_c = self.cmb_cat.SelectedItem
                reasons = sorted(set(KIND_LABELS[k] for i in self.items for k in i.Kinds))
                cats = sorted(set(i.CategoryName for i in self.items))
                self.cmb_reason.Items.Clear()
                self.cmb_reason.Items.Add(ALL_REASONS)
                for r in reasons:
                    self.cmb_reason.Items.Add(r)
                self.cmb_cat.Items.Clear()
                self.cmb_cat.Items.Add(ALL_CATS)
                for c in cats:
                    self.cmb_cat.Items.Add(c)
                self.cmb_reason.SelectedItem = cur_r if cur_r in reasons else ALL_REASONS
                self.cmb_cat.SelectedItem = cur_c if cur_c in cats else ALL_CATS
            finally:
                self._loading = False

        def apply_filter(self):
            if self._loading:
                return
            txt = (self.txt_search.Text or u"").lower()
            reason = self.cmb_reason.SelectedItem
            cat = self.cmb_cat.SelectedItem
            show_unex = bool(self.chk_unexplained.IsChecked)
            shown = []
            for it in self.items:
                if not show_unex and it.Kinds == set(["unexplained"]):
                    continue
                if reason and reason != ALL_REASONS:
                    if not any(KIND_LABELS[k] == reason for k in it.Kinds):
                        continue
                if cat and cat != ALL_CATS and it.CategoryName != cat:
                    continue
                if txt:
                    hay = u" ".join([it.Name, it.CategoryName, it.IdText, it.ReasonText]).lower()
                    if txt not in hay:
                        continue
                shown.append(it)
            self.lv.ItemsSource = List[object](shown)
            self.txt_status.Text = (
                u"Showing %d of %d hidden element(s)  |  %d candidate(s), %d visible, %.1f s" %
                (len(shown), len(self.items), self.stats["candidates"],
                 self.stats["visible"], self.stats["seconds"]))

        # ---- UI events
        def filter_changed(self, sender, args):
            try:
                self.apply_filter()
            except Exception:
                pass

        def deep_click(self, sender, args):
            self.run(self.refresh)

        def refresh_click(self, sender, args):
            self.run(self.refresh)

        def select_click(self, sender, args):
            if not self.lv.SelectedItems.Count:
                return
            self.run(self.do_select)

        def do_select(self, uiapp):
            ids = self.selected_ids()
            uiapp.ActiveUIDocument.Selection.SetElementIds(ids)

        def zoom_click(self, sender, args):
            if not self.lv.SelectedItems.Count:
                return
            self.run(self.do_zoom)

        def do_zoom(self, uiapp):
            """Zoom the ACTIVE view. ShowElements is not used because Revit
            may switch to a different view to display the element."""
            uidoc = uiapp.ActiveUIDocument
            doc = uidoc.Document
            bbs = []
            for eid in self.selected_ids():
                el = doc.GetElement(eid)
                bb = el.get_BoundingBox(None) if el is not None else None
                if bb is not None:
                    bbs.append(bb)
            if not bbs:
                self.Title = u"Hidden Element Manager v5  [zoom: no geometry]"
                return
            mn = [min(b.Min.X for b in bbs), min(b.Min.Y for b in bbs), min(b.Min.Z for b in bbs)]
            mx = [max(b.Max.X for b in bbs), max(b.Max.Y for b in bbs), max(b.Max.Z for b in bbs)]
            # pad 25 percent (minimum 1 ft) so the element is not edge to edge
            for i in range(3):
                pad = max((mx[i] - mn[i]) * 0.25, 1.0)
                mn[i] -= pad
                mx[i] += pad
            for uiv in uidoc.GetOpenUIViews():
                if uiv.ViewId == uidoc.ActiveView.Id:
                    uiv.ZoomAndCenterRectangle(DB.XYZ(mn[0], mn[1], mn[2]),
                                               DB.XYZ(mx[0], mx[1], mx[2]))
                    break

        def unhide_click(self, sender, args):
            if not self.lv.SelectedItems.Count:
                return
            self.run(lambda uiapp: self.do_fix(uiapp, self.selected_items(), set(["element"]), False), "unhide")

        def fix_click(self, sender, args):
            items = self.selected_items() or list(self.lv.ItemsSource)
            if not items:
                return
            self.run(lambda uiapp: self.do_fix(
                uiapp, items,
                set(["element", "category", "filter", "workset", "section", "temp"]), True), "fix all")

        def do_fix(self, uiapp, items, allowed, confirm):
            uidoc = uiapp.ActiveUIDocument
            doc = uidoc.Document
            view = uidoc.ActiveView
            if self.scan_view_id is None or view.Id != self.scan_view_id:
                TaskDialog.Show("Hidden Element Manager",
                                "The active view changed. Refresh and try again.")
                return

            elem_ids = List[DB.ElementId]()
            seen = set()
            cats, filters, wss = {}, {}, {}
            temp = section = False
            skipped = set()
            for it in items:
                for kind, payload, text in it.Reasons:
                    if kind not in allowed:
                        if kind in ("phase", "unexplained"):
                            skipped.add(kind)
                        continue
                    if kind == "element" and iid(it.ElementId) not in seen:
                        seen.add(iid(it.ElementId))
                        elem_ids.Add(it.ElementId)
                    elif kind == "category":
                        cats[iid(payload)] = (payload, text)
                    elif kind == "filter":
                        filters[iid(payload)] = (payload, text)
                    elif kind == "workset":
                        wss[iid(payload)] = (payload, text)
                    elif kind == "section":
                        section = True
                    elif kind == "temp":
                        temp = True

            todo = []
            if elem_ids.Count:
                todo.append(u"Unhide %d element(s) hidden individually" % elem_ids.Count)
            for p, t in cats.values():
                todo.append(u"Turn on category: " + t.split(u": ", 1)[-1])
            for p, t in filters.values():
                todo.append(u"Turn on filter visibility: " + t.split(u": ", 1)[-1] +
                            u" (affects every element it matches)")
            for p, t in wss.values():
                todo.append(u"Show workset: " + t.split(u": ", 1)[-1])
            if section:
                todo.append(u"Turn off the section box")
            if temp:
                todo.append(u"End temporary hide/isolate")
            if not todo:
                msg = "Nothing in the selection can be fixed automatically."
                if skipped:
                    msg += " Phase filter and unexplained items need manual review."
                TaskDialog.Show("Hidden Element Manager", msg)
                return

            if confirm:
                td = TaskDialog("Fix all reasons")
                td.MainInstruction = "Apply these changes to the active view?"
                td.MainContent = u"\n".join(todo)
                if skipped:
                    td.MainContent += u"\n\nNot changed (manual review): phase filter and unexplained items."
                td.CommonButtons = TaskDialogCommonButtons.Yes | TaskDialogCommonButtons.No
                if td.Show() != TaskDialogResult.Yes:
                    return

            errors = []
            if temp:
                try:
                    view.DisableTemporaryViewMode(DB.TemporaryViewMode.TemporaryHideIsolate)
                except Exception as ex:
                    errors.append("Temporary hide/isolate: %s" % ex)

            t = DB.Transaction(doc, "Hidden Element Manager: unhide")
            t.Start()
            try:
                if elem_ids.Count:
                    try:
                        view.UnhideElements(elem_ids)
                    except Exception as ex:
                        errors.append("Unhide elements: %s" % ex)
                for p, tx in cats.values():
                    try:
                        view.SetCategoryHidden(p, False)
                    except Exception as ex:
                        errors.append("%s: %s" % (tx, ex))
                for p, tx in filters.values():
                    try:
                        view.SetFilterVisibility(p, True)
                    except Exception as ex:
                        errors.append("%s: %s" % (tx, ex))
                for p, tx in wss.values():
                    try:
                        view.SetWorksetVisibility(p, DB.WorksetVisibility.Visible)
                    except Exception as ex:
                        errors.append("%s: %s" % (tx, ex))
                if section:
                    try:
                        view.IsSectionBoxActive = False
                    except Exception as ex:
                        errors.append("Section box: %s" % ex)
                t.Commit()
            except Exception as ex:
                t.RollBack()
                errors.append("Transaction rolled back: %s" % ex)

            if errors:
                TaskDialog.Show("Hidden Element Manager",
                                "Some changes could not be applied (a view template may control them):\n\n" +
                                "\n".join(errors))
            self.refresh(uiapp)

        def reveal_click(self, sender, args):
            self.run(self.do_reveal)

        def do_reveal(self, uiapp):
            doc = uiapp.ActiveUIDocument.Document
            view = uiapp.ActiveUIDocument.ActiveView
            mode = DB.TemporaryViewMode.RevealHiddenElements
            t = DB.Transaction(doc, "Hidden Element Manager: reveal hidden")
            t.Start()
            try:
                if view.IsInTemporaryViewMode(mode):
                    view.DisableTemporaryViewMode(mode)
                    label = "Reveal hidden"
                else:
                    view.EnableRevealHiddenMode()
                    label = "Close reveal"
                t.Commit()
            except Exception:
                t.RollBack()
                raise
            self.btn_reveal.Content = label

        def csv_click(self, sender, args):
            if not self.items:
                return
            safe = re.sub(r'[\\/:*?"<>|]', "_", self.view_name or "view")
            path = forms.save_file(file_ext="csv",
                                   default_name="HiddenElements_%s" % safe,
                                   title="Export hidden elements")
            if not path:
                return

            def q(v):
                return u'"%s"' % (u"%s" % v).replace(u'"', u'""')

            lines = [u",".join(q(h) for h in
                               ("View", "Element Id", "Category", "Element", "Reasons"))]
            for it in self.items:
                lines.append(u",".join([q(self.view_name), q(it.IdText), q(it.CategoryName),
                                        q(it.Name), q(u" | ".join(r[2] for r in it.Reasons))]))
            with io.open(path, "w", encoding="utf-8-sig", newline="") as f:
                f.write(u"\r\n".join(lines) + u"\r\n")
            self.Title = u"Hidden Element Manager v5  [exported %d rows]" % len(self.items)


    # wrap every UI handler so an exception cannot crash Revit
    for _n in ("filter_changed", "deep_click", "refresh_click", "select_click",
               "zoom_click", "unhide_click", "fix_click", "reveal_click", "csv_click"):
        setattr(HiddenWindow, _n, guarded(HiddenWindow.__dict__[_n]))


    # -------------------------------------------------------------------- run
    _uiapp = uiapp_in
    _handler = _Handler()
    _ext = ExternalEvent.Create(_handler)
    _win = HiddenWindow(_uiapp, _handler, _ext)
    _win.refresh(_uiapp)
    _win.show()



_main(__revit__)
