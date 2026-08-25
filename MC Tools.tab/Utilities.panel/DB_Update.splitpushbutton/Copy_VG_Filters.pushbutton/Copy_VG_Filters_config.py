# -*- coding: utf-8 -*-
import clr
import os
import os.path as op
import pickle as pl
import traceback

clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')
clr.AddReference('System.Xaml')

import System
from System.Windows import Window, Thickness, HorizontalAlignment, ResizeMode, WindowStartupLocation, GridLength, GridUnitType
from System.Windows.Controls import Label, TextBox, Button, ScrollViewer, StackPanel, CheckBox, RowDefinition, ColumnDefinition
from System.Windows.Controls import Grid as WpfGrid
from System.Windows.Controls import Orientation, ScrollBarVisibility
from System.Windows.Media import FontFamily

from Autodesk.Revit import DB
from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory, Transaction
from Autodesk.Revit.UI import TaskDialog

from pyrevit import script

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
app = doc.Application
logger = script.get_logger()

USER_TEMP = os.getenv('Temp')
PROJECT_NAME = op.splitext(op.basename(doc.PathName))[0] if doc.PathName else "UnsavedProject"
DATAFILE = op.join(USER_TEMP, PROJECT_NAME + '_pyChecked_Templates.pym')

ALLOWED_TYPES = (
    DB.ViewPlan,
    DB.View3D,
    DB.ViewSection,
    DB.ViewSheet,
    DB.ViewDrafting
)


def get_id_value(obj):
    try:
        eid = obj.Id
    except:
        eid = obj
    try:
        return str(eid.Value)
    except:
        return str(eid.IntegerValue)


def read_checkboxes_state():
    try:
        if not op.exists(DATAFILE):
            return set()
        with open(DATAFILE, 'rb') as f:
            data = pl.load(f)
        return set([str(x) for x in data])
    except Exception as ex:
        logger.warning("Could not read saved state: {}".format(ex))
        return set()


def save_checkboxes_state(elements):
    try:
        data = set([get_id_value(x) for x in elements])
        with open(DATAFILE, 'wb') as f:
            pl.dump(data, f)
    except Exception as ex:
        logger.error("Could not save state: {}".format(ex))


def get_active_view():
    av = doc.ActiveView
    if isinstance(av, ALLOWED_TYPES):
        return av

    sel_ids = list(uidoc.Selection.GetElementIds())
    if sel_ids:
        el = doc.GetElement(sel_ids[0])
        if isinstance(el, ALLOWED_TYPES):
            return el

    return None


def get_view_filters(view):
    items = []
    for fid in view.GetFilters():
        f = doc.GetElement(fid)
        if f:
            items.append(f)
    items.sort(key=lambda x: x.Name.lower())
    return items


def get_view_templates():
    items = []
    collector = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Views).WhereElementIsNotElementType()
    for v in collector:
        try:
            if v.IsTemplate:
                items.append(v)
        except:
            pass
    items.sort(key=lambda x: x.Name.lower())
    return items


class MultiSelectWindow(Window):
    def __init__(self, items, title, button_text, prechecked_ids=None):
        Window.__init__(self)

        self.all_items = items
        self.result_items = []
        self.checkboxes = []
        self.state_map = {}
        self.check_all_state = False
        self.button_text = button_text

        if prechecked_ids is None:
            prechecked_ids = set()

        for item in self.all_items:
            self.state_map[get_id_value(item)] = get_id_value(item) in prechecked_ids

        self.Title = title
        self.Width = 420
        self.Height = 520
        self.MinWidth = 420
        self.MinHeight = 520
        self.ResizeMode = ResizeMode.NoResize
        self.WindowStartupLocation = WindowStartupLocation.CenterScreen

        self.build_ui(title)
        self.update_checkboxes(self.all_items)

    def build_ui(self, title):
        root = WpfGrid()
        root.Margin = Thickness(8)

        root.RowDefinitions.Add(RowDefinition(Height=GridLength.Auto))
        root.RowDefinitions.Add(RowDefinition(Height=GridLength.Auto))
        root.RowDefinitions.Add(RowDefinition(Height=GridLength(1, GridUnitType.Star)))
        root.RowDefinitions.Add(RowDefinition(Height=GridLength.Auto))
        root.ColumnDefinitions.Add(ColumnDefinition())

        lbl = Label()
        lbl.Content = title
        lbl.FontFamily = FontFamily("Arial")
        lbl.FontSize = 16
        lbl.Margin = Thickness(0, 0, 0, 6)
        WpfGrid.SetRow(lbl, 0)
        root.Children.Add(lbl)

        self.search_box = TextBox()
        self.search_box.Height = 24
        self.search_box.FontFamily = FontFamily("Arial")
        self.search_box.FontSize = 12
        self.search_box.Margin = Thickness(0, 0, 0, 6)
        self.search_box.TextChanged += self.on_search_changed
        WpfGrid.SetRow(self.search_box, 1)
        root.Children.Add(self.search_box)

        self.checkbox_panel = StackPanel()
        self.checkbox_panel.Orientation = Orientation.Vertical

        scroll = ScrollViewer()
        scroll.Content = self.checkbox_panel
        scroll.VerticalScrollBarVisibility = ScrollBarVisibility.Auto
        WpfGrid.SetRow(scroll, 2)
        root.Children.Add(scroll)

        btn_panel = StackPanel()
        btn_panel.Orientation = Orientation.Horizontal
        btn_panel.HorizontalAlignment = HorizontalAlignment.Center
        btn_panel.Margin = Thickness(0, 8, 0, 0)

        self.ok_btn = Button()
        self.ok_btn.Content = self.button_text
        self.ok_btn.Width = 110
        self.ok_btn.Height = 28
        self.ok_btn.Margin = Thickness(0, 0, 10, 0)
        self.ok_btn.Click += self.on_ok
        btn_panel.Children.Add(self.ok_btn)

        self.all_btn = Button()
        self.all_btn.Content = "Check All"
        self.all_btn.Width = 90
        self.all_btn.Height = 28
        self.all_btn.Margin = Thickness(0, 0, 10, 0)
        self.all_btn.Click += self.on_check_all
        btn_panel.Children.Add(self.all_btn)

        self.cancel_btn = Button()
        self.cancel_btn.Content = "Cancel"
        self.cancel_btn.Width = 80
        self.cancel_btn.Height = 28
        self.cancel_btn.Click += self.on_cancel
        btn_panel.Children.Add(self.cancel_btn)

        WpfGrid.SetRow(btn_panel, 3)
        root.Children.Add(btn_panel)

        self.Content = root

    def update_checkboxes(self, items):
        self.checkbox_panel.Children.Clear()
        self.checkboxes = []

        for item in items:
            cb = CheckBox()
            cb.Content = item.Name
            cb.Tag = item
            cb.Margin = Thickness(2)
            cb.IsChecked = self.state_map.get(get_id_value(item), False)
            cb.Click += self.on_checkbox_click
            self.checkbox_panel.Children.Add(cb)
            self.checkboxes.append(cb)

    def on_checkbox_click(self, sender, args):
        self.state_map[get_id_value(sender.Tag)] = bool(sender.IsChecked)

    def on_check_all(self, sender, args):
        self.check_all_state = not self.check_all_state
        for cb in self.checkboxes:
            cb.IsChecked = self.check_all_state
            self.state_map[get_id_value(cb.Tag)] = self.check_all_state
        self.all_btn.Content = "Uncheck All" if self.check_all_state else "Check All"

    def on_ok(self, sender, args):
        self.result_items = [x for x in self.all_items if self.state_map.get(get_id_value(x), False)]
        self.DialogResult = True
        self.Close()

    def on_cancel(self, sender, args):
        self.DialogResult = False
        self.Close()

    def on_search_changed(self, sender, args):
        txt = self.search_box.Text.lower().strip()
        if not txt:
            filtered = self.all_items
        else:
            filtered = [x for x in self.all_items if txt in x.Name.lower()]
        self.update_checkboxes(filtered)


def main():
    active_view = get_active_view()
    if not active_view:
        TaskDialog.Show("Error", "Open or select a valid source view.")
        return

    source_filters = get_view_filters(active_view)
    if not source_filters:
        TaskDialog.Show("Warning", "Source view has no filters.")
        return

    templates = get_view_templates()
    if not templates:
        TaskDialog.Show("Warning", "No view templates found.")
        return

    dlg_filters = MultiSelectWindow(source_filters, "Select filters to copy", "Select Filters")
    if not dlg_filters.ShowDialog():
        return
    selected_filters = dlg_filters.result_items
    if not selected_filters:
        return

    saved_ids = read_checkboxes_state()
    dlg_templates = MultiSelectWindow(templates, "Select templates to apply filters", "Apply Filters", saved_ids)
    if not dlg_templates.ShowDialog():
        return
    selected_templates = dlg_templates.result_items
    if not selected_templates:
        return

    save_checkboxes_state(selected_templates)

    t = Transaction(doc, "Copy VG Filters")
    t.Start()
    try:
        for vt in selected_templates:
            for filt in selected_filters:
                try:
                    try:
                        vt.RemoveFilter(filt.Id)
                    except:
                        pass

                    overrides = active_view.GetFilterOverrides(filt.Id)
                    vt.SetFilterOverrides(filt.Id, overrides)

                    try:
                        vis = active_view.GetFilterVisibility(filt.Id)
                        vt.SetFilterVisibility(filt.Id, vis)
                    except:
                        pass

                except Exception as ex:
                    logger.warning("Could not apply filter '{}' to template '{}': {}".format(filt.Name, vt.Name, ex))
        t.Commit()
    except Exception:
        t.RollBack()
        raise

    TaskDialog.Show("Success", "Filters copied.")


if __name__ == "__main__":
    try:
        main()
    except Exception as ex:
        TaskDialog.Show("Script Error", "{}\n\n{}".format(str(ex), traceback.format_exc()))