# -*- coding: utf-8 -*-
import clr
clr.AddReference('System')
import System
import System.Diagnostics
clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')
clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')

from System.Windows import Window, Thickness, HorizontalAlignment, WindowStartupLocation, ResizeMode
from System.Windows.Controls import Grid, RowDefinition, Label, TextBox, Button, StackPanel, Orientation
from System.Windows.Interop import WindowInteropHelper

from Autodesk.Revit import DB
from Autodesk.Revit.DB import (
    FilteredElementCollector,
    Transaction,
    Family,
    ViewType
)
from Autodesk.Revit.UI import TaskDialog
from Autodesk.Revit.UI.Events import TaskDialogShowingEventArgs
from Autodesk.Revit.Exceptions import OperationCanceledException

import os
import sys


# --------------------------------------------------
# Basic environment
# --------------------------------------------------
uidoc = __revit__.ActiveUIDocument
doc = uidoc.Document
uiapp = __revit__
view = uidoc.ActiveView
revit_version = int(doc.Application.VersionNumber)

if view.ViewType == ViewType.ThreeD:
    TaskDialog.Show("Error", "Cannot use in 3D view.")
    sys.exit()

script_dir, script_file = os.path.split(__file__)
family_filename = 'FixtureCL.rfa'
family_path = os.path.join(script_dir, family_filename)
family_name = 'FixtureCL'

folder_name = r"c:\temp"
filepath = os.path.join(folder_name, 'Ribbon_FixtureType.txt')

if not os.path.exists(folder_name):
    os.makedirs(folder_name)


def read_last_fixture_type(txt_path, default_value='WC-1'):
    try:
        if not os.path.exists(txt_path):
            with open(txt_path, 'w') as f:
                f.write(default_value)
            return default_value

        with open(txt_path, 'r') as f:
            value = f.read().strip()

        return value if value else default_value
    except:
        return default_value


def save_last_fixture_type(txt_path, value):
    try:
        with open(txt_path, 'w') as f:
            f.write(value or '')
    except:
        pass


prev_input = read_last_fixture_type(filepath, 'WC-1')


# --------------------------------------------------
# Family load options
# --------------------------------------------------
class FamilyLoaderOptionsHandler(DB.IFamilyLoadOptions):
    def OnFamilyFound(self, familyInUse, overwriteParameterValues):
        overwriteParameterValues.Value = False
        return True

    def OnSharedFamilyFound(self, sharedFamily, familyInUse, source, overwriteParameterValues):
        source.Value = DB.FamilySource.Project
        overwriteParameterValues.Value = False
        return True


def shared_family_dialog_fallback(sender, args):
    try:
        if isinstance(args, TaskDialogShowingEventArgs):
            msg = (args.Message or "").lower()
            dialog_id = (args.DialogId or "").lower()

            if ("shared" in msg and "already exists" in msg and "project" in msg) \
               or ("shared" in dialog_id and "family" in dialog_id):
                args.OverrideResult(1003)
    except:
        pass


# --------------------------------------------------
# Family/type cache manager
# --------------------------------------------------
class FixtureFamilyManager(object):
    def __init__(self, document, target_family_name, target_family_path):
        self.doc = document
        self.family_name = target_family_name
        self.family_path = target_family_path
        self.family = None
        self.symbol_cache = {}

    def get_family_by_name(self):
        if self.family and self.family.IsValidObject:
            return self.family

        collector = FilteredElementCollector(self.doc).OfClass(Family)
        for fam in collector:
            if fam.Name == self.family_name:
                self.family = fam
                return fam

        self.family = None
        return None

    def get_symbol_name(self, symbol):
        try:
            if symbol.Name:
                return symbol.Name
        except:
            pass

        try:
            p = symbol.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM)
            if p:
                return p.AsString()
        except:
            pass

        return None

    def load_family_if_missing(self):
        fam = self.get_family_by_name()
        if fam:
            return fam, False

        if not os.path.exists(self.family_path):
            TaskDialog.Show("Error", "Family file not found:\n{}".format(self.family_path))
            return None, False

        t = None
        uiapp.DialogBoxShowing += shared_family_dialog_fallback
        try:
            t = Transaction(self.doc, 'Load Fixture CL Family')
            t.Start()

            fload_handler = FamilyLoaderOptionsHandler()
            loaded_family_ref = clr.StrongBox[Family]()
            result = self.doc.LoadFamily(self.family_path, fload_handler, loaded_family_ref)

            t.Commit()

            if result and loaded_family_ref.Value:
                self.family = loaded_family_ref.Value
                self.symbol_cache = {}
                return self.family, True

            self.family = None
            fallback = self.get_family_by_name()
            return fallback, False

        except Exception as e:
            if t and t.HasStarted() and not t.HasEnded():
                t.RollBack()
            TaskDialog.Show("Error", "Error loading family: {}".format(str(e)))
            return None, False

        finally:
            uiapp.DialogBoxShowing -= shared_family_dialog_fallback
            if t:
                t.Dispose()

    def build_symbol_cache(self):
        self.symbol_cache = {}

        fam = self.get_family_by_name()
        if not fam:
            return

        for symbol_id in fam.GetFamilySymbolIds():
            sym = self.doc.GetElement(symbol_id)
            if sym:
                type_name = self.get_symbol_name(sym)
                if type_name:
                    self.symbol_cache[type_name.strip().upper()] = sym

    def get_symbol_by_type_name(self, type_name):
        if not type_name:
            return None

        key = type_name.strip().upper()

        if not self.symbol_cache:
            self.build_symbol_cache()

        sym = self.symbol_cache.get(key)
        if sym and sym.IsValidObject:
            return sym

        return None

    def get_any_base_symbol(self):
        fam = self.get_family_by_name()
        if not fam:
            return None

        for sid in fam.GetFamilySymbolIds():
            sym = self.doc.GetElement(sid)
            if sym:
                return sym

        return None

    def create_type_if_missing(self, fixture_type):
        fixture_type = (fixture_type or '').strip().upper()
        if not fixture_type:
            return None

        existing = self.get_symbol_by_type_name(fixture_type)
        if existing:
            return existing

        base_symbol = self.get_any_base_symbol()
        if not base_symbol:
            TaskDialog.Show("Error", "No symbols found in family '{}'.".format(self.family_name))
            return None

        t = Transaction(self.doc, 'Create New Family Type')
        t.Start()
        try:
            new_symbol = base_symbol.Duplicate(fixture_type)

            param = new_symbol.LookupParameter("TS_Point_Description")
            if not param:
                t.RollBack()
                TaskDialog.Show("Error", "Parameter 'TS_Point_Description' not found in family type '{}'.".format(fixture_type))
                return None

            if param.IsReadOnly:
                t.RollBack()
                TaskDialog.Show("Error", "Parameter 'TS_Point_Description' is read-only in family type '{}'.".format(fixture_type))
                return None

            if param.StorageType != DB.StorageType.String:
                t.RollBack()
                TaskDialog.Show("Error", "Parameter 'TS_Point_Description' is not a text parameter in '{}'.".format(fixture_type))
                return None

            param.Set(fixture_type)
            t.Commit()

            self.symbol_cache[fixture_type] = new_symbol
            return new_symbol

        except Exception as e:
            if t.HasStarted():
                t.RollBack()
            TaskDialog.Show("Error", "Error creating new type '{}': {}".format(fixture_type, str(e)))
            return None

    def activate_symbol_if_needed(self, symbol):
        if not symbol:
            return False

        if symbol.IsActive:
            return True

        t = Transaction(self.doc, 'Activate Family Symbol')
        t.Start()
        try:
            symbol.Activate()
            self.doc.Regenerate()
            t.Commit()
            return True
        except Exception as e:
            if t.HasStarted():
                t.RollBack()
            TaskDialog.Show("Error", "Error activating symbol: {}".format(str(e)))
            return False


# --------------------------------------------------
# WPF modal dialog
# --------------------------------------------------
class FixtureTypeWindow(Window):
    def __init__(self, revit_window_handle, default_input):
        Window.__init__(self)

        self.Title = "Set Fixture Type"
        self.Width = 300
        self.Height = 150
        self.ResizeMode = ResizeMode.NoResize
        self.WindowStartupLocation = WindowStartupLocation.CenterScreen

        self.selected_type = None
        self.placement_triggered = False
        self.default_input = default_input

        self.InitializeComponents()
        WindowInteropHelper(self).Owner = revit_window_handle

    def InitializeComponents(self):
        grid = Grid()
        self.Content = grid

        grid.RowDefinitions.Add(RowDefinition(Height=System.Windows.GridLength.Auto))
        grid.RowDefinitions.Add(RowDefinition(Height=System.Windows.GridLength.Auto))
        grid.RowDefinitions.Add(RowDefinition(Height=System.Windows.GridLength.Auto))

        self.label = Label()
        self.label.Content = "Enter Fixture Type (e.g., WC-1, L-1):"
        self.label.Margin = Thickness(10, 5, 10, 5)
        Grid.SetRow(self.label, 0)
        grid.Children.Add(self.label)

        self.textbox = TextBox()
        self.textbox.Text = self.default_input
        self.textbox.Margin = Thickness(10, 0, 10, 5)
        self.textbox.TextChanged += self.on_text_changed
        Grid.SetRow(self.textbox, 1)
        grid.Children.Add(self.textbox)

        button_panel = StackPanel()
        button_panel.Orientation = Orientation.Horizontal
        button_panel.HorizontalAlignment = HorizontalAlignment.Center
        button_panel.Margin = Thickness(0, 10, 0, 10)
        Grid.SetRow(button_panel, 2)
        grid.Children.Add(button_panel)

        self.place_button = Button()
        self.place_button.Content = "Place"
        self.place_button.Width = 75
        self.place_button.Height = 25
        self.place_button.Margin = Thickness(5, 0, 5, 0)
        self.place_button.Click += self.on_place_click
        button_panel.Children.Add(self.place_button)

        self.close_button = Button()
        self.close_button.Content = "Close"
        self.close_button.Width = 75
        self.close_button.Height = 25
        self.close_button.Margin = Thickness(5, 0, 5, 0)
        self.close_button.Click += self.on_close_click
        button_panel.Children.Add(self.close_button)

    def on_text_changed(self, sender, event):
        current_text = self.textbox.Text
        uppercase_text = current_text.upper()
        if current_text != uppercase_text:
            caret_position = self.textbox.CaretIndex
            self.textbox.Text = uppercase_text
            self.textbox.CaretIndex = caret_position

    def on_place_click(self, sender, event):
        fixture_type = self.textbox.Text.strip().upper()

        if not fixture_type:
            TaskDialog.Show("Error", "Please enter a valid fixture type.")
            return

        self.selected_type = fixture_type
        self.placement_triggered = True
        self.Close()

    def on_close_click(self, sender, event):
        self.placement_triggered = False
        self.Close()


# --------------------------------------------------
# Main
# --------------------------------------------------
def main():
    family_manager = FixtureFamilyManager(doc, family_name, family_path)

    revit_window_handle = uiapp.MainWindowHandle
    form = FixtureTypeWindow(revit_window_handle, prev_input)
    form.ShowDialog()

    if not (form.placement_triggered and form.selected_type):
        return

    fixture_type = form.selected_type

    fam, newly_loaded = family_manager.load_family_if_missing()
    if not fam:
        return

    family_manager.build_symbol_cache()

    if revit_version == 2022 and newly_loaded:
        save_last_fixture_type(filepath, fixture_type)
        TaskDialog.Show(
            "Fixture CL",
            "Family loaded successfully.\n\nRevit 2022: run the command again to place the family."
        )
        return

    target_symbol = family_manager.get_symbol_by_type_name(fixture_type)
    if not target_symbol:
        target_symbol = family_manager.create_type_if_missing(fixture_type)

    if not target_symbol:
        return

    if not family_manager.activate_symbol_if_needed(target_symbol):
        return

    save_last_fixture_type(filepath, fixture_type)

    try:
        uidoc.PromptForFamilyInstancePlacement(target_symbol)
    except OperationCanceledException:
        pass
    except Exception as e:
        TaskDialog.Show("Error", "Placement error: {}".format(str(e)))


if __name__ == '__main__':
    main()