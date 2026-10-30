# -*- coding: UTF-8 -*-
import clr
clr.AddReference('System')
import System
import System.Diagnostics
clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')

from System.Windows import Window, Thickness, HorizontalAlignment, WindowStartupLocation, ResizeMode
from System.Windows.Controls import Grid, RowDefinition, Label, TextBox, Button, StackPanel, Orientation
from System.Windows.Interop import WindowInteropHelper

from Autodesk.Revit import DB
from Autodesk.Revit.DB import FilteredElementCollector, Transaction, Family, FamilyInstance, ViewType
from Autodesk.Revit.UI import TaskDialog
from Autodesk.Revit.UI.Events import TaskDialogShowingEventArgs

import os
import sys

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
uiapp = __revit__
revit_version = int(doc.Application.VersionNumber)

path, filename = os.path.split(__file__)
family_file = '\\Control Point.rfa'
family_path = path + family_file

FamilyName = 'Control Point'
TypeName = 'CP'

# Save file
folder_name = "c:\\temp"
filepath = os.path.join(folder_name, 'Ribbon_Control Point.txt')

if not os.path.exists(folder_name):
    os.makedirs(folder_name)

default_point_number = ""
default_point_name = ""

if os.path.exists(filepath):
    try:
        with open(filepath, 'r') as f:
            lines = f.read().splitlines()
            if len(lines) > 0:
                default_point_number = lines[0].strip()
            if len(lines) > 1:
                default_point_name = lines[1].strip()
    except:
        pass


class FamilyLoaderOptionsHandler(DB.IFamilyLoadOptions):
    def OnFamilyFound(self, familyInUse, overwriteParameterValues):
        overwriteParameterValues.Value = False
        return True

    def OnSharedFamilyFound(self, sharedFamily, familyInUse, source, overwriteParameterValues):
        source.Value = DB.FamilySource.Project
        overwriteParameterValues.Value = False
        return True


def get_family(document, family_name):
    families = FilteredElementCollector(document).OfClass(Family)
    for fam in families:
        if fam.Name == family_name:
            return fam
    return None


def get_symbol_name(symbol):
    try:
        return symbol.Name
    except:
        pass

    try:
        param = symbol.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM)
        if param:
            return param.AsString()
    except:
        pass

    return None


def get_family_symbol_by_name(document, family_name, type_name):
    family = get_family(document, family_name)
    if not family:
        return None

    for symbol_id in family.GetFamilySymbolIds():
        symbol = document.GetElement(symbol_id)
        if symbol:
            symbol_name = get_symbol_name(symbol)
            if symbol_name and symbol_name.upper() == type_name.upper():
                return symbol

    return None


def get_instance_ids_for_symbol(document, symbol_id):
    ids = set()
    collector = FilteredElementCollector(document).OfClass(FamilyInstance)
    for inst in collector:
        try:
            if inst.Symbol and inst.Symbol.Id == symbol_id:
                ids.add(inst.Id.IntegerValue)
        except:
            pass
    return ids


def set_text_parameter(element, param_name, value):
    param = element.LookupParameter(param_name)
    if not param:
        return "Parameter '{}' not found.".format(param_name)
    if param.IsReadOnly:
        return "Parameter '{}' is read-only.".format(param_name)
    if param.StorageType != DB.StorageType.String:
        return "Parameter '{}' is not a text parameter.".format(param_name)

    param.Set(value)
    return None


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


def load_family_if_missing(document, family_name, family_path):
    family = get_family(document, family_name)
    if family:
        return family, False

    if not os.path.exists(family_path):
        TaskDialog.Show("Error", "Family file not found:\n{}".format(family_path))
        return None, False

    fload_handler = FamilyLoaderOptionsHandler()
    loaded_family_ref = clr.StrongBox[DB.Family]()

    uiapp.DialogBoxShowing += shared_family_dialog_fallback
    try:
        t = Transaction(document, 'Load Control Point Family')
        t.Start()
        try:
            loaded_ok = document.LoadFamily(family_path, fload_handler, loaded_family_ref)
            t.Commit()
        except Exception as e:
            if t.HasStarted():
                t.RollBack()
            TaskDialog.Show("Error", "Error loading family: {}".format(str(e)))
            return None, False
    finally:
        uiapp.DialogBoxShowing -= shared_family_dialog_fallback

    family = get_family(document, family_name)
    if not family and loaded_family_ref.Value:
        family = loaded_family_ref.Value

    return family, True if family else False


class ControlPointWindow(Window):
    def __init__(self, revit_window_handle, default_number, default_name):
        Window.__init__(self)

        self.Title = "Set Control Point Data"
        self.Width = 320
        self.Height = 200
        self.ResizeMode = ResizeMode.NoResize
        self.WindowStartupLocation = WindowStartupLocation.CenterScreen

        self.point_number = default_number
        self.point_name = default_name
        self.confirmed = False

        self.InitializeComponents()
        WindowInteropHelper(self).Owner = revit_window_handle

    def InitializeComponents(self):
        grid = Grid()
        self.Content = grid

        row_definitions = [
            RowDefinition(Height=System.Windows.GridLength.Auto),
            RowDefinition(Height=System.Windows.GridLength.Auto),
            RowDefinition(Height=System.Windows.GridLength.Auto),
            RowDefinition(Height=System.Windows.GridLength.Auto),
            RowDefinition(Height=System.Windows.GridLength.Auto)
        ]
        for row in row_definitions:
            grid.RowDefinitions.Add(row)

        row_index = 0

        self.label_number = Label()
        self.label_number.Content = "Point Number:"
        self.label_number.Margin = Thickness(10, 5, 10, 5)
        Grid.SetRow(self.label_number, row_index)
        grid.Children.Add(self.label_number)
        row_index += 1

        self.textbox_number = TextBox()
        self.textbox_number.Text = self.point_number
        self.textbox_number.Margin = Thickness(10, 0, 10, 5)
        Grid.SetRow(self.textbox_number, row_index)
        grid.Children.Add(self.textbox_number)
        row_index += 1

        self.label_name = Label()
        self.label_name.Content = "Point Description:"
        self.label_name.Margin = Thickness(10, 5, 10, 5)
        Grid.SetRow(self.label_name, row_index)
        grid.Children.Add(self.label_name)
        row_index += 1

        self.textbox_name = TextBox()
        self.textbox_name.Text = self.point_name
        self.textbox_name.Margin = Thickness(10, 0, 10, 5)
        Grid.SetRow(self.textbox_name, row_index)
        grid.Children.Add(self.textbox_name)
        row_index += 1

        button_panel = StackPanel()
        button_panel.Orientation = Orientation.Horizontal
        button_panel.HorizontalAlignment = HorizontalAlignment.Center
        button_panel.Margin = Thickness(0, 10, 0, 10)
        Grid.SetRow(button_panel, row_index)
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

    def on_place_click(self, sender, event):
        self.point_number = self.textbox_number.Text.strip()
        self.point_name = self.textbox_name.Text.strip()

        if not self.point_number:
            TaskDialog.Show("Error", "Please enter a Point Number.")
            return

        if not self.point_name:
            TaskDialog.Show("Error", "Please enter a Point Name.")
            return

        self.confirmed = True
        self.Close()

    def on_close_click(self, sender, event):
        self.confirmed = False
        self.Close()


def main():
    # Check active view type
    view = uidoc.ActiveView
    if view.ViewType == ViewType.ThreeD:
        TaskDialog.Show("Error", "Cannot use in 3D view.")
        return

    # Show dialog first
    revit_window_handle = uiapp.MainWindowHandle
    form = ControlPointWindow(revit_window_handle, default_point_number, default_point_name)
    form.ShowDialog()

    if not form.confirmed:
        return

    # Save entered values
    try:
        with open(filepath, 'w') as f:
            f.write(form.point_number + '\n')
            f.write(form.point_name)
    except:
        pass

    # Load family if needed
    family, newly_loaded = load_family_if_missing(doc, FamilyName, family_path)
    if not family:
        return

    # Revit 2022 only: force rerun after load
    if revit_version == 2022 and newly_loaded:
        TaskDialog.Show(
            "Control Point",
            "Family loaded successfully.\n\nRevit 2022: run the command again to place the family."
        )
        return

    target_symbol = get_family_symbol_by_name(doc, FamilyName, TypeName)
    if not target_symbol:
        TaskDialog.Show("Error", "Type '{}' was not found in family '{}'.".format(TypeName, FamilyName))
        return

    # Activate symbol if needed
    if not target_symbol.IsActive:
        t = Transaction(doc, 'Activate Control Point Type')
        t.Start()
        try:
            target_symbol.Activate()
            doc.Regenerate()
            t.Commit()
        except Exception as e:
            TaskDialog.Show("Error", "Error activating type '{}': {}".format(TypeName, str(e)))
            if t.HasStarted():
                t.RollBack()
            return

    # Track placed instances
    before_ids = get_instance_ids_for_symbol(doc, target_symbol.Id)

    try:
        uidoc.PromptForFamilyInstancePlacement(target_symbol)
    except Exception:
        pass

    after_ids = get_instance_ids_for_symbol(doc, target_symbol.Id)
    new_ids = after_ids - before_ids

    if not new_ids:
        return

    # Write values to newly placed instances
    t = Transaction(doc, 'Set Control Point Parameters')
    t.Start()
    try:
        errors = []

        for int_id in new_ids:
            elem = doc.GetElement(DB.ElementId(int_id))
            if not elem:
                continue

            err1 = set_text_parameter(elem, "TS_Point_Number", form.point_number)
            err2 = set_text_parameter(elem, "TS_Point_Description", form.point_name)

            if err1:
                errors.append(err1)
            if err2:
                errors.append(err2)

        t.Commit()

        if errors:
            unique_errors = sorted(set(errors))
            TaskDialog.Show("Warning", "\n".join(unique_errors))

    except Exception as e:
        if t.HasStarted():
            t.RollBack()
        TaskDialog.Show("Error", "Error setting control point parameters: {}".format(str(e)))


if __name__ == '__main__':
    main()