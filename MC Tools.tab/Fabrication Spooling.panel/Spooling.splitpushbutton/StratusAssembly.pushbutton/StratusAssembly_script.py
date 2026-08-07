# -*- coding: utf-8 -*-
__persistentengine__ = True

import os
import clr

from Parameters.Add_SharedParameters import Shared_Params
from Parameters.Get_Set_Params import set_parameter_by_name

Shared_Params()

clr.AddReference('PresentationCore')
clr.AddReference('PresentationFramework')
clr.AddReference('WindowsBase')
clr.AddReference('System')
clr.AddReference('System.Windows.Forms')

from System.Windows import Window, Thickness, WindowStartupLocation, ResizeMode, FontWeights
from System.Windows.Controls import Button, TextBox, Label, Grid, RowDefinition, ColumnDefinition
from System.Windows.Media import FontFamily
from System.Windows.Interop import WindowInteropHelper
from System.Windows.Forms import MessageBox

from Autodesk.Revit.DB import Transaction, FabricationConfiguration
from Autodesk.Revit.UI import TaskDialog, ExternalEvent, IExternalEventHandler

uiapp = __revit__

folder_name = r"c:\Temp"
filepath = os.path.join(folder_name, "Ribbon_StratusAssembly.txt")


def ensure_file():
    if not os.path.exists(folder_name):
        os.makedirs(folder_name)

    if not os.path.exists(filepath):
        with open(filepath, 'w') as f:
            f.writelines([
                "L1-A1-CW-01\n",
                "L1-A1-HGR-MAP\n"
            ])


def read_defaults():
    ensure_file()
    with open(filepath, 'r') as f:
        lines = [x.rstrip() for x in f.readlines()]

    while len(lines) < 2:
        lines.append("")

    return lines[0], lines[1]


def write_defaults(spool_name, map_name):
    ensure_file()
    with open(filepath, 'w') as f:
        f.writelines([
            spool_name + "\n",
            map_name + "\n"
        ])


def increment_spool_name(value):
    parts = value.rsplit('-', 1)
    if len(parts) == 2 and parts[1].isdigit():
        width = len(parts[1])
        next_num = int(parts[1]) + 1
        return "{}-{}".format(parts[0], str(next_num).zfill(width))
    return value


class SetSpoolHandler(IExternalEventHandler):
    def __init__(self):
        self.spool_name = ""
        self.map_name = ""
        self.form = None

    def Execute(self, app):
        try:
            uidoc = app.ActiveUIDocument
            if uidoc is None:
                TaskDialog.Show("Spool Data", "No active document.")
                return

            doc = uidoc.Document
            FabricationConfiguration.GetFabricationConfiguration(doc)
            selection_ids = list(uidoc.Selection.GetElementIds())

            if not selection_ids:
                TaskDialog.Show("Spool Data", "No elements selected.")
                return

            missing_elements = []

            t = Transaction(doc, "Set Spool Data")
            t.Start()

            for eid in selection_ids:
                el = doc.GetElement(eid)
                if el is None:
                    continue

                try:
                    param_exist = el.LookupParameter("STRATUS Assembly")
                    isfabpart = el.LookupParameter("Fabrication Service")

                    if isfabpart:
                        set_parameter_by_name(el, "STRATUS Assembly", self.spool_name)
                        set_parameter_by_name(el, "FP_Spool Map", self.map_name)
                        set_parameter_by_name(el, "STRATUS Status", "Modeled")

                        try:
                            el.SpoolName = self.spool_name
                        except:
                            pass

                        try:
                            el.PartStatus = 1
                        except:
                            pass

                        try:
                            el.Pinned = True
                        except:
                            pass

                    elif param_exist:
                        set_parameter_by_name(el, "STRATUS Assembly", self.spool_name)
                        set_parameter_by_name(el, "FP_Spool Map", self.map_name)
                        set_parameter_by_name(el, "STRATUS Status", "Modeled")

                        try:
                            el.Pinned = True
                        except:
                            pass

                    else:
                        missing_elements.append(str(eid.IntegerValue))

                except Exception as inner_ex:
                    missing_elements.append("{} : {}".format(eid.IntegerValue, str(inner_ex)))

            t.Commit()

            if missing_elements:
                td = TaskDialog("Missing Parameters Warning")
                td.MainInstruction = "Some selected elements were skipped."
                td.MainContent = "The following Element IDs were missing required parameters or failed to update:"
                td.ExpandedContent = "\n".join(missing_elements)
                td.Show()

            new_spool_name = increment_spool_name(self.spool_name)
            write_defaults(new_spool_name, self.map_name)

            if self.form:
                self.form.spool_box.Text = new_spool_name
                self.form.map_box.Text = self.map_name

        except Exception as ex:
            TaskDialog.Show("Execute Error", str(ex))

    def GetName(self):
        return "Set Spool Data Handler"


class TXT_Form(Window):
    def __init__(self, ext_event, handler):
        self.ext_event = ext_event
        self.handler = handler
        self.handler.form = self

        self.Title = "Spool Data"
        self.Width = 320
        self.Height = 170
        self.MinWidth = 320
        self.MinHeight = 170
        self.WindowStartupLocation = WindowStartupLocation.CenterScreen
        self.ResizeMode = ResizeMode.NoResize
        self.Topmost = True
        self.ShowInTaskbar = False

        try:
            helper = WindowInteropHelper(self)
            helper.Owner = uiapp.MainWindowHandle
        except:
            pass

        spool_name, map_name = read_defaults()

        grid = Grid()
        for i in range(3):
            grid.RowDefinitions.Add(RowDefinition())
        for i in range(2):
            grid.ColumnDefinitions.Add(ColumnDefinition())

        l1 = Label()
        l1.Content = "Spool Name:"
        l1.FontFamily = FontFamily("Arial")
        l1.FontSize = 14
        l1.FontWeight = FontWeights.Bold
        l1.Margin = Thickness(5, 5, 0, 0)
        Grid.SetRow(l1, 0)
        Grid.SetColumn(l1, 0)
        grid.Children.Add(l1)

        self.spool_box = TextBox()
        self.spool_box.Text = spool_name
        self.spool_box.Margin = Thickness(-10, 5, 5, 0)
        self.spool_box.FontFamily = FontFamily("Arial")
        self.spool_box.FontSize = 12
        self.spool_box.Height = 22
        Grid.SetRow(self.spool_box, 0)
        Grid.SetColumn(self.spool_box, 1)
        grid.Children.Add(self.spool_box)

        l2 = Label()
        l2.Content = "Map Name:"
        l2.FontFamily = FontFamily("Arial")
        l2.FontSize = 14
        l2.FontWeight = FontWeights.Bold
        l2.Margin = Thickness(5, 5, 0, 0)
        Grid.SetRow(l2, 1)
        Grid.SetColumn(l2, 0)
        grid.Children.Add(l2)

        self.map_box = TextBox()
        self.map_box.Text = map_name
        self.map_box.Margin = Thickness(-10, 5, 5, 0)
        self.map_box.FontFamily = FontFamily("Arial")
        self.map_box.FontSize = 12
        self.map_box.Height = 22
        Grid.SetRow(self.map_box, 1)
        Grid.SetColumn(self.map_box, 1)
        grid.Children.Add(self.map_box)

        btn = Button()
        btn.Content = "Set Spool Data"
        btn.Margin = Thickness(80, 10, 80, 5)
        btn.FontFamily = FontFamily("Arial")
        btn.FontSize = 12
        btn.Height = 28
        btn.Click += self.on_click
        Grid.SetRow(btn, 2)
        Grid.SetColumnSpan(btn, 2)
        grid.Children.Add(btn)

        self.Content = grid
        self.Closed += self.on_closed

        self.spool_box.Focus()
        self.spool_box.SelectAll()

    def on_click(self, sender, args):
        try:
            spool = self.spool_box.Text.strip()
            mapname = self.map_box.Text.strip()

            if not spool or not mapname:
                MessageBox.Show("Enter both Spool Name and Map Name.", "Input Error")
                return

            self.handler.spool_name = spool
            self.handler.map_name = mapname
            self.ext_event.Raise()

        except Exception as ex:
            MessageBox.Show(str(ex), "Button Error")

    def on_closed(self, sender, args):
        try:
            self.ext_event.Dispose()
        except:
            pass

        globals()["_spool_form"] = None
        globals()["_spool_handler"] = None
        globals()["_spool_event"] = None


try:
    if globals().get("_spool_form", None):
        try:
            if _spool_form.IsVisible:
                _spool_form.Activate()
            else:
                _spool_form.Show()
        except:
            pass
    else:
        _spool_handler = SetSpoolHandler()
        _spool_event = ExternalEvent.Create(_spool_handler)
        _spool_form = TXT_Form(_spool_event, _spool_handler)
        _spool_form.Show()

except Exception as ex:
    TaskDialog.Show("Startup Error", str(ex))