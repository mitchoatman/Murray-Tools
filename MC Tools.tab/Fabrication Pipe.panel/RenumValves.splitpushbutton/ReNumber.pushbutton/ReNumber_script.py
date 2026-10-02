import os
import re
import clr
clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')
from System.Windows import Window, Thickness, WindowStyle, ResizeMode, WindowStartupLocation, HorizontalAlignment
from System.Windows.Controls import Label, TextBox, Button, StackPanel

from Autodesk.Revit.DB import *
from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter

# Get document and active view directly from the Revit services API
uiapp = __revit__
uidoc = uiapp.ActiveUIDocument
doc = uidoc.Document
active_view = doc.ActiveView

# Setup persistence file path in C:\Temp
folder_name = "c:\\Temp"
filepath = os.path.join(folder_name, 'Ribbon_ValveNumber.txt')

if not os.path.exists(folder_name):
    os.makedirs(folder_name)

if not os.path.exists(filepath):
    with open(filepath, 'w') as f:
        f.write('1')

# Read last used/next suggested number
with open(filepath, 'r') as f:
    last_number = f.read().strip()
    if not last_number:
        last_number = "1"


class NumberInputForm(Window):
    """Simple WPF input form using StackPanel for starting number."""
    def __init__(self, initial_value="1"):
        self.Title = "ReNumber FP_Valve Number"
        self.Width = 300
        self.Height = 170
        self.WindowStyle = WindowStyle.SingleBorderWindow
        self.ResizeMode = ResizeMode.NoResize
        self.WindowStartupLocation = WindowStartupLocation.CenterScreen
        self.result_value = None

        panel = StackPanel()
        panel.Margin = Thickness(10)
        self.Content = panel

        label = Label()
        label.Content = "Enter starting number:"
        label.Margin = Thickness(0, 0, 0, 5)
        panel.Children.Add(label)

        self.textbox = TextBox()
        self.textbox.Text = str(initial_value)
        self.textbox.Height = 26
        self.textbox.Margin = Thickness(0, 0, 0, 10)
        panel.Children.Add(self.textbox)

        ok_button = Button()
        ok_button.Content = "OK"
        ok_button.Width = 75
        ok_button.Height = 25
        ok_button.HorizontalAlignment = HorizontalAlignment.Center
        ok_button.Click += self.ok_clicked
        panel.Children.Add(ok_button)

        self.textbox.Focus()
        self.textbox.SelectAll()

    def ok_clicked(self, sender, args):
        if self.textbox.Text.strip():
            self.result_value = self.textbox.Text.strip()
            self.DialogResult = True
            self.Close()
        else:
            self.textbox.Text = "1"


class AnyElementFilter(ISelectionFilter):
    """Allows picking any element."""
    def AllowElement(self, elem):
        return True

    def AllowReference(self, reference, position):
        return True


def increment(number_str):
    """Native string incrementer without external dependencies."""
    match = re.search(r'\d+$', number_str)
    if match:
        num_str = match.group(0)
        next_num_str = str(int(num_str) + 1).zfill(len(num_str))
        return number_str[:match.start()] + next_num_str
    return number_str + "1"


def set_number(target_element, new_number):
    """Set target element FP_Valve Number."""
    fp_num_param = target_element.LookupParameter('FP_Valve Number')
    if not fp_num_param:
        logger.debug("FP_Valve Number not found on element %s", target_element.Id)
        return

    if fp_num_param.IsReadOnly:
        logger.debug("FP_Valve Number is read-only on element %s", target_element.Id)
        return

    try:
        if fp_num_param.StorageType == StorageType.String:
            fp_num_param.Set(str(new_number))
        elif fp_num_param.StorageType == StorageType.Integer:
            fp_num_param.Set(int(new_number))
        else:
            logger.debug("Unsupported storage type for FP_Valve Number on element %s", target_element.Id)
    except Exception as ex:
        logger.debug("Failed to set FP_Valve Number on element %s | %s", target_element.Id, ex)
        raise


def mark_element_as_renumbered(target_view, element):
    """Override element VG to transparent and halftone."""
    try:
        ogs = OverrideGraphicSettings()
        ogs.SetHalftone(True)
        ogs.SetSurfaceTransparency(100)
        target_view.SetElementOverrides(element.Id, ogs)
    except Exception as ex:
        logger.debug("Failed to mark element %s | %s", element.Id, ex)


def unmark_renamed_elements(target_view, marked_element_ids):
    """Reset element VG to default."""
    for marked_element_id in marked_element_ids:
        try:
            ogs = OverrideGraphicSettings()
            target_view.SetElementOverrides(marked_element_id, ogs)
        except Exception as ex:
            logger.debug("Failed to unmark element %s | %s", marked_element_id, ex)


def pick_and_renumber(starting_index):
    """Main sequential renumbering loop with visual tracking using native transactions."""
    index = str(starting_index)
    sel_filter = AnyElementFilter()
    renumbered_element_ids = []
    last_written_index = starting_index

    while True:
        try:
            ref = uidoc.Selection.PickObject(
                ObjectType.Element, 
                sel_filter, 
                "Select element to renumber (ESC to finish)"
            )
            if not ref:
                break

            picked_id = ref.ElementId
            
            # Use native Revit API Transaction directly
            t = Transaction(doc, "Set FP_Valve Number")
            t.Start()
            try:
                fresh_element = doc.GetElement(picked_id)
                if fresh_element:
                    set_number(fresh_element, index)
                    mark_element_as_renumbered(active_view, fresh_element)
                    renumbered_element_ids.append(picked_id)
                    last_written_index = index
                t.Commit()
            except Exception as ex:
                if t.HasStarted():
                    t.RollBack()
                logger.debug("Failed transaction: %s", ex)

            index = increment(index)

        except Exception:
            # User pressed ESC or canceled selection
            break

    # Unmark/restore elements when done using native transaction
    if renumbered_element_ids:
        t_unmark = Transaction(doc, "Unmark renumbered elements")
        t_unmark.Start()
        try:
            unmark_renamed_elements(active_view, renumbered_element_ids)
            t_unmark.Commit()
        except Exception:
            if t_unmark.HasStarted():
                t_unmark.RollBack()
                
        # Save the next sequential number after the last one written
        next_suggested_number = increment(last_written_index)
        with open(filepath, 'w') as f:
            f.write(str(next_suggested_number))


# Ensure active view is a valid model view before running (using native API)
if active_view and not active_view.IsTemplate and active_view.ViewType not in [ViewType.DrawingSheet, ViewType.Schedule, ViewType.Report]:
    form = NumberInputForm(last_number)
    starting_number = None
    if form.ShowDialog() and form.DialogResult:
        starting_number = form.result_value
    
    if starting_number:
        pick_and_renumber(starting_number)
else:
    TaskDialog.Show("ReNumber", "Please open a valid model view to run this tool.")