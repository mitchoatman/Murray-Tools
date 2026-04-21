# -*- coding: utf-8 -*-
from Autodesk.Revit.DB import FilteredElementCollector, Revision, ViewSheet, BuiltInParameter, Transaction, ElementId
from Autodesk.Revit.UI import TaskDialog
import clr, sys
clr.AddReference('PresentationCore')
clr.AddReference('PresentationFramework')
clr.AddReference('WindowsBase')
clr.AddReference('System')
from System.Windows import (
    Application, Window, Thickness, HorizontalAlignment, VerticalAlignment,
    ResizeMode, WindowStartupLocation, GridLength, GridUnitType, FontWeights
)
from System.Windows.Controls import (
    Button, TextBox, CheckBox, Grid, RowDefinition, ColumnDefinition,
    Label, StackPanel, ScrollViewer, Orientation, ScrollBarVisibility
)
from System.Windows.Media import Brushes

app = __revit__.Application
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

# Collect non-issued revisions
all_revisions = FilteredElementCollector(doc).OfClass(Revision).ToElements()
non_issued_revisions = [r for r in all_revisions if not r.Issued]
non_issued_revisions.sort(key=lambda r: r.SequenceNumber)

if not non_issued_revisions:
    TaskDialog.Show("No Revisions", "No non-issued revisions found in the project.")
    sys.exit()

# Revision display format: Sequence - Description (Date)
def get_revision_display(rev):
    seq = str(rev.SequenceNumber).zfill(2) if rev.SequenceNumber else "??"
    desc = rev.Description or "(No Description)"
    date = rev.RevisionDate or "(No Date)"
    return "{} - {} ({})".format(seq, desc, date)

revision_items = [(rev, get_revision_display(rev)) for rev in non_issued_revisions]

# Collect all sheets (including placeholders)
all_sheets = FilteredElementCollector(doc).OfClass(ViewSheet).ToElements()
all_sheets = sorted(all_sheets, key=lambda s: s.SheetNumber)

def get_sheet_display(sheet):
    num = sheet.SheetNumber
    name = sheet.Name
    return "{} - {}".format(num, name)

sheet_items = [(s, get_sheet_display(s)) for s in all_sheets]

class RevisionSheetSelectionForm(object):
    def __init__(self, revision_items, sheet_items):
        self.all_revision_items = revision_items
        self.all_sheet_items = sheet_items
        
        self.revision_checked = {r.Id.IntegerValue: False for r, _ in revision_items}
        self.sheet_checked = {s.Id.IntegerValue: False for s, _ in sheet_items}
        
        self.revision_check_all_state = False
        self.sheet_check_all_state = False
        
        self.revision_checkboxes = []
        self.sheet_checkboxes = []
        
        self.InitializeComponents()
    
    def InitializeComponents(self):
        self._window = Window()
        self._window.Title = "Set Revision on Sheets"
        self._window.Width = 700
        self._window.Height = 800
        self._window.WindowStartupLocation = WindowStartupLocation.CenterScreen
        self._window.ResizeMode = ResizeMode.CanResizeWithGrip
        
        main_grid = Grid()
        main_grid.Margin = Thickness(10)
        
        # Row definitions
        rows = [
            GridLength.Auto,                    # 0 - Revision label
            GridLength.Auto,                    # 1 - Revision search + All/None
            GridLength(1, GridUnitType.Star),   # 2 - Revision list
            GridLength.Auto,                    # 3 - Sheets label
            GridLength.Auto,                    # 4 - Sheets search + All/None
            GridLength(2, GridUnitType.Star),   # 5 - Sheets list (larger)
            GridLength.Auto                     # 6 - Buttons
        ]
        for h in rows:
            rd = RowDefinition()
            rd.Height = h
            main_grid.RowDefinitions.Add(rd)
        
        # Main grid has one stretching column
        main_col = ColumnDefinition()
        main_col.Width = GridLength(1, GridUnitType.Star)
        main_grid.ColumnDefinitions.Add(main_col)
        
        # === Revisions section ===
        rev_label = Label()
        rev_label.Content = "Select non-issued revisions (multiple allowed):"
        rev_label.FontSize = 14
        rev_label.FontWeight = FontWeights.Bold
        Grid.SetRow(rev_label, 0)
        main_grid.Children.Add(rev_label)
        
        # Revision search row
        rev_search_grid = Grid()
        rev_search_grid.Margin = Thickness(0, 5, 0, 5)
        
        # Columns: stretch for textbox, auto for button
        rev_col_star = ColumnDefinition()
        rev_col_star.Width = GridLength(1, GridUnitType.Star)
        rev_search_grid.ColumnDefinitions.Add(rev_col_star)
        
        rev_col_auto = ColumnDefinition()
        rev_col_auto.Width = GridLength.Auto
        rev_search_grid.ColumnDefinitions.Add(rev_col_auto)
        
        self.revision_search_box = TextBox()
        self.revision_search_box.Height = 25
        self.revision_search_box.Margin = Thickness(0, 0, 10, 0)
        self.revision_search_box.TextChanged += self.revision_search_changed
        Grid.SetColumn(self.revision_search_box, 0)
        rev_search_grid.Children.Add(self.revision_search_box)
        
        self.revision_all_button = Button()
        self.revision_all_button.Content = "All / None"
        self.revision_all_button.Width = 100
        self.revision_all_button.Margin = Thickness(0, 0, 0, 0)
        self.revision_all_button.Click += self.revision_all_clicked
        Grid.SetColumn(self.revision_all_button, 1)
        rev_search_grid.Children.Add(self.revision_all_button)
        
        Grid.SetRow(rev_search_grid, 1)
        main_grid.Children.Add(rev_search_grid)
        
        # Revision list
        self.revision_panel = StackPanel()
        rev_scroll = ScrollViewer()
        rev_scroll.Content = self.revision_panel
        rev_scroll.VerticalScrollBarVisibility = ScrollBarVisibility.Auto
        Grid.SetRow(rev_scroll, 2)
        main_grid.Children.Add(rev_scroll)
        
        self.update_revision_checkboxes(self.all_revision_items)
        
        # === Sheets section ===
        sheet_label = Label()
        sheet_label.Content = "Select sheets (including placeholders):"
        sheet_label.FontSize = 14
        sheet_label.FontWeight = FontWeights.Bold
        sheet_label.Margin = Thickness(0, 20, 0, 10)
        Grid.SetRow(sheet_label, 3)
        main_grid.Children.Add(sheet_label)
        
        # Sheet search row
        sheet_search_grid = Grid()
        sheet_search_grid.Margin = Thickness(0, 5, 0, 5)
        
        # Columns: stretch for textbox, auto for button
        sheet_col_star = ColumnDefinition()
        sheet_col_star.Width = GridLength(1, GridUnitType.Star)
        sheet_search_grid.ColumnDefinitions.Add(sheet_col_star)
        
        sheet_col_auto = ColumnDefinition()
        sheet_col_auto.Width = GridLength.Auto
        sheet_search_grid.ColumnDefinitions.Add(sheet_col_auto)
        
        self.sheet_search_box = TextBox()
        self.sheet_search_box.Height = 25
        self.sheet_search_box.Margin = Thickness(0, 0, 10, 0)
        self.sheet_search_box.TextChanged += self.sheet_search_changed
        Grid.SetColumn(self.sheet_search_box, 0)
        sheet_search_grid.Children.Add(self.sheet_search_box)
        
        self.sheet_all_button = Button()
        self.sheet_all_button.Content = "All / None"
        self.sheet_all_button.Width = 100
        self.sheet_all_button.Margin = Thickness(0, 0, 0, 0)
        self.sheet_all_button.Click += self.sheet_all_clicked
        Grid.SetColumn(self.sheet_all_button, 1)
        sheet_search_grid.Children.Add(self.sheet_all_button)
        
        Grid.SetRow(sheet_search_grid, 4)
        main_grid.Children.Add(sheet_search_grid)
        
        # Sheet list
        self.sheet_panel = StackPanel()
        sheet_scroll = ScrollViewer()
        sheet_scroll.Content = self.sheet_panel
        sheet_scroll.VerticalScrollBarVisibility = ScrollBarVisibility.Auto
        Grid.SetRow(sheet_scroll, 5)
        main_grid.Children.Add(sheet_scroll)
        
        self.update_sheet_checkboxes(self.all_sheet_items)
        
        # === Buttons ===
        button_panel = StackPanel()
        button_panel.Orientation = Orientation.Horizontal
        button_panel.HorizontalAlignment = HorizontalAlignment.Center
        button_panel.Margin = Thickness(0, 20, 0, 10)
        
        cancel_btn = Button()
        cancel_btn.Content = "Cancel"
        cancel_btn.Width = 120
        cancel_btn.Height = 25
        cancel_btn.Margin = Thickness(10)
        cancel_btn.Click += self.cancel_clicked
        button_panel.Children.Add(cancel_btn)
        
        apply_btn = Button()
        apply_btn.Content = "Apply"
        apply_btn.Width = 120
        apply_btn.Height = 25
        apply_btn.Margin = Thickness(10)
        apply_btn.Click += self.apply_clicked
        button_panel.Children.Add(apply_btn)
        
        Grid.SetRow(button_panel, 6)
        main_grid.Children.Add(button_panel)
        
        self._window.Content = main_grid
    
    # Update methods
    def update_revision_checkboxes(self, items):
        self.revision_panel.Children.Clear()
        self.revision_checkboxes = []
        for rev, display in items:
            cb = CheckBox()
            cb.Content = display
            cb.Tag = rev
            cb.IsChecked = self.revision_checked.get(rev.Id.IntegerValue, False)
            cb.Checked += self.revision_checkbox_changed
            cb.Unchecked += self.revision_checkbox_changed
            cb.Margin = Thickness(4)
            self.revision_panel.Children.Add(cb)
            self.revision_checkboxes.append(cb)
    
    def update_sheet_checkboxes(self, items):
        self.sheet_panel.Children.Clear()
        self.sheet_checkboxes = []
        for sheet, display in items:
            cb = CheckBox()
            cb.Content = display
            cb.Tag = sheet
            cb.IsChecked = self.sheet_checked.get(sheet.Id.IntegerValue, False)
            cb.Checked += self.sheet_checkbox_changed
            cb.Unchecked += self.sheet_checkbox_changed
            cb.Margin = Thickness(4)
            self.sheet_panel.Children.Add(cb)
            self.sheet_checkboxes.append(cb)
    
    # Search handlers
    def revision_search_changed(self, sender, args):
        text = sender.Text.lower()
        filtered = [item for item in self.all_revision_items if text in item[1].lower()]
        self.update_revision_checkboxes(filtered)
    
    def sheet_search_changed(self, sender, args):
        text = sender.Text.lower()
        filtered = [item for item in self.all_sheet_items if text in item[1].lower()]
        self.update_sheet_checkboxes(filtered)
    
    # All/None handlers
    def revision_all_clicked(self, sender, args):
        self.revision_check_all_state = not self.revision_check_all_state
        for cb in self.revision_checkboxes:
            cb.IsChecked = self.revision_check_all_state
            self.revision_checked[cb.Tag.Id.IntegerValue] = self.revision_check_all_state
    
    def sheet_all_clicked(self, sender, args):
        self.sheet_check_all_state = not self.sheet_check_all_state
        for cb in self.sheet_checkboxes:
            cb.IsChecked = self.sheet_check_all_state
            self.sheet_checked[cb.Tag.Id.IntegerValue] = self.sheet_check_all_state
    
    # Checkbox changed handlers
    def revision_checkbox_changed(self, sender, args):
        cb = sender
        rev = cb.Tag
        self.revision_checked[rev.Id.IntegerValue] = cb.IsChecked == True
    
    def sheet_checkbox_changed(self, sender, args):
        cb = sender
        sheet = cb.Tag
        self.sheet_checked[sheet.Id.IntegerValue] = cb.IsChecked == True
    
    # Button handlers
    def cancel_clicked(self, sender, args):
        self._window.Close()
    
    def apply_clicked(self, sender, args):
        selected_rev_ids = [eid for eid, checked in self.revision_checked.items() if checked]
        selected_sheet_ids = [eid for eid, checked in self.sheet_checked.items() if checked]
        
        if not selected_rev_ids:
            TaskDialog.Show("Selection Required", "Please select at least one revision.")
            return
        if not selected_sheet_ids:
            TaskDialog.Show("Selection Required", "Please select at least one sheet.")
            return
        
        selected_revisions = [doc.GetElement(ElementId(eid)) for eid in selected_rev_ids]
        selected_sheets = [doc.GetElement(ElementId(eid)) for eid in selected_sheet_ids]
        
        try:
            from pyrevit import revit
            with revit.Transaction('Set Revision on Sheets'):
                updated_sheets = revit.update.update_sheet_revisions(selected_revisions, selected_sheets)
            
            if updated_sheets:
                TaskDialog.Show("Message", 'Revisions Applied')
            else:
                TaskDialog.Show("Message", 'No sheets were updated')
                
        except Exception as ex:
            TaskDialog.Show("Error", str(ex))
        
        self._window.Close()
    
    def ShowDialog(self):
        self._window.ShowDialog()

# Launch the custom dialog
form = RevisionSheetSelectionForm(revision_items, sheet_items)
form.ShowDialog()