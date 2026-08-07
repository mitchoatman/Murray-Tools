# -*- coding: utf-8 -*-
from Autodesk.Revit.DB import Transaction, FabricationPart, FabricationConfiguration, TransactionStatus
from Autodesk.Revit.UI.Selection import ObjectType
from Autodesk.Revit.UI import TaskDialog
from System import Array

# WPF Namespace Imports (No external XAML required)
import clr
clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')
from System.Windows import Window, Thickness, WindowStyle, ResizeMode, WindowStartupLocation, HorizontalAlignment
from System.Windows.Controls import Label, ComboBox, Button, Grid, RowDefinition

uidoc = __revit__.ActiveUIDocument
doc = uidoc.Document

def change_fab_part_specs():
    try:
        # 1. Select multiple fabrication parts in the model
        references = uidoc.Selection.PickObjects(ObjectType.Element, "Select Fabrication Parts to update spec")
        
        if not references:
            print("Operation cancelled by user.")
            return
            
        fab_parts = []
        for ref in references:
            element = doc.GetElement(ref.ElementId)
            if isinstance(element, FabricationPart):
                fab_parts.append(element)
                
        if not fab_parts:
            TaskDialog.Show("Error", "No valid Fabrication Parts were selected.")
            return
            
        # Use the first part as a reference to pull compatible specifications from the config
        sample_part = fab_parts[0]
        config = FabricationConfiguration.GetFabricationConfiguration(doc)
        if not config:
            TaskDialog.Show("Error", "No fabrication configuration loaded in this project.")
            return
            
        spec_ids = config.GetAllSpecifications(sample_part)
        
        # Build a mapping of Specification Name -> ID
        spec_dict = {}
        for spec_id in spec_ids:
            spec_name = config.GetSpecificationName(spec_id)
            if spec_name:
                spec_dict[spec_name] = spec_id
                
        sorted_spec_names = sorted(spec_dict.keys())
        
        if not sorted_spec_names:
            TaskDialog.Show("Error", "No compatible specifications found for the selected parts.")
            return

        current_spec_id = sample_part.Specification
        current_spec_name = next((name for name, spec_id in spec_dict.items() if spec_id == current_spec_id), "Unknown")

        # 2. Native WPF Form (matching your architectural style)
        class SpecSelectForm(Window):
            def __init__(self, spec_names, current_name, part_count):
                self.Title = "Select Target Specification"
                self.Width = 320
                self.Height = 160
                self.WindowStyle = WindowStyle.SingleBorderWindow
                self.ResizeMode = ResizeMode.NoResize
                self.WindowStartupLocation = WindowStartupLocation.CenterScreen
                self.result = None

                grid = Grid()
                grid.Margin = Thickness(12)

                for _ in range(3):
                    grid.RowDefinitions.Add(RowDefinition())

                self.Content = grid

                label = Label()
                label.Content = "Current Spec: {} | Update {} part(s):".format(current_name, part_count)
                label.Margin = Thickness(0, -2, 0, 4)
                Grid.SetRow(label, 0)
                grid.Children.Add(label)

                self.combo = ComboBox()
                self.combo.ItemsSource = Array[object](spec_names)
                self.combo.SelectedIndex = 0
                self.combo.Height = 25
                self.combo.Margin = Thickness(0, 0, 0, 10)
                Grid.SetRow(self.combo, 1)
                grid.Children.Add(self.combo)

                ok_button = Button()
                ok_button.Content = "Update Specification"
                ok_button.Width = 140
                ok_button.Height = 28
                ok_button.HorizontalAlignment = HorizontalAlignment.Center
                ok_button.Click += self.on_ok
                Grid.SetRow(ok_button, 2)
                grid.Children.Add(ok_button)
                self.combo.Focus()

            def on_ok(self, sender, args):
                self.result = self.combo.SelectedItem
                self.DialogResult = True
                self.Close()

        # Show the custom WPF form
        form = SpecSelectForm(sorted_spec_names, current_spec_name, len(fab_parts))
        
        if form.ShowDialog() and form.DialogResult:
            selected_name = form.result
            if not selected_name:
                print("Operation cancelled by user.")
                return
                
            new_spec_id = spec_dict[selected_name]

            # 3. Execute the batch specification update inside a single transaction
            success_count = 0
            failed_count = 0
            
            t = Transaction(doc, "Batch Update Fabrication Specifications")
            try:
                t.Start()
                for fab_part in fab_parts:
                    try:
                        fab_part.Specification = new_spec_id
                        success_count += 1
                    except:
                        failed_count += 1
                t.Commit()
                
                TaskDialog.Show(
                    "Batch Complete", 
                    "Successfully updated spec to '{}' for {} part(s).{}".format(
                        selected_name, 
                        success_count, 
                        "\n({} part(s) failed or were incompatible)".format(failed_count) if failed_count > 0 else ""
                    )
                )
                
            except Exception as trans_ex:
                if t.GetStatus() == TransactionStatus.Started:
                    t.RollBack()
                TaskDialog.Show(
                    "Transaction Failed", 
                    "Revit rejected the batch specification change.\n\nReason: {}".format(trans_ex)
                )
        else:
            TaskDialog.Show("Selection Cancelled", "No specification selected.")

    except Exception as ex:
        if "Canceled" not in str(ex):
            print("An error occurred: {}".format(ex))

if __name__ == "__main__":
    change_fab_part_specs()