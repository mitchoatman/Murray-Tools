# -*- coding: utf-8 -*-
from Autodesk.Revit.DB import Transaction, SubTransaction, FabricationPart, FabricationConfiguration, TransactionStatus
from Autodesk.Revit.UI.Selection import ObjectType
from Autodesk.Revit.UI import TaskDialog
from System import Array

import clr
clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')
from System.Windows import Window, Thickness, WindowStyle, ResizeMode, WindowStartupLocation, HorizontalAlignment
from System.Windows.Controls import Label, ComboBox, Button, Grid, RowDefinition

uidoc = __revit__.ActiveUIDocument
doc = uidoc.Document

def change_fab_part_materials():
    try:
        references = uidoc.Selection.PickObjects(ObjectType.Element, "Select Fabrication Parts to update material")
        if not references:
            return
            
        fab_parts = [doc.GetElement(ref.ElementId) for ref in references if isinstance(doc.GetElement(ref.ElementId), FabricationPart)]
        if not fab_parts:
            TaskDialog.Show("Error", "No valid Fabrication Parts selected.")
            return
            
        sample_part = fab_parts[0]
        config = FabricationConfiguration.GetFabricationConfiguration(doc)
        if not config:
            TaskDialog.Show("Error", "No fabrication configuration loaded.")
            return
            
        # Get compatible materials for this specific part type safely
        material_ids = config.GetAllMaterials(sample_part)

        material_dict = {}
        for mat_id in material_ids:
            mat_name = config.GetMaterialName(mat_id)
            if mat_name:
                material_dict[mat_name] = mat_id
                
        # Fallback to global if item-specific list comes back empty
        if len(material_dict) <= 1:
            try:
                for m_id in config.GetAllMaterials():
                    m_name = config.GetMaterialName(m_id)
                    if m_name:
                        material_dict[m_name] = m_id
            except:
                pass

        sorted_material_names = sorted(material_dict.keys())
        if not sorted_material_names:
            TaskDialog.Show("Error", "No compatible materials found.")
            return

        current_material_id = sample_part.Material
        current_material_name = config.GetMaterialName(current_material_id) if hasattr(config, "GetMaterialName") else "Unknown"

        class MaterialSelectForm(Window):
            def __init__(self, material_names, current_name, part_count):
                self.Title = "Select Compatible Material"
                self.Width = 350
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
                label.Content = "Current Material: {} | Update {} part(s):".format(current_name, part_count)
                label.Margin = Thickness(0, -2, 0, 4)
                Grid.SetRow(label, 0)
                grid.Children.Add(label)

                self.combo = ComboBox()
                self.combo.ItemsSource = Array[object](material_names)
                if current_name in material_names:
                    self.combo.SelectedItem = current_name
                else:
                    self.combo.SelectedIndex = 0
                self.combo.Height = 25
                self.combo.Margin = Thickness(0, 0, 0, 10)
                Grid.SetRow(self.combo, 1)
                grid.Children.Add(self.combo)

                ok_button = Button()
                ok_button.Content = "Update Material"
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

        form = MaterialSelectForm(sorted_material_names, current_material_name, len(fab_parts))
        if form.ShowDialog() and form.DialogResult:
            selected_name = form.result
            if not selected_name:
                return
            new_material_id = material_dict[selected_name]

            success_count = 0
            failed_count = 0
            
            # Wrap in a single main Transaction, using SubTransactions per part 
            # to prevent a single bad item from crashing the whole batch or locking Revit up.
            main_t = Transaction(doc, "Batch Update Fabrication Materials")
            try:
                main_t.Start()
                for fab_part in fab_parts:
                    sub_t = SubTransaction(doc)
                    sub_t.Start()
                    try:
                        fab_part.Material = new_material_id
                        sub_t.Commit()
                        success_count += 1
                    except:
                        sub_t.RollBack()
                        failed_count += 1
                main_t.Commit()
                
                TaskDialog.Show(
                    "Batch Complete", 
                    "Successfully updated material to '{}' for {} part(s).{}".format(
                        selected_name, 
                        success_count, 
                        "\n({} part(s) skipped due to database mapping limits)".format(failed_count) if failed_count > 0 else ""
                    )
                )
                
            except Exception as trans_ex:
                if main_t.GetStatus() == TransactionStatus.Started:
                    main_t.RollBack()
                TaskDialog.Show("Transaction Failed", "Reason: {}".format(trans_ex))

    except Exception as ex:
        if "Canceled" not in str(ex):
            print("An error occurred: {}".format(ex))

if __name__ == "__main__":
    change_fab_part_materials()