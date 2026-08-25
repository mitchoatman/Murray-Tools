# -*- coding: utf-8 -*-
from Autodesk.Revit.DB import (
    Transaction, SubTransaction, FabricationPart, FabricationConfiguration,
    TransactionStatus, ConnectorDomainType, ConnectorProfileType
)
from Autodesk.Revit.UI.Selection import ObjectType
from Autodesk.Revit.UI import TaskDialog
from System import Array

import clr
clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')
from System.Windows import Window, Thickness, WindowStyle, ResizeMode, WindowStartupLocation, HorizontalAlignment
from System.Windows.Controls import Label, ComboBox, Button, Grid, RowDefinition, TextBox

uidoc = __revit__.ActiveUIDocument
doc = uidoc.Document


def get_fabrication_connectors(part):
    results = []
    try:
        if not part or not part.ConnectorManager:
            return results

        for conn in part.ConnectorManager.Connectors:
            try:
                fab_info = conn.GetFabricationConnectorInfo()
                if fab_info and fab_info.IsValid():
                    results.append((conn, fab_info))
            except:
                pass
    except:
        pass
    return results


def get_connector_name(config, connector_id):
    try:
        if connector_id and connector_id > 0:
            return config.GetFabricationConnectorName(connector_id)
    except:
        pass
    return "None"


def get_candidate_connector_ids(config, fab_info):
    """
    Older-version-safe connector list:
    use config.GetAllFabricationConnectorDefinitions(domain, shape)
    based on the current connector's fabrication connector id.
    """
    try:
        current_id = fab_info.BodyConnectorId
        if current_id and current_id > 0:
            domain = config.GetFabricationConnectorDomain(current_id)
            shape = config.GetFabricationConnectorShape(current_id)
        else:
            domain = ConnectorDomainType.Undefined
            shape = ConnectorProfileType.Invalid

        ids = config.GetAllFabricationConnectorDefinitions(domain, shape)
        return [i for i in ids if i and i > 0]
    except:
        return []


class SelectFromListForm(Window):
    def __init__(self, title_text, label_text, item_names, button_text="OK"):
        self.Title = title_text
        self.Width = 420
        self.Height = 170
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
        label.Content = label_text
        label.Margin = Thickness(0, -2, 0, 4)
        Grid.SetRow(label, 0)
        grid.Children.Add(label)

        self.combo = ComboBox()
        self.combo.ItemsSource = Array[object](item_names)
        self.combo.SelectedIndex = 0 if item_names else -1
        self.combo.Height = 25
        self.combo.Margin = Thickness(0, 0, 0, 10)
        Grid.SetRow(self.combo, 1)
        grid.Children.Add(self.combo)

        ok_button = Button()
        ok_button.Content = button_text
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


class SearchableSelectForm(Window):
    def __init__(self, title_text, label_text, item_names, button_text="OK"):
        self.Title = title_text
        self.Width = 420
        self.Height = 210
        self.WindowStyle = WindowStyle.SingleBorderWindow
        self.ResizeMode = ResizeMode.NoResize
        self.WindowStartupLocation = WindowStartupLocation.CenterScreen
        self.result = None
        self.all_items = list(item_names)

        grid = Grid()
        grid.Margin = Thickness(12)
        for _ in range(4):
            grid.RowDefinitions.Add(RowDefinition())
        self.Content = grid

        label = Label()
        label.Content = label_text
        label.Margin = Thickness(0, -2, 0, 4)
        Grid.SetRow(label, 0)
        grid.Children.Add(label)

        self.search_box = TextBox()
        self.search_box.Height = 25
        self.search_box.Margin = Thickness(0, 0, 0, 6)
        self.search_box.TextChanged += self.on_search_changed
        Grid.SetRow(self.search_box, 1)
        grid.Children.Add(self.search_box)

        self.combo = ComboBox()
        self.combo.ItemsSource = Array[object](self.all_items)
        self.combo.SelectedIndex = 0 if self.all_items else -1
        self.combo.Height = 25
        self.combo.Margin = Thickness(0, 0, 0, 10)
        Grid.SetRow(self.combo, 2)
        grid.Children.Add(self.combo)

        ok_button = Button()
        ok_button.Content = button_text
        ok_button.Width = 140
        ok_button.Height = 28
        ok_button.HorizontalAlignment = HorizontalAlignment.Center
        ok_button.Click += self.on_ok
        Grid.SetRow(ok_button, 3)
        grid.Children.Add(ok_button)

        self.search_box.Focus()

    def on_search_changed(self, sender, args):
        search_text = self.search_box.Text.lower().strip()
        if not search_text:
            filtered = self.all_items
        else:
            filtered = [x for x in self.all_items if search_text in x.lower()]

        self.combo.ItemsSource = Array[object](filtered)
        self.combo.SelectedIndex = 0 if filtered else -1

    def on_ok(self, sender, args):
        self.result = self.combo.SelectedItem
        self.DialogResult = True
        self.Close()


def change_fab_part_connectors():
    try:
        references = uidoc.Selection.PickObjects(
            ObjectType.Element,
            "Select Fabrication Parts to update connector"
        )
        if not references:
            return

        fab_parts = [doc.GetElement(ref.ElementId) for ref in references
                     if isinstance(doc.GetElement(ref.ElementId), FabricationPart)]
        if not fab_parts:
            TaskDialog.Show("Error", "No valid Fabrication Parts selected.")
            return

        config = FabricationConfiguration.GetFabricationConfiguration(doc)
        if not config:
            TaskDialog.Show("Error", "No fabrication configuration loaded.")
            return

        sample_part = fab_parts[0]
        sample_connectors = get_fabrication_connectors(sample_part)
        if not sample_connectors:
            TaskDialog.Show("Error", "No valid fabrication connectors found on the sample part.")
            return

        connector_choice_map = {}
        sorted_sample_connectors = sorted(sample_connectors, key=lambda x: x[1].FabricationIndex)

        for conn, fab_info in sorted_sample_connectors:
            current_id = fab_info.BodyConnectorId
            current_name = get_connector_name(config, current_id)
            state = "Connected" if conn.IsConnected else "Open"
            display_connector = "C{}".format(fab_info.FabricationIndex + 1)
            label = "{} | Current: {} [{}] | {}".format(
                display_connector, current_name, current_id, state
            )
            connector_choice_map[label] = fab_info.FabricationIndex

        connector_labels = sorted(connector_choice_map.keys())

        end_form = SelectFromListForm(
            "Select Connector End",
            "Choose which fabrication connector/end to update:",
            connector_labels,
            "Next"
        )

        if not (end_form.ShowDialog() and end_form.DialogResult):
            return

        selected_end_label = end_form.result
        if not selected_end_label:
            return

        selected_fab_index = connector_choice_map[selected_end_label]

        sample_connector = None
        sample_fab_info = None
        for conn, fab_info in sample_connectors:
            if fab_info.FabricationIndex == selected_fab_index:
                sample_connector = conn
                sample_fab_info = fab_info
                break

        if not sample_connector or not sample_fab_info:
            TaskDialog.Show("Error", "Could not resolve the selected connector/end.")
            return

        valid_connector_ids = get_candidate_connector_ids(config, sample_fab_info)

        connector_def_map = {}
        for conn_id in valid_connector_ids:
            conn_name = get_connector_name(config, conn_id)
            display = "{} [{}]".format(conn_name, conn_id)
            connector_def_map[display] = conn_id

        connector_def_labels = sorted(connector_def_map.keys())
        if not connector_def_labels:
            TaskDialog.Show("Error", "No compatible fabrication connector definitions found.")
            return

        current_body_id = sample_fab_info.BodyConnectorId
        current_display = None
        for k, v in connector_def_map.items():
            if v == current_body_id:
                current_display = k
                break

        select_form = SearchableSelectForm(
            "Select Connector Type",
            "Current connector on {}: {}".format(
                "C{}".format(selected_fab_index + 1),
                get_connector_name(config, current_body_id)
            ),
            connector_def_labels,
            "Update Connector"
        )

        if current_display and current_display in connector_def_labels:
            select_form.combo.SelectedItem = current_display

        if select_form.ShowDialog() and select_form.DialogResult:
            selected_connector_label = select_form.result
            if not selected_connector_label:
                return

            new_connector_id = connector_def_map[selected_connector_label]

            success_count = 0
            failed_count = 0

            main_t = Transaction(doc, "Batch Update Fabrication Connectors")
            try:
                main_t.Start()

                for fab_part in fab_parts:
                    sub_t = SubTransaction(doc)
                    sub_t.Start()
                    try:
                        target_connector = None
                        target_fab_info = None

                        for conn, fab_info in get_fabrication_connectors(fab_part):
                            if fab_info.FabricationIndex == selected_fab_index:
                                target_connector = conn
                                target_fab_info = fab_info
                                break

                        if not target_connector or not target_fab_info:
                            sub_t.RollBack()
                            failed_count += 1
                            continue

                        target_fab_info.BodyConnectorId = new_connector_id
                        sub_t.Commit()
                        success_count += 1

                    except:
                        sub_t.RollBack()
                        failed_count += 1

                main_t.Commit()

                TaskDialog.Show(
                    "Batch Complete",
                    "Updated connector on {} to '{}' for {} part(s).{}".format(
                        "C{}".format(selected_fab_index + 1),
                        get_connector_name(config, new_connector_id),
                        success_count,
                        "\n({} part(s) skipped due to connectivity/compatibility limits)".format(failed_count)
                        if failed_count > 0 else ""
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
    change_fab_part_connectors()