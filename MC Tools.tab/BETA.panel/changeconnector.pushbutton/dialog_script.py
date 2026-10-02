# -*- coding: utf-8 -*-

from Autodesk.Revit.DB import (
    Transaction, SubTransaction, FabricationPart, FabricationConfiguration,
    TransactionStatus, ConnectorDomainType, ConnectorProfileType
)
from Autodesk.Revit.UI.Selection import ObjectType
from Autodesk.Revit.UI import TaskDialog, TaskDialogCommonButtons, TaskDialogResult
from System import Array

import clr
clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')
from System.Windows import (
    Window, Thickness, WindowStyle, ResizeMode, WindowStartupLocation,
    HorizontalAlignment, GridLength, GridUnitType
)
from System.Windows.Controls import (
    Label, ComboBox, Button, Grid, RowDefinition, ColumnDefinition,
    ScrollViewer, StackPanel, Orientation
)

uidoc = __revit__.ActiveUIDocument
doc = uidoc.Document

NO_CHANGE_LABEL = "-- No Change --"


# ---------------------------------------------------------------------------
# Fabrication connector helpers (unchanged from prior version)
# ---------------------------------------------------------------------------

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


def get_connected_partner(conn):
    """
    [CONFIRM] see module docstring.
    """
    try:
        if not conn.IsConnected:
            return None
        for other in conn.AllRefs:
            try:
                if other.Owner is not None and other.Owner.Id != conn.Owner.Id:
                    return other
            except:
                pass
    except:
        pass
    return None


# ---------------------------------------------------------------------------
# Dialog: one row per connector end, each with its own new-type dropdown
# ---------------------------------------------------------------------------

class MultiEndConnectorForm(Window):
    """
    end_defs: list of dicts, one per connector end on the sample part, sorted
    by FabricationIndex. Each dict:
        {
            'fab_index':   int,
            'display':     "C1",
            'current_id':  int,
            'current_name':str,
            'state':       "Connected" | "Open",
            'label_to_id': {label_string: connector_id_or_None, ...}
        }
    Result after ShowDialog(): self.result = {fab_index: new_connector_id}
    containing ONLY the ends where the user picked something other than
    "No Change".
    """

    def __init__(self, end_defs, sample_part_name):
        self.end_defs = end_defs
        self.combos = {}
        self.result = None

        self.Title = "Update Fabrication Connectors"
        self.Width = 640
        row_height = 44
        self.Height = 400
        self.WindowStyle = WindowStyle.SingleBorderWindow
        self.ResizeMode = ResizeMode.NoResize
        self.WindowStartupLocation = WindowStartupLocation.CenterScreen

        outer = Grid()
        outer.Margin = Thickness(12)
        outer.RowDefinitions.Add(RowDefinition())               # header
        rows_host_row = RowDefinition()
        rows_host_row.Height = GridLength(1, GridUnitType.Star)
        outer.RowDefinitions.Add(rows_host_row)                 # scrollable rows
        outer.RowDefinitions.Add(RowDefinition())                # button row
        self.Content = outer

        header = Label()
        header.Content = (
            "Ends and current values shown are from: {}\n"
            "Set a new type only for the end(s) you want to change; leave "
            "others as 'No Change'.".format(sample_part_name)
        )
        header.Margin = Thickness(0, 0, 0, 6)
        Grid.SetRow(header, 0)
        outer.Children.Add(header)

        scroll = ScrollViewer()
        scroll.VerticalScrollBarVisibility = 0
        rows_panel = StackPanel()
        rows_panel.Orientation = Orientation.Vertical
        scroll.Content = rows_panel
        Grid.SetRow(scroll, 1)
        outer.Children.Add(scroll)

        for end_def in self.end_defs:
            row = Grid()
            row.Margin = Thickness(0, 4, 0, 4)
            col_label = ColumnDefinition()
            col_label.Width = GridLength(260)
            col_combo = ColumnDefinition()
            col_combo.Width = GridLength(1, GridUnitType.Star)
            row.ColumnDefinitions.Add(col_label)
            row.ColumnDefinitions.Add(col_combo)

            lbl = Label()
            lbl.Content = "{} | Current: {} [{}] | {}".format(
                end_def['display'], end_def['current_name'],
                end_def['current_id'], end_def['state']
            )
            Grid.SetColumn(lbl, 0)
            row.Children.Add(lbl)

            combo = ComboBox()
            labels = sorted(end_def['label_to_id'].keys())
            # Keep "No Change" pinned first regardless of sort.
            labels = [NO_CHANGE_LABEL] + [l for l in labels if l != NO_CHANGE_LABEL]
            combo.ItemsSource = Array[object](labels)
            combo.SelectedIndex = 0  # default to No Change
            combo.Height = 25
            combo.Margin = Thickness(6, 0, 0, 0)
            Grid.SetColumn(combo, 1)
            row.Children.Add(combo)

            self.combos[end_def['fab_index']] = (combo, end_def)
            rows_panel.Children.Add(row)

        ok_button = Button()
        ok_button.Content = "Apply"
        ok_button.Width = 160
        ok_button.Height = 28
        ok_button.Margin = Thickness(0, 10, 0, 0)
        ok_button.HorizontalAlignment = HorizontalAlignment.Center
        ok_button.Click += self.on_ok
        Grid.SetRow(ok_button, 2)
        outer.Children.Add(ok_button)

    def on_ok(self, sender, args):
        chosen = {}
        for fab_index, (combo, end_def) in self.combos.items():
            selected_label = combo.SelectedItem
            if not selected_label or selected_label == NO_CHANGE_LABEL:
                continue
            new_id = end_def['label_to_id'].get(selected_label)
            if new_id:
                chosen[fab_index] = new_id
        self.result = chosen
        self.DialogResult = True
        self.Close()


# ---------------------------------------------------------------------------
# Main routine
# ---------------------------------------------------------------------------

def change_fab_part_connectors():
    try:
        references = uidoc.Selection.PickObjects(
            ObjectType.Element,
            "Select Fabrication Parts to update connector(s)"
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

        sample_connectors = sorted(sample_connectors, key=lambda x: x[1].FabricationIndex)

        # Build one end_def per connector on the sample part.
        end_defs = []
        for conn, fab_info in sample_connectors:
            current_id = fab_info.BodyConnectorId
            current_name = get_connector_name(config, current_id)
            state = "Connected" if conn.IsConnected else "Open"

            candidate_ids = get_candidate_connector_ids(config, fab_info)
            label_to_id = {NO_CHANGE_LABEL: None}
            for cid in candidate_ids:
                cname = get_connector_name(config, cid)
                label_to_id["{} [{}]".format(cname, cid)] = cid

            end_defs.append({
                'fab_index': fab_info.FabricationIndex,
                'display': "C{}".format(fab_info.FabricationIndex + 1),
                'current_id': current_id,
                'current_name': current_name,
                'state': state,
                'label_to_id': label_to_id,
            })

        try:
            sample_name = sample_part.get_Parameter(
                __import__("Autodesk.Revit.DB", fromlist=["BuiltInParameter"]).BuiltInParameter.ALL_MODEL_TYPE_NAME
            ).AsString() or "Selected part 1"
        except:
            sample_name = "Selected part 1"

        form = MultiEndConnectorForm(end_defs, sample_name)
        if not (form.ShowDialog() and form.DialogResult):
            return

        end_changes = form.result  # {fab_index: new_connector_id}
        if not end_changes:
            TaskDialog.Show("Nothing To Do", "No connector ends were changed from 'No Change'.")
            return

        # --- Determine, per targeted end, how many parts have a connected joint there ---
        connected_count_by_end = {}
        for fab_index in end_changes:
            count = 0
            for fab_part in fab_parts:
                for conn, fab_info in get_fabrication_connectors(fab_part):
                    if fab_info.FabricationIndex == fab_index and conn.IsConnected:
                        count += 1
                        break
            connected_count_by_end[fab_index] = count

        any_connected = any(c > 0 for c in connected_count_by_end.values())

        # --- Required confirmation before disconnecting anything ---
        if any_connected:
            confirm = TaskDialog("Confirm Disconnect")
            confirm.MainInstruction = "This will change {} connector end(s) across {} selected part(s).".format(
                len(end_changes), len(fab_parts)
            )
            detail_lines = []
            for fab_index, new_id in sorted(end_changes.items()):
                detail_lines.append(
                    "  C{}: -> {}  ({} connected joint(s) will be disconnected/reconnected)".format(
                        fab_index + 1, get_connector_name(config, new_id),
                        connected_count_by_end.get(fab_index, 0)
                    )
                )
            confirm.MainContent = (
                "The fabrication connector type on a connected end cannot be changed directly "
                "through the API. For each connected end above, this tool will:\n\n"
                "  1. Disconnect the joint\n"
                "  2. Change the connector type\n"
                "  3. Attempt to reconnect to the original mating part\n\n"
                "Reconnection is not guaranteed if the new connector type has a different "
                "shape or size than the mating connector. Any end that fails to reconnect "
                "will be left OPEN and reported at the end.\n\n"
                + "\n".join(detail_lines) +
                "\n\nDo you want to continue?"
            )
            confirm.CommonButtons = TaskDialogCommonButtons.Yes | TaskDialogCommonButtons.No
            confirm.DefaultButton = TaskDialogResult.No

            if confirm.Show() != TaskDialogResult.Yes:
                return

        # --- Apply ---
        success_by_end = {}
        reconnected_by_end = {}
        left_open_by_end = {}
        skipped_no_end_by_end = {}
        failed_parts = 0

        main_t = Transaction(doc, "Batch Update Fabrication Connectors")
        try:
            main_t.Start()

            for fab_part in fab_parts:
                sub_t = SubTransaction(doc)
                sub_t.Start()
                try:
                    conn_by_index = {}
                    for conn, fab_info in get_fabrication_connectors(fab_part):
                        conn_by_index[fab_info.FabricationIndex] = (conn, fab_info)

                    for fab_index, new_connector_id in end_changes.items():
                        if fab_index not in conn_by_index:
                            skipped_no_end_by_end[fab_index] = skipped_no_end_by_end.get(fab_index, 0) + 1
                            continue

                        target_connector, target_fab_info = conn_by_index[fab_index]
                        was_connected = target_connector.IsConnected
                        partner_conn = None

                        if was_connected:
                            partner_conn = get_connected_partner(target_connector)
                            if partner_conn is None:
                                # Reported connected but partner could not be resolved;
                                # skip this end on this part rather than guess.
                                continue
                            target_connector.DisconnectFrom(partner_conn)

                        target_fab_info.BodyConnectorId = new_connector_id
                        success_by_end[fab_index] = success_by_end.get(fab_index, 0) + 1

                        if was_connected and partner_conn is not None:
                            try:
                                target_connector.ConnectTo(partner_conn)
                                if target_connector.IsConnected:
                                    reconnected_by_end[fab_index] = reconnected_by_end.get(fab_index, 0) + 1
                                else:
                                    left_open_by_end[fab_index] = left_open_by_end.get(fab_index, 0) + 1
                            except:
                                left_open_by_end[fab_index] = left_open_by_end.get(fab_index, 0) + 1

                    sub_t.Commit()

                except:
                    sub_t.RollBack()
                    failed_parts += 1

            main_t.Commit()

            # --- Summary ---
            summary_lines = []
            for fab_index, new_id in sorted(end_changes.items()):
                cname = get_connector_name(config, new_id)
                n = success_by_end.get(fab_index, 0)
                summary_lines.append("C{}: updated to '{}' on {} part(s).".format(
                    fab_index + 1, cname, n
                ))
                if reconnected_by_end.get(fab_index):
                    summary_lines.append(
                        "    {} previously-connected joint(s) reconnected.".format(
                            reconnected_by_end[fab_index]
                        )
                    )
                if left_open_by_end.get(fab_index):
                    summary_lines.append(
                        "    {} joint(s) could NOT be reconnected (mismatched type/size) "
                        "and are now OPEN.".format(left_open_by_end[fab_index])
                    )
                if skipped_no_end_by_end.get(fab_index):
                    summary_lines.append(
                        "    {} part(s) skipped: no C{} on that part.".format(
                            skipped_no_end_by_end[fab_index], fab_index + 1
                        )
                    )

            if failed_parts > 0:
                summary_lines.append(
                    "{} part(s) were skipped entirely due to an error.".format(failed_parts)
                )

            TaskDialog.Show("Batch Complete", "\n".join(summary_lines))

        except Exception as trans_ex:
            if main_t.GetStatus() == TransactionStatus.Started:
                main_t.RollBack()
            TaskDialog.Show("Transaction Failed", "Reason: {}".format(trans_ex))

    except Exception as ex:
        if "Canceled" not in str(ex):
            print("An error occurred: {}".format(ex))


if __name__ == "__main__":
    change_fab_part_connectors()