
from collections import OrderedDict

from pyrevit.coreutils import applocales
from pyrevit import revit, DB
from pyrevit import coreutils
from pyrevit import forms
from pyrevit import script


logger = script.get_logger()
output = script.get_output()

# shortcut for DB.BuiltInCategory
BIC = DB.BuiltInCategory


class RNOpts(object):
    """Renumber tool option"""
    def __init__(self, cat, by_bicat=None):
        self.bicat = cat
        self._cat = revit.query.get_category(self.bicat)
        self.by_bicat = by_bicat
        self._by_cat = revit.query.get_category(self.by_bicat) if self.by_bicat else None

    @property
    def name(self):
        """Renumber option name derived from option categories."""
        if self.by_bicat:
            applocale = applocales.get_host_applocale()
            if 'english' in applocale.lang_name.lower():
                return '{} by {}'.format(self._cat.Name, self._by_cat.Name)
            return '{} <- {}'.format(self._cat.Name, self._by_cat.Name)
        return self._cat.Name


def toggle_element_selection_handles(target_view, bicat, state=True):
    """Toggle handles for spatial elements"""
    with revit.Transaction("Toggle handles"):
        if state:
            target_view.EnableTemporaryViewPropertiesMode(target_view.Id)

        rr_cat = revit.query.get_subcategory(bicat, 'Reference')
        try:
            if rr_cat:
                rr_cat.Visible[target_view] = state
        except Exception as vex:
            logger.debug(
                'Failed changing category visibility for "%s" to "%s" on view "%s" | %s',
                bicat, state, target_view.Name, str(vex)
            )

        rr_int = revit.query.get_subcategory(bicat, 'Interior Fill')
        if not rr_int:
            rr_int = revit.query.get_subcategory(bicat, 'Interior')

        try:
            if rr_int:
                rr_int.Visible[target_view] = state
        except Exception as vex:
            logger.debug(
                'Failed changing interior fill visibility for "%s" to "%s" on view "%s" | %s',
                bicat, state, target_view.Name, str(vex)
            )

        if not state:
            target_view.DisableTemporaryViewMode(
                DB.TemporaryViewMode.TemporaryViewProperties
            )


class EasilySelectableElements(object):
    """Toggle spatial element handles for easy selection."""
    def __init__(self, target_view, bicat):
        self.supported_categories = [
            BIC.OST_Rooms,
            BIC.OST_Areas,
            BIC.OST_MEPSpaces
        ]
        self.target_view = target_view
        self.bicat = bicat

    def __enter__(self):
        if self.bicat in self.supported_categories:
            toggle_element_selection_handles(self.target_view, self.bicat)
        return self

    def __exit__(self, exception, exception_value, traceback):
        if self.bicat in self.supported_categories:
            toggle_element_selection_handles(self.target_view, self.bicat, state=False)


def increment(number):
    """Increment given item number by one."""
    return coreutils.increment_str(str(number), expand=True)


def get_number(target_element):
    """Get target element number."""
    if hasattr(target_element, "Number"):
        try:
            return target_element.Number
        except Exception as ex:
            logger.debug("Failed reading built-in Number on element %s | %s", target_element.Id, ex)

    fp_num_param = target_element.LookupParameter('FP_Valve Number')
    if fp_num_param and fp_num_param.HasValue:
        try:
            if fp_num_param.StorageType == DB.StorageType.String:
                return fp_num_param.AsString()
            elif fp_num_param.StorageType == DB.StorageType.Integer:
                return str(fp_num_param.AsInteger())
        except Exception as ex:
            logger.debug("Failed reading FP_Valve Number on element %s | %s", target_element.Id, ex)

    return None


def set_number(target_element, new_number):
    """Set target element number."""
    if hasattr(target_element, "Number"):
        target_element.Number = str(new_number)
        return

    fp_num_param = target_element.LookupParameter('FP_Valve Number')
    if not fp_num_param:
        logger.debug("FP_Valve Number not found on element %s", target_element.Id)
        return

    if fp_num_param.IsReadOnly:
        logger.debug("FP_Valve Number is read-only on element %s", target_element.Id)
        return

    try:
        if fp_num_param.StorageType == DB.StorageType.String:
            fp_num_param.Set(str(new_number))
        elif fp_num_param.StorageType == DB.StorageType.Integer:
            fp_num_param.Set(int(new_number))
        else:
            logger.debug("Unsupported storage type for FP_Valve Number on element %s", target_element.Id)
    except Exception as ex:
        logger.debug("Failed to set FP_Valve Number on element %s | %s", target_element.Id, ex)
        raise


def mark_element_as_renumbered(target_view, element):
    """Override element VG to transparent and halftone."""
    try:
        ogs = DB.OverrideGraphicSettings()
        ogs.SetHalftone(True)
        ogs.SetSurfaceTransparency(100)
        target_view.SetElementOverrides(element.Id, ogs)
    except Exception as ex:
        logger.debug("Failed to mark element %s | %s", element.Id, ex)


def unmark_renamed_elements(target_view, marked_element_ids):
    """Reset element VG to default."""
    for marked_element_id in marked_element_ids:
        try:
            ogs = DB.OverrideGraphicSettings()
            target_view.SetElementOverrides(marked_element_id, ogs)
        except Exception as ex:
            logger.debug("Failed to unmark element %s | %s", marked_element_id, ex)


def get_elements_dict(builtin_cat):
    """Collect number:id information about target elements."""
    all_elements = revit.query.get_elements_by_categories([builtin_cat])
    elements_dict = {}
    for elem in all_elements:
        try:
            num = get_number(elem)
            if num:
                elements_dict[num] = elem.Id
        except Exception as ex:
            logger.debug("Failed building number map for element %s | %s", elem.Id, ex)
    return elements_dict


def find_replacement_number(existing_number, elements_dict):
    """Find an appropriate replacement number for conflicting numbers."""
    replaced_number = increment(existing_number)
    while replaced_number in elements_dict:
        replaced_number = increment(replaced_number)
    return replaced_number


def renumber_element(target_element, new_number, elements_dict):
    """Renumber given element."""
    target_id = target_element.Id

    if new_number in elements_dict:
        element_with_same_number = revit.doc.GetElement(elements_dict[new_number])
        if element_with_same_number and element_with_same_number.Id != target_id:
            current_number = get_number(element_with_same_number)
            if current_number:
                replaced_number = find_replacement_number(current_number, elements_dict)
                set_number(element_with_same_number, replaced_number)
                elements_dict[replaced_number] = element_with_same_number.Id
                if current_number in elements_dict:
                    elements_dict.pop(current_number, None)

    existing_number = get_number(target_element)
    if existing_number in elements_dict:
        elements_dict.pop(existing_number, None)

    logger.debug('Applying %s to element %s', new_number, target_id)
    set_number(target_element, new_number)
    elements_dict[str(new_number)] = target_id

    refreshed_element = revit.doc.GetElement(target_id)
    if refreshed_element:
        mark_element_as_renumbered(revit.active_view, refreshed_element)


def ask_for_starting_number(category_name):
    """Ask user for starting number."""
    return forms.ask_for_string(
        prompt="Enter starting number",
        title="ReNumber {}".format(category_name)
    )


def _unmark_collected(category_name, renumbered_element_ids):
    with revit.Transaction("Unmark {}".format(category_name)):
        unmark_renamed_elements(revit.active_view, renumbered_element_ids)


def pick_and_renumber(rnopts, starting_index):
    """Main renumbering routine for elements of given category."""
    with revit.TransactionGroup("Renumber {}".format(rnopts.name)):
        with EasilySelectableElements(revit.active_view, rnopts.bicat):
            index = str(starting_index)
            existing_elements_data = get_elements_dict(rnopts.bicat)
            renumbered_element_ids = []

            for picked_element in revit.get_picked_elements_by_category(
                    rnopts.bicat,
                    message="Select {} in order".format(rnopts.name.lower())):

                try:
                    picked_id = picked_element.Id
                except Exception as ex:
                    logger.debug("Picked element became invalid before transaction | %s", ex)
                    continue

                with revit.Transaction("Renumber {}".format(rnopts.name)):
                    fresh_element = revit.doc.GetElement(picked_id)
                    if not fresh_element:
                        logger.debug("Element %s is no longer valid.", picked_id)
                        continue

                    renumber_element(fresh_element, index, existing_elements_data)
                    renumbered_element_ids.append(picked_id)

                index = increment(index)

            _unmark_collected(rnopts.name, renumbered_element_ids)


# ensure active view is a model view
if forms.check_modelview(revit.active_view):
    renumber_options = [
        RNOpts(cat=BIC.OST_PipeAccessory),
        RNOpts(cat=BIC.OST_FabricationPipework),
        RNOpts(cat=BIC.OST_FabricationDuctwork),
        RNOpts(cat=BIC.OST_StructuralFraming)
    ]

    if revit.active_view.ViewType == DB.ViewType.AreaPlan:
        renumber_options.insert(1, RNOpts(cat=BIC.OST_Areas))

    options_dict = OrderedDict()
    for renumber_option in renumber_options:
        options_dict[renumber_option.name] = renumber_option

    selected_option_name = forms.CommandSwitchWindow.show(
        options_dict,
        message='Pick element type to renumber:',
        width=400
    )

    if selected_option_name:
        selected_option = options_dict[selected_option_name]

        if selected_option.by_bicat:
            if selected_option.bicat == BIC.OST_Doors \
                    and selected_option.by_bicat == BIC.OST_Rooms:
                with forms.WarningBar(title='Pick Pairs of Door and Room. ESCAPE to end.'):
                    door_by_room_renumber(selected_option)
        else:
            starting_number = ask_for_starting_number(selected_option.name)
            if starting_number:
                with forms.WarningBar(
                    title='Pick {} One by One. ESCAPE to end.'.format(selected_option.name)
                ):
                    pick_and_renumber(selected_option, starting_number)