"""ReNumber Fabrication elements in order of selection."""

__title__ = 'Renumber'


#pylint: disable=import-error,invalid-name,broad-except
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
        self._by_cat = revit.query.get_category(self.by_bicat)

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
        # if view has template, toggle temp VG overrides
        if state:
            target_view.EnableTemporaryViewPropertiesMode(target_view.Id)

        rr_cat = revit.query.get_subcategory(bicat, 'Reference')
        try:
            rr_cat.Visible[target_view] = state
        except Exception as vex:
            logger.debug(
                'Failed changing category visibility for \"%s\" '
                'to \"%s\" on view \"%s\" | %s',
                bicat,
                state,
                target_view.Name,
                str(vex)
                )
        rr_int = revit.query.get_subcategory(bicat, 'Interior Fill')
        if not rr_int:
            rr_int = revit.query.get_subcategory(bicat, 'Interior')
        try:
            rr_int.Visible[target_view] = state
        except Exception as vex:
            logger.debug(
                'Failed changing interior fill visibility for \"%s\" '
                'to \"%s\" on view \"%s\" | %s',
                bicat,
                state,
                target_view.Name,
                str(vex)
                )
        # disable the temp VG overrides after making changes to categories
        if not state:
            target_view.DisableTemporaryViewMode(
                DB.TemporaryViewMode.TemporaryViewProperties)


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
            toggle_element_selection_handles(
                self.target_view,
                self.bicat
                )
        return self

    def __exit__(self, exception, exception_value, traceback):
        if self.bicat in self.supported_categories:
            toggle_element_selection_handles(
                self.target_view,
                self.bicat,
                state=False
                )


def increment(number):
    """Increment given item number by one."""
    return coreutils.increment_str(number, expand=True)


def get_number(target_element):
    """Get target elemnet number (might be from Number or other fields)"""
    if hasattr(target_element, "Number"):
        return target_element.Number

    # determine target parameter
    mark_param = target_element.Parameter[DB.BuiltInParameter.ALL_MODEL_MARK]
    fnum_param = target_element.LookupParameter ('Item Number')
    fp_num_param = target_element.LookupParameter ('FP_Valve Number')
    if isinstance(target_element, (DB.Level, DB.Grid)):
        mark_param = target_element.Parameter[DB.BuiltInParameter.DATUM_TEXT]
    # get now
    if mark_param:
        return mark_param.AsString()


def set_number(target_element, new_number):
    """Set target elemnet number (might be at Number or other fields)"""
    if hasattr(target_element, "Number"):
        target_element.Number = new_number
        return

    # determine target parameter
    mark_param = target_element.Parameter[DB.BuiltInParameter.ALL_MODEL_MARK]
    fnum_param = target_element.LookupParameter ('Item Number')
    fp_num_param = target_element.LookupParameter ('FP_Valve Number')
    try:
        if mark_param:
            mark_param.Set(new_number)
        if fnum_param:
            fnum_param.Set(new_number)
        if fp_num_param:
            fp_num_param.Set(new_number)
    except:
        pass


def mark_element_as_renumbered(target_view, room):
    """Override element VG to transparent and halftone.

    Intended to mark processed renumbered elements visually.
    """
    ogs = DB.OverrideGraphicSettings()
    ogs.SetHalftone(True)
    ogs.SetSurfaceTransparency(100)
    target_view.SetElementOverrides(room.Id, ogs)


def unmark_renamed_elements(target_view, marked_element_ids):
    """Rest element VG to default."""
    for marked_element_id in marked_element_ids:
        ogs = DB.OverrideGraphicSettings()
        target_view.SetElementOverrides(marked_element_id, ogs)


def get_elements_dict(builtin_cat):
    """Collect number:id information about target elements."""
    all_elements = \
        revit.query.get_elements_by_categories([builtin_cat])
    return {get_number(x):x.Id for x in all_elements}


def find_replacement_number(existing_number, elements_dict):
    """Find an appropriate replacement number for conflicting numbers."""
    replaced_number = increment(existing_number)
    while replaced_number in elements_dict:
        replaced_number = increment(replaced_number)
    return replaced_number


def renumber_element(target_element, new_number, elements_dict):
    """Renumber given element."""
    # check if elements with same number exists
    if new_number in elements_dict:
        element_with_same_number = \
            revit.doc.GetElement(elements_dict[new_number])
        # make sure its not the same as target_element
        if element_with_same_number \
                and element_with_same_number.Id != target_element.Id:
            # replace its number with something else that is not conflicting
            current_number = get_number(element_with_same_number)
            replaced_number = \
                find_replacement_number(current_number, elements_dict)
            set_number(element_with_same_number, replaced_number)
            # record the element with its new number for later renumber jobs
            elements_dict[replaced_number] = element_with_same_number.Id

    # check if target element is already listed
    # remove the existing number entry since we are renumbering
    existing_number = get_number(target_element)
    if existing_number in elements_dict:
        elements_dict.pop(existing_number)

    # renumber the given element
    logger.debug('applying %s', new_number)
    set_number(target_element, new_number)
    elements_dict[new_number] = target_element.Id
    # mark the element visually to renumbered
    mark_element_as_renumbered(revit.active_view, target_element)


def ask_for_starting_number(category_name):
    """Ask user for starting number."""
    return forms.ask_for_string(
        prompt="Enter starting number",
        title="ReNumber {}".format(category_name)
        )


def _unmark_collected(category_name, renumbered_element_ids):
    # unmark all renumbered elements
    with revit.Transaction("Unmark {}".format(category_name)):
        unmark_renamed_elements(revit.active_view, renumbered_element_ids)


def pick_and_renumber(rnopts, starting_index):
    """Main renumbering routine for elements of given category."""
    # all actions under one transaction
    with revit.TransactionGroup("Renumber {}".format(rnopts.name)):
        # make sure target elements are easily selectable
        with EasilySelectableElements(revit.active_view, rnopts.bicat):
            index = starting_index
            # collect existing elements number:id data
            existing_elements_data = get_elements_dict(rnopts.bicat)
            # list to collect renumbered elements
            renumbered_element_ids = []
            # ask user to pick elements and renumber them
            for picked_element in revit.get_picked_elements_by_category(
                    rnopts.bicat,
                    message="Select {} in order".format(rnopts.name.lower())):
                # need nested transactions to push revit to update view
                # on each renumber task
                with revit.Transaction("Renumber {}".format(rnopts.name)):
                    # actual renumber task
                    renumber_element(picked_element,
                                     index, existing_elements_data)
                    # record the renumbered element
                    renumbered_element_ids.append(picked_element.Id)
                index = increment(index)
            # unmark all renumbered elements
            _unmark_collected(rnopts.name, renumbered_element_ids)


# ensure active view is a model view
if forms.check_modelview(revit.active_view):
    # prepare options
    renumber_options = [
        RNOpts(cat=BIC.OST_PipeAccessory),
        RNOpts(cat=BIC.OST_FabricationPipework),
        RNOpts(cat=BIC.OST_FabricationDuctwork),
        RNOpts(cat=BIC.OST_StructuralFraming)
        ]
    # add areas if active view is an Area Plan
    if revit.active_view.ViewType == DB.ViewType.AreaPlan:
        renumber_options.insert(1, RNOpts(cat=BIC.OST_Areas))

    options_dict = OrderedDict()
    for renumber_option in renumber_options:
        options_dict[renumber_option.name] = renumber_option
    selected_option_name = \
        forms.CommandSwitchWindow.show(
            options_dict,
            message='Pick element type to renumber:',
            width=400
        )

    if selected_option_name:
        selected_option = options_dict[selected_option_name]
        if selected_option.by_bicat:
            # if renumber doors by room
            if selected_option.bicat == BIC.OST_Doors \
                    and selected_option.by_bicat == BIC.OST_Rooms:
                with forms.WarningBar(
                    title='Pick Pairs of Door and Room. ESCAPE to end.'):
                    door_by_room_renumber(selected_option)
        else:
            starting_number = ask_for_starting_number(selected_option.name)
            if starting_number:
                with forms.WarningBar(
                    title='Pick {} One by One. ESCAPE to end.'.format(
                        selected_option.name)):
                    pick_and_renumber(selected_option, starting_number)
