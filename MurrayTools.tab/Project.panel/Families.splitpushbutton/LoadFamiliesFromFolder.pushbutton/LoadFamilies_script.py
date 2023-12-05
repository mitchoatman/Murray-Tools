import os
from pyrevit.framework import clr
from pyrevit import forms, revit, DB, script
import re
import pathlib
import itertools

logger = script.get_logger()

class FamilyLoader:
    """
    Enables loading a family from an absolute path.

    Attributes
    ----------
    path : str
        Absolute path to family .rfa file
    name : str
        File name
    is_loaded : bool
        Checks if family name already exists in project

    Methods
    -------
    get_symbols()
        Loads family in a fake transaction to return all symbols
    load_selective()
        Loads the family and selected symbols
    load_all()
        Loads family and all its symbols

    Credit
    ------
    Based on Ehsan Iran-Nejads 'Load More Types'
    """
    def __init__(self, path):
        """
        Parameters
        ----------
        path : str
            Absolute path to family .rfa file
        """
        self.path = path
        self.name = os.path.basename(path).replace(".rfa", "")

    @property
    def is_loaded(self):
        """
        Checks if family name already exists in project

        Returns
        -------
        bool
            Flag indicating if family is already loaded
        """
        collector = DB.FilteredElementCollector(revit.doc).OfClass(DB.Family)
        condition = (x for x in collector if x.Name == self.name)
        return next(condition, None) is not None

    def get_symbols(self):
        """
        Loads family in a fake transaction to return all symbols.

        Returns
        -------
        set()
            Set of family symbols

        Remark
        ------
        Uses SmartSortableFamilySymbol for effective sorting
        """
        logger.debug('Fake loading family: {}'.format(self.name))
        symbol_set = set()
        with revit.ErrorSwallower():
            # DryTransaction will rollback all the changes
            with revit.DryTransaction('Fake load'):
                ret_ref = clr.Reference[DB.Family]()
                revit.doc.LoadFamily(self.path, ret_ref)
                loaded_fam = ret_ref.Value
                # Get the symbols
                for symbol_id in loaded_fam.GetFamilySymbolIds():
                    symbol = revit.doc.GetElement(symbol_id)
                    symbol_name = revit.query.get_name(symbol)
                    sortable_sym = SmartSortableFamilySymbol(symbol_name)
                    logger.debug('Importable Symbol: {}'.format(sortable_sym))
                    symbol_set.add(sortable_sym)
        return sorted(symbol_set)

    def load_selective(self):
        """ Loads the family and selected symbols. """
        symbols = self.get_symbols()

        # Dont prompt if only 1 symbol available
        if len(symbols) == 1:
            self.load_all()
            return

        # User input -> Select family symbols
        selected_symbols = forms.SelectFromList.show(
            symbols,
            title=self.name,
            button_name="Load type(s)",
            multiselect=True)
        if selected_symbols is None:
            logger.debug('No family symbols selected.')
            return
        logger.debug('Selected symbols are: {}'.format(selected_symbols))

        # Load family with selected symbols
        with revit.Transaction('Loaded {}'.format(self.name)):
            try:
                for symbol in selected_symbols:
                    logger.debug('Loading symbol: {}'.format(symbol))
                    revit.doc.LoadFamilySymbol(self.path, symbol.symbol_name)
                logger.debug('Successfully loaded all selected symbols')
            except Exception as load_err:
                logger.error(
                    'Error loading family symbol from {} | {}'
                    .format(self.path, load_err))
                raise load_err

    def load_all(self):
        """ Loads family and all its symbols. """
        with revit.Transaction('Loaded {}'.format(self.name)):
            try:
                revit.doc.LoadFamily(self.path)
                logger.debug(
                    'Successfully loaded family: {}'.format(self.name))
            except Exception as load_err:
                logger.error(
                    'Error loading family symbol from {} | {}'
                    .format(self.path, load_err))
                raise load_err


class SmartSortableFamilySymbol:
    """
    Enables smart sorting of family symbols.

    Attributes
    ----------
    symbol_name : str
        name of the family symbol

    Example
    -------
    symbol_set = set()
    for family_symbol in familiy_symbols:
        family_symbol_name = revit.query.get_name(family_symbol)
        sortable_sym = SmartSortableFamilySymbol(family_symbol_name)
        symbol_set.add(sortable_sym)
    sorted_symbols = sorted(symbol_set)

    Credit
    ------
    Copied from Ehsan Iran-Nejads SmartSortableFamilyType
    in 'Load More Types'.
    """
    def __init__(self, symbol_name):
        self.symbol_name = symbol_name
        self.sort_alphabetically = False
        self.number_list = [
            int(x)
            for x in re.findall(r'\d+', self.symbol_name)]
        if not self.number_list:
            self.sort_alphabetically = True

    def __str__(self):
        return self.symbol_name

    def __repr__(self):
        return '<SmartSortableFamilySymbol Name:{} Values:{} StringSort:{}>'\
               .format(self.symbol_name,
                       self.number_list,
                       self.sort_alphabetically)

    def __eq__(self, other):
        return self.symbol_name == other.symbol_name

    def __hash__(self):
        return hash(self.symbol_name)

    def __lt__(self, other):
        if self.sort_alphabetically or other.sort_alphabetically:
            return self.symbol_name < other.symbol_name
        else:
            return self.number_list < other.number_list



class FileFinder:
    """
    Handles the file search in a directory

    Attributes
    ----------
    directory : str
        Path of the target directory
    paths : set()
        Holds absolute paths of search result

    Methods
    -------
    search(str)
        Searches in the target directory for the given glob pattern.
        Adds absolute paths to self.paths.

    exclude_by_pattern(str)
        Filters self.paths by the given regex pattern.
    """
    def __init__(self, directory):
        """
        Parameters
        ----------
        directory : str
            Absolute path to target directory.
        """
        self.directory = directory
        self.paths = set()

    def search(self, pattern):
        """
        Searches in the target directory for the given glob pattern.
        Adds absolute paths to self.paths.

        Parameters
        ----------
        pattern : str
            Glob pattern
        """
        result = pathlib.Path(self.directory).rglob(pattern)
        for path in result:
            logger.debug('Found file: {}'.format(path))
            self.paths.add(str(path))
        if len(self.paths) == 0:
            logger.debug(
                'No {} files in "{}" found.'.format(pattern, self.directory))
            forms.alert(
                'No {} files in "{}" found.'.format(pattern, self.directory))
            script.exit()

    def exclude_by_pattern(self, pattern):
        """
        Filters self.paths by the given regex pattern.

        Parameters
        ----------
        pattern : str
            Regular expression pattern
        """
        self.paths = itertools.ifilterfalse(
            re.compile(pattern).match, self.paths)



# Get directory with families
directory = forms.pick_folder("Select parent folder of families")
logger.debug('Selected parent folder: {}'.format(directory))
if directory is None:
    logger.debug('No directory selected. Calling script.exit')
    script.exit()

# Find family files in directory
finder = FileFinder(directory)
finder.search('*.rfa')

# Excluding backup files
backup_pattern = r'^.*\.\d{4}\.rfa$'
finder.exclude_by_pattern(backup_pattern)
paths = finder.paths

# Dictionary to look up absolute paths by relative paths
path_dict = dict()
for path in paths:
    path_dict.update({os.path.relpath(path, directory): path})

# User input -> Select families from directory
family_select_options = sorted(
    path_dict.keys(),
    key=lambda x: (x.count(os.sep), x))  # Sort by nesting level
selected_families = forms.SelectFromList.show(
    family_select_options,
    title="Select Families",
    width=500,
    button_name="Load Families",
    multiselect=True)
if selected_families is None:
    logger.debug('No families selected. Calling script.exit()')
    script.exit()
logger.debug('Selected Families: {}'.format(selected_families))

# Dictionary to look up FamilyLoader method by selected option
family_loading_options = {
    "Load all types": "load_all",
    "Load types by selecting individually": "load_selective"}
selected_loading_option = forms.CommandSwitchWindow.show(
    family_loading_options.keys(),
    message='Select loading option:',)
if selected_loading_option is None:
    logger.debug('No loading option selected. Calling script.exit()')
    script.exit()

# User input -> Select loading option (load all, load certain symbols)
logger.debug('Selected loading option: {}'.format(selected_loading_option))
laoding_option = family_loading_options[selected_loading_option]


# Feedback on already loaded families
already_loaded = set()

# Loading selected families
max_value = len(selected_families)
with forms.ProgressBar(title='Loading Family {value} of {max_value}',
                       cancellable=True) as pb:
    for count, family_path in enumerate(selected_families, 1):
        if pb.cancelled:
            break
        pb.update_progress(count, max_value)

        family = FamilyLoader(path_dict[family_path])
        logger.debug('Loading family: {}'.format(family.name))
        loaded = family.is_loaded
        if loaded:
            logger.debug('Family is already loaded: {}'.format(family.path))
            already_loaded.add(family)
            continue
        getattr(family, laoding_option)()

# Feedback on already loaded families
if len(already_loaded) != 0:
    for family in sorted(already_loaded):
        print(family.path)
