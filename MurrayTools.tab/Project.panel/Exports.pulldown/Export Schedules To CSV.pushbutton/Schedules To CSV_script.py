__title__ = 'Schedules\nTo CSV'
__doc__ = """Exports selected schedules to CSV files."""
#__highlight__ = 'new'

import os.path as op

from pyrevit import forms
from pyrevit import coreutils
from pyrevit import revit, DB
from pyrevit import script

logger = script.get_logger()
output = script.get_output()


#open_exported = True
#incl_headers = True
basefolder = forms.pick_folder()


if basefolder:
    logger.debug(basefolder)
    schedules_to_export = forms.select_schedules()

    if schedules_to_export:
        vseop = DB.ViewScheduleExportOptions()
        vseop.ColumnHeaders = coreutils.get_enum_value(DB.ExportColumnHeaders,"OneRow")

        # determine which separator to use
        csv_sp = ','

        if csv_sp:
            vseop.FieldDelimiter = csv_sp
            vseop.Title = False
            vseop.HeadersFootersBlanks = True

            for sched in schedules_to_export:
                fname = \
                    coreutils.cleanup_filename(revit.query.get_name(sched)) \
                    + '.csv'
                sched.Export(basefolder, fname, vseop)
                exported = op.join(basefolder, fname)
                revit.files.correct_text_encoding(exported)
#                if open_exported:
#                    coreutils.run_process('"%s"' % exported)