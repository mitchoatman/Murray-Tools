import os
from pyrevit import EXEC_PARAMS, script, forms
from pyrevit.loader import sessionmgr
from pyrevit.loader import sessioninfo
from SharedParam.Add_Parameters import Shared_Params

Shared_Params()

try:
    my_file = 'C:\Egnyte\Shared\BIM\Murray CADetailing Dept\REVIT\Murray Ribbon\Murray.extension\hooks\UpdateFPData_doc-updater.txt'
    base = os.path.splitext(my_file)[0]
    os.rename(my_file, base + '.py')

    if EXEC_PARAMS.executed_from_ui:
        logger = script.get_logger()
        results = script.get_results()
        # re-load pyrevit session.
        logger.info('Reloading....')
        sessionmgr.reload_pyrevit()
        results.newsession = sessioninfo.get_session_uuid()
except:
    my_file = 'C:\Egnyte\Shared\BIM\Murray CADetailing Dept\REVIT\Murray Ribbon\Murray.extension\hooks\UpdateFPData_doc-updater.py'
    base = os.path.splitext(my_file)[0]
    os.rename(my_file, base + '.txt')

    if EXEC_PARAMS.executed_from_ui:
        logger = script.get_logger()
        results = script.get_results()
        # re-load pyrevit session.
        logger.info('Reloading....')
        sessionmgr.reload_pyrevit()
        results.newsession = sessioninfo.get_session_uuid()
