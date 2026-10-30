# -*- coding: utf-8 -*-
"""Murray Hub launcher. Opens the modeless hub window (WebView2)."""
import os

# Keep the script engine alive so the window and ExternalEvent keep working
# after this script returns. [CONFIRM flag name in pyRevit 7.0 docs]
__persistentengine__ = True

from mchub import webhost

ui_dir = os.path.join(os.path.dirname(__file__), 'ui')
webhost.open_hub(__revit__, ui_dir)
