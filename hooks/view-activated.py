# -*- coding: utf-8 -*-
"""Workset Auto-Switch Hook - Only active when toggle is enabled."""

import os
import clr
clr.AddReference('System')

from pyrevit import revit, EXEC_PARAMS
from Autodesk.Revit.DB import FilteredWorksetCollector, WorksetKind

# ── Guard: bail out immediately if hook is disabled ───────────────────────────
FLAG_FILE = r'C:\temp\Ribbon_Workset-hook-status.txt'

def hook_is_enabled():
    try:
        if os.path.exists(FLAG_FILE):
            with open(FLAG_FILE, 'r') as f:
                state = f.read().strip().lower()
                return state == 'true'
    except Exception:
        pass
    return True  # Default enabled

# ── Main Logic ───────────────────────────────────────────────────────────────
if not hook_is_enabled():
    # print("Workset Hook: Disabled via toggle")
    pass
else:
    try:
        args = EXEC_PARAMS.event_args
        doc = revit.doc
       
        active_view = args.CurrentActiveView
        # print("Workset Hook: Event fired - ViewType = {}".format(
            # str(getattr(active_view, 'ViewType', None))
        # ))
        
        if str(getattr(active_view, 'ViewType', None)) == 'FloorPlan':
            # print("Workset Hook: FloorPlan detected - attempting workset switch")
           
            try:
                WorksetNames = []
                WorksetIds = []
                
                # Use the imported WorksetKind
                collector = FilteredWorksetCollector(doc)
                AllWorksets = collector.OfKind(WorksetKind.UserWorkset).ToWorksets()
               
                for ws in AllWorksets:
                    WorksetNames.append(ws.Name)
                    WorksetIds.append(ws.Id)
                
                level = active_view.GenLevel
                if level is None:
                    pass
                    # print("Workset Hook: No level found on view")
                else:
                    pass
                    # print("Workset Hook: Looking for workset named '{}'".format(level.Name))
                   
                    try:
                        index = WorksetNames.index(level.Name)
                        doc.GetWorksetTable().SetActiveWorksetId(WorksetIds[index])
                        # print("Workset Hook: SUCCESS - Set active workset to {}".format(level.Name))
                    except ValueError:
                        pass
                        # print("Workset Hook: No workset found with name '{}'".format(level.Name))
                    except Exception as e:
                        pass
                        # print("Workset Hook: Error setting workset: {}".format(str(e)))
            except Exception as e:
                pass
                # print("Workset Hook: Error collecting worksets: {}".format(str(e)))
    except Exception as e:
        pass
        # print("Workset Hook: General error: {}".format(str(e)))