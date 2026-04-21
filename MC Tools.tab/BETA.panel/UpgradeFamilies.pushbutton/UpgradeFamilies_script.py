# -*- coding: utf-8 -*-
import os
import clr

from pyrevit import forms, revit, DB

# -----------------------------
# Failure Suppression
# -----------------------------
class SilentFailuresPreprocessor(DB.IFailuresPreprocessor):
    def PreprocessFailures(self, failuresAccessor):
        for f in failuresAccessor.GetFailureMessages():
            failuresAccessor.DeleteWarning(f)
        return DB.FailureProcessingResult.Continue


# -----------------------------
# Dialog Suppression (PREVENT FREEZE)
# -----------------------------
uiapp = revit.uidoc.Application

class DialogSuppressor(object):
    def handler(self, sender, args):
        dialog_id = args.DialogId

        # Handles "newer version" dialog (prevents freeze)
        if dialog_id == "TaskDialog_Opening_File_From_Later_Version":
            args.OverrideResult(1)


dialog_handler = DialogSuppressor()
uiapp.DialogBoxShowing += dialog_handler.handler


# -----------------------------
# Pick Folder
# -----------------------------
source_folder = forms.pick_folder(title="Select Library Folder to Upgrade")

if not source_folder:
    uiapp.DialogBoxShowing -= dialog_handler.handler
    forms.alert("No folder selected.", exitscript=True)


# -----------------------------
# Collect RFA files
# -----------------------------
rfa_files = []
for root, dirs, files in os.walk(source_folder):
    for file in files:
        if file.lower().endswith(".rfa"):
            rfa_files.append(os.path.join(root, file))

if not rfa_files:
    uiapp.DialogBoxShowing -= dialog_handler.handler
    forms.alert("No RFA files found.", exitscript=True)


# -----------------------------
# Results tracking
# -----------------------------
processed_count = 0
failed_files = []
skipped_files = []


# -----------------------------
# Main Processing
# -----------------------------
app = revit.doc.Application

with forms.ProgressBar(title="Upgrading Families", max_value=len(rfa_files), cancellable=True) as pb:
    for i, file_path in enumerate(rfa_files):

        if pb.cancelled:
            break

        try:
            # Convert to ModelPath (required)
            model_path = DB.ModelPathUtils.ConvertUserVisiblePathToModelPath(file_path)

            # Open options
            open_opts = DB.OpenOptions()
            open_opts.Audit = True

            # Open document
            doc = app.OpenDocumentFile(model_path, open_opts)

            # Suppress warnings
            t = DB.Transaction(doc, "Suppress Warnings")
            t.Start()

            fail_options = t.GetFailureHandlingOptions()
            fail_options.SetFailuresPreprocessor(SilentFailuresPreprocessor())
            t.SetFailureHandlingOptions(fail_options)

            t.Commit()

            # Save + Close
            doc.Save()
            doc.Close(False)

            processed_count += 1

        except Exception as e:
            msg = str(e).lower()

            if "later version" in msg:
                skipped_files.append(file_path)
            else:
                failed_files.append((file_path, str(e)))

        pb.update_progress(i + 1, len(rfa_files))


# -----------------------------
# CLEANUP (IMPORTANT)
# -----------------------------
uiapp.DialogBoxShowing -= dialog_handler.handler


# -----------------------------
# Final Report
# -----------------------------
report = []
report.append("=== BATCH UPGRADE COMPLETE ===\n")
report.append("Total Files: {}".format(len(rfa_files)))
report.append("Processed: {}".format(processed_count))
report.append("Skipped (Newer Version): {}".format(len(skipped_files)))
report.append("Failed: {}\n".format(len(failed_files)))

if skipped_files:
    report.append("=== SKIPPED FILES ===")
    report.extend(skipped_files)

if failed_files:
    report.append("\n=== FAILED FILES ===")
    for f, err in failed_files:
        report.append("{}\n  -> {}".format(f, err))

report_text = "\n".join(report)

forms.alert(report_text, title="Upgrade Report")