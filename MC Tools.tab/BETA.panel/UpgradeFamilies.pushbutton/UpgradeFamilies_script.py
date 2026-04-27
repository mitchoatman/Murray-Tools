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

        # Handles "newer version" dialog
        if dialog_id == "TaskDialog_Opening_File_From_Later_Version":
            args.OverrideResult(1)


dialog_handler = DialogSuppressor()
uiapp.DialogBoxShowing += dialog_handler.handler


# -----------------------------
# Helpers
# -----------------------------
def get_or_create_preview_3d_view(doc):
    """Return a valid 3D view for preview. Create one if needed."""
    settings = doc.GetDocumentPreviewSettings()

    # Try existing 3D views first
    collector = DB.FilteredElementCollector(doc).OfClass(DB.View3D)
    for v in collector:
        if v.IsTemplate:
            continue
        if settings.IsViewIdValidForPreview(v.Id):
            return v

    # If none found, create a new isometric 3D view
    view_type_id = None
    vft_collector = DB.FilteredElementCollector(doc).OfClass(DB.ViewFamilyType)
    for vft in vft_collector:
        if vft.ViewFamily == DB.ViewFamily.ThreeDimensional:
            view_type_id = vft.Id
            break

    if view_type_id is None:
        return None

    t = DB.Transaction(doc, "Create Preview 3D View")
    t.Start()
    new_view = DB.View3D.CreateIsometric(doc, view_type_id)

    try:
        new_view.Name = "{3D}_BatchPreview"
    except:
        pass

    t.Commit()

    if settings.IsViewIdValidForPreview(new_view.Id):
        return new_view

    return None


def set_preview_view(doc, view):
    """Set the document preview to the supplied view."""
    settings = doc.GetDocumentPreviewSettings()

    t = DB.Transaction(doc, "Set Family Preview View")
    t.Start()

    fail_options = t.GetFailureHandlingOptions()
    fail_options.SetFailuresPreprocessor(SilentFailuresPreprocessor())
    t.SetFailureHandlingOptions(fail_options)

    settings.PreviewViewId = view.Id
    t.Commit()


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
preview_updated = 0


# -----------------------------
# Main Processing
# -----------------------------
app = revit.doc.Application

with forms.ProgressBar(title="Upgrading Families", max_value=len(rfa_files), cancellable=True) as pb:
    for i, file_path in enumerate(rfa_files):

        if pb.cancelled:
            break

        doc = None

        try:
            model_path = DB.ModelPathUtils.ConvertUserVisiblePathToModelPath(file_path)

            open_opts = DB.OpenOptions()
            open_opts.Audit = True

            doc = app.OpenDocumentFile(model_path, open_opts)

            preview_view = get_or_create_preview_3d_view(doc)

            save_opts = DB.SaveOptions()

            if preview_view:
                set_preview_view(doc, preview_view)
                save_opts.PreviewViewId = preview_view.Id
                preview_updated += 1

            doc.Save(save_opts)
            doc.Close(False)

            processed_count += 1

        except Exception as e:
            msg = str(e).lower()

            try:
                if doc and doc.IsModifiable:
                    pass
            except:
                pass

            try:
                if doc:
                    doc.Close(False)
            except:
                pass

            if "later version" in msg:
                skipped_files.append(file_path)
            else:
                failed_files.append((file_path, str(e)))

        pb.update_progress(i + 1, len(rfa_files))


# -----------------------------
# CLEANUP
# -----------------------------
uiapp.DialogBoxShowing -= dialog_handler.handler


# -----------------------------
# Final Report
# -----------------------------
report = []
report.append("=== BATCH UPGRADE COMPLETE ===\n")
report.append("Total Files: {}".format(len(rfa_files)))
report.append("Processed: {}".format(processed_count))
report.append("Preview Updated: {}".format(preview_updated))
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