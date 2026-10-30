# -*- coding: utf-8 -*-

"""
Copy View Templates - Phase 3 Production
========================================

Copies selected View Templates from another OPEN Revit project
into the ACTIVE Revit project.

TARGET ENVIRONMENT
------------------
Revit 2024 / 2025
pyRevit 6.x
IronPython 2.7

WORKFLOW
--------
1. Active document = DESTINATION.
2. Select another open Revit project = SOURCE.
3. Select one or more View Templates.
4. New templates are copied.
5. Existing same-name templates can be:
       - Skipped
       - Replaced
       - Cancelled
6. When replacing:
       - Views using old template are captured.
       - Assignments are temporarily cleared.
       - Old template is deleted.
       - Source template is copied.
       - Destination views are reassigned.
7. Entire operation occurs in ONE transaction.
   Any failure rolls back all changes.

PHASE 3
-------
Production polish:
    - Clear source/destination identification
    - Searchable sorted template list
    - Conflict impact reporting
    - Number of affected views shown before replacement
    - Clean completion report
    - Revit version validation
    - Stable Phase 2 copy/replace engine retained
"""

from pyrevit import DB, forms, script

from System.Collections.Generic import List


# =============================================================================
# CONSTANTS
# =============================================================================

TITLE = "Copy View Templates"

SUPPORTED_REVIT_VERSIONS = [
    2024,
    2025
]


# =============================================================================
# REVIT CONTEXT
# =============================================================================

uidoc = __revit__.ActiveUIDocument
doc = uidoc.Document
app = doc.Application

output = script.get_output()


# =============================================================================
# WRAPPER CLASSES
# =============================================================================

class DocumentOption(object):
    """Display wrapper for an open Revit project."""

    def __init__(self, revit_doc):
        self.doc = revit_doc

    @property
    def name(self):
        return self.doc.Title

    @property
    def description(self):
        try:
            path = self.doc.PathName

            if path:
                return path
        except Exception:
            pass

        return "Unsaved project"

    def __str__(self):
        return self.doc.Title


class TemplateOption(object):
    """Display wrapper for source View Templates."""

    def __init__(self, view):
        self.view = view

    @property
    def name(self):
        return self.view.Name

    def __str__(self):
        return self.view.Name


# =============================================================================
# DUPLICATE TYPE HANDLER
# =============================================================================

class DuplicateTypeNamesHandler(DB.IDuplicateTypeNamesHandler):
    """
    Tell Revit to use an existing destination type whenever a copied
    dependency encounters a duplicate type name.

    This prevents Revit's duplicate-type modal warning.
    """

    def OnDuplicateTypeNamesFound(self, args):

        return DB.DuplicateTypeAction.UseDestinationTypes


# =============================================================================
# VERSION CHECK
# =============================================================================

def check_revit_version():
    """
    Phase 3 has been written/tested for Revit 2024/2025.

    Do not hard-fail future versions; warn instead so testing can
    continue when moving the tool forward.
    """

    try:
        version = int(app.VersionNumber)
    except Exception:
        return

    if version not in SUPPORTED_REVIT_VERSIONS:

        result = forms.alert(
            "This tool was built and validated for:\n\n"
            "Revit 2024\n"
            "Revit 2025\n\n"
            "Current Revit version:\n"
            "{}\n\n"
            "The required API methods may still work, but this "
            "version has not been validated.\n\n"
            "Continue?".format(version),
            title=TITLE,
            options=[
                "Continue",
                "Cancel"
            ],
            warn_icon=True
        )

        if result != "Continue":
            script.exit()


# =============================================================================
# DOCUMENT HELPERS
# =============================================================================

def get_source_documents():
    """
    Get valid source documents currently open in this Revit session.

    Excludes:
        - Active destination document
        - Family documents
        - Linked documents
    """

    documents = []

    for open_doc in app.Documents:

        try:

            if open_doc.Equals(doc):
                continue

            if open_doc.IsFamilyDocument:
                continue

            try:
                if open_doc.IsLinked:
                    continue
            except Exception:
                pass

            documents.append(open_doc)

        except Exception:
            continue

    documents.sort(
        key=lambda x: x.Title.lower()
    )

    return documents


# =============================================================================
# VIEW TEMPLATE HELPERS
# =============================================================================

def get_view_templates(revit_doc):
    """
    Get all View Templates in a document.

    Collector is limited to DB.View and therefore does NOT scan
    the production model's physical elements.
    """

    templates = []

    collector = (
        DB.FilteredElementCollector(revit_doc)
        .OfClass(DB.View)
    )

    for view in collector:

        try:

            if view.IsTemplate:
                templates.append(view)

        except Exception:
            continue

    templates.sort(
        key=lambda x: x.Name.lower()
    )

    return templates


def get_template_dictionary(revit_doc):
    """
    Return View Templates indexed by name.

    Example:

        {
            "FP - Plumbing": <View>,
            "FP - Hangers": <View>
        }
    """

    result = {}

    for template in get_view_templates(revit_doc):

        result[template.Name] = template

    return result


def get_views_using_template(revit_doc, template_id):
    """
    Return all non-template views currently assigned to template_id.

    Only DB.View objects are collected.
    """

    result = []

    collector = (
        DB.FilteredElementCollector(revit_doc)
        .OfClass(DB.View)
    )

    for view in collector:

        try:

            if view.IsTemplate:
                continue

            if view.ViewTemplateId == template_id:
                result.append(view)

        except Exception:
            continue

    return result


# =============================================================================
# ELEMENT ID HELPERS
# =============================================================================

def make_element_id_list(elements):
    """
    Convert Python list of Revit elements into the generic collection
    required by ElementTransformUtils.CopyElements.
    """

    element_ids = List[DB.ElementId]()

    for element in elements:
        element_ids.Add(element.Id)

    return element_ids


# =============================================================================
# UTILITY
# =============================================================================

def safe_view_name(view):
    """Return a useful view name without allowing an error to stop reporting."""

    try:
        return view.Name
    except Exception:
        return "<Unknown View>"


# =============================================================================
# START
# =============================================================================

check_revit_version()


# =============================================================================
# VERIFY DESTINATION
# =============================================================================

if doc.IsFamilyDocument:

    forms.alert(
        "The active document is a Revit family.\n\n"
        "Open or activate the PROJECT that should receive the "
        "View Templates and run the tool again.",
        title=TITLE,
        warn_icon=True
    )

    script.exit()


# =============================================================================
# GET SOURCE PROJECTS
# =============================================================================

source_docs = get_source_documents()

if not source_docs:

    forms.alert(
        "No available SOURCE projects were found.\n\n"
        "DESTINATION:\n"
        "{}\n\n"
        "Open the Revit project containing the View Templates, "
        "then run this tool again.".format(
            doc.Title
        ),
        title=TITLE,
        warn_icon=True
    )

    script.exit()


# =============================================================================
# SELECT SOURCE
# =============================================================================

document_options = []

for source_doc_item in source_docs:

    document_options.append(
        DocumentOption(source_doc_item)
    )


selected_document = forms.SelectFromList.show(
    document_options,
    title="SOURCE Project  |  DESTINATION: {}".format(
        doc.Title
    ),
    button_name="Select Source",
    multiselect=False,
    width=700,
    height=500
)

if not selected_document:
    script.exit()


source_doc = selected_document.doc


# =============================================================================
# GET SOURCE VIEW TEMPLATES
# =============================================================================

source_templates = get_view_templates(
    source_doc
)

if not source_templates:

    forms.alert(
        "No View Templates were found in the selected SOURCE.\n\n"
        "SOURCE:\n"
        "{}\n\n"
        "DESTINATION:\n"
        "{}".format(
            source_doc.Title,
            doc.Title
        ),
        title=TITLE,
        warn_icon=True
    )

    script.exit()


# =============================================================================
# SELECT VIEW TEMPLATES
# =============================================================================

template_options = []

for template in source_templates:

    template_options.append(
        TemplateOption(template)
    )


selected_options = forms.SelectFromList.show(
    template_options,
    title="Copy View Templates  |  SOURCE: {}".format(
        source_doc.Title
    ),
    button_name="Copy Selected",
    multiselect=True,
    width=750,
    height=750
)

if not selected_options:
    script.exit()


selected_templates = []

for option in selected_options:

    selected_templates.append(
        option.view
    )


# =============================================================================
# DESTINATION TEMPLATE LOOKUP
# =============================================================================

destination_templates = get_template_dictionary(
    doc
)


# =============================================================================
# CLASSIFY SELECTED TEMPLATES
# =============================================================================

new_templates = []
conflicting_templates = []

for source_template in selected_templates:

    if source_template.Name in destination_templates:

        conflicting_templates.append(
            source_template
        )

    else:

        new_templates.append(
            source_template
        )


# =============================================================================
# PREFLIGHT CONFLICT INFORMATION
# =============================================================================

replacement_data = {}

total_affected_views = 0


for source_template in conflicting_templates:

    template_name = source_template.Name

    old_template = destination_templates[
        template_name
    ]

    affected_views = get_views_using_template(
        doc,
        old_template.Id
    )

    replacement_data[template_name] = {
        "source_template": source_template,
        "old_template": old_template,
        "affected_views": affected_views
    }

    total_affected_views += len(
        affected_views
    )


# =============================================================================
# CONFLICT DECISION
# =============================================================================

replace_existing = False
skip_existing = False


if conflicting_templates:

    conflict_names = []

    for template in conflicting_templates:

        template_name = template.Name

        affected_count = len(
            replacement_data[
                template_name
            ]["affected_views"]
        )

        if affected_count == 1:

            conflict_names.append(
                "• {}  [1 view using template]".format(
                    template_name
                )
            )

        else:

            conflict_names.append(
                "• {}  [{} views using template]".format(
                    template_name,
                    affected_count
                )
            )


    conflict_names.sort()


    conflict_message = (
        "SOURCE\n"
        "{}\n\n"
        "DESTINATION\n"
        "{}\n\n"
        "{} selected template(s) already exist in the destination.\n\n"
        "{}\n\n"
        "TOTAL DESTINATION VIEWS AFFECTED IF REPLACED: {}\n\n"
        "REPLACE EXISTING\n"
        "Deletes the destination template, copies the SOURCE version, "
        "and reconnects all affected destination views.\n\n"
        "SKIP EXISTING\n"
        "Leaves existing destination templates untouched and copies "
        "only templates that do not already exist."
    ).format(
        source_doc.Title,
        doc.Title,
        len(conflicting_templates),
        "\n".join(conflict_names),
        total_affected_views
    )


    choice = forms.alert(
        conflict_message,
        title="View Template Conflicts",
        options=[
            "Replace Existing",
            "Skip Existing",
            "Cancel"
        ],
        warn_icon=True
    )


    if choice == "Replace Existing":

        replace_existing = True

    elif choice == "Skip Existing":

        skip_existing = True

    else:

        script.exit()


# =============================================================================
# BUILD FINAL COPY LIST
# =============================================================================

templates_to_copy = []


# -----------------------------------------------------------------------------
# Always copy templates that don't exist yet
# -----------------------------------------------------------------------------

for template in new_templates:

    templates_to_copy.append(
        template
    )


# -----------------------------------------------------------------------------
# Add conflicts only when Replace Existing was selected
# -----------------------------------------------------------------------------

if replace_existing:

    for template in conflicting_templates:

        templates_to_copy.append(
            template
        )


# =============================================================================
# NOTHING TO DO
# =============================================================================

if not templates_to_copy:

    forms.alert(
        "No View Templates were copied.\n\n"
        "{} existing template(s) were skipped.".format(
            len(conflicting_templates)
        ),
        title=TITLE
    )

    script.exit()


# =============================================================================
# COPY OPTIONS
# =============================================================================

copy_options = DB.CopyPasteOptions()

duplicate_handler = DuplicateTypeNamesHandler()

copy_options.SetDuplicateTypeNamesHandler(
    duplicate_handler
)


# =============================================================================
# TRANSACTION
# =============================================================================

transaction = DB.Transaction(
    doc,
    "Copy View Templates"
)


# =============================================================================
# RESULT TRACKING
# =============================================================================

copied_new_names = []
replaced_names = []
skipped_names = []

reassigned_view_count = 0

reassigned_views_by_template = {}


# =============================================================================
# EXECUTE
# =============================================================================

try:

    # =========================================================================
    # START ONE ATOMIC TRANSACTION
    # =========================================================================

    transaction.Start()


    # =========================================================================
    # STEP 1
    # TEMPORARILY REMOVE TEMPLATE ASSIGNMENTS
    # =========================================================================

    if replace_existing:

        for template_name in replacement_data:

            data = replacement_data[
                template_name
            ]

            affected_views = data[
                "affected_views"
            ]

            for view in affected_views:

                view.ViewTemplateId = (
                    DB.ElementId.InvalidElementId
                )


    # =========================================================================
    # STEP 2
    # DELETE DESTINATION TEMPLATES BEING REPLACED
    # =========================================================================

    if replace_existing:

        for template_name in replacement_data:

            data = replacement_data[
                template_name
            ]

            old_template = data[
                "old_template"
            ]

            doc.Delete(
                old_template.Id
            )


        # Make deletion state available before copying
        doc.Regenerate()


    # =========================================================================
    # STEP 3
    # COPY SOURCE VIEW TEMPLATES
    # =========================================================================

    source_ids = make_element_id_list(
        templates_to_copy
    )


    copied_ids = (
        DB.ElementTransformUtils.CopyElements(
            source_doc,
            source_ids,
            doc,
            DB.Transform.Identity,
            copy_options
        )
    )


    # =========================================================================
    # STEP 4
    # REGENERATE TO MAKE NEW TEMPLATES AVAILABLE
    # =========================================================================

    doc.Regenerate()


    # =========================================================================
    # STEP 5
    # REBUILD DESTINATION TEMPLATE LOOKUP
    # =========================================================================

    destination_after_copy = (
        get_template_dictionary(doc)
    )


    # =========================================================================
    # STEP 6
    # VALIDATE EVERY REQUESTED TEMPLATE EXISTS
    # =========================================================================

    for source_template in templates_to_copy:

        template_name = (
            source_template.Name
        )


        if template_name not in destination_after_copy:

            raise Exception(
                "Copied View Template could not be found "
                "after the copy operation.\n\n"
                "Template:\n"
                "{}".format(
                    template_name
                )
            )


    # =========================================================================
    # STEP 7
    # TRACK NEW TEMPLATES
    # =========================================================================

    for source_template in new_templates:

        template_name = source_template.Name

        if template_name in destination_after_copy:

            copied_new_names.append(
                template_name
            )


    # =========================================================================
    # STEP 8
    # REASSIGN DESTINATION VIEWS
    # =========================================================================

    if replace_existing:

        for template_name in replacement_data:

            if template_name not in destination_after_copy:

                raise Exception(
                    "Replacement View Template could not be found.\n\n"
                    "Template:\n"
                    "{}".format(
                        template_name
                    )
                )


            new_template = (
                destination_after_copy[
                    template_name
                ]
            )


            affected_views = (
                replacement_data[
                    template_name
                ]["affected_views"]
            )


            reassigned_views_by_template[
                template_name
            ] = []


            for view in affected_views:

                try:

                    view.ViewTemplateId = (
                        new_template.Id
                    )

                    reassigned_view_count += 1

                    reassigned_views_by_template[
                        template_name
                    ].append(
                        safe_view_name(view)
                    )


                except Exception as assignment_error:

                    raise Exception(
                        "A destination view could not be reassigned "
                        "to its replacement View Template.\n\n"
                        "VIEW\n"
                        "{}\n\n"
                        "TEMPLATE\n"
                        "{}\n\n"
                        "ERROR\n"
                        "{}".format(
                            safe_view_name(view),
                            template_name,
                            str(assignment_error)
                        )
                    )


            replaced_names.append(
                template_name
            )


    # =========================================================================
    # STEP 9
    # COMMIT
    # =========================================================================

    transaction.Commit()


# =============================================================================
# ERROR / ROLLBACK
# =============================================================================

except Exception as ex:

    try:

        if (
            transaction.GetStatus()
            == DB.TransactionStatus.Started
        ):

            transaction.RollBack()

    except Exception:
        pass


    forms.alert(
        "VIEW TEMPLATE TRANSFER FAILED\n\n"
        "No changes were kept.\n\n"
        "SOURCE\n"
        "{}\n\n"
        "DESTINATION\n"
        "{}\n\n"
        "ERROR\n"
        "{}".format(
            source_doc.Title,
            doc.Title,
            str(ex)
        ),
        title=TITLE,
        warn_icon=True
    )

    script.exit()


# =============================================================================
# SKIPPED TEMPLATE TRACKING
# =============================================================================

if skip_existing:

    for template in conflicting_templates:

        skipped_names.append(
            template.Name
        )


# =============================================================================
# CLEAN RESULT LISTS
# =============================================================================

copied_new_names = sorted(
    set(copied_new_names)
)

replaced_names = sorted(
    set(replaced_names)
)

skipped_names = sorted(
    set(skipped_names)
)


# =============================================================================
# FINAL REPORT
# =============================================================================

report_lines = []


report_lines.append(
    "SOURCE"
)

report_lines.append(
    source_doc.Title
)

report_lines.append("")

report_lines.append(
    "DESTINATION"
)

report_lines.append(
    doc.Title
)

report_lines.append("")


# -----------------------------------------------------------------------------
# NEW
# -----------------------------------------------------------------------------

report_lines.append(
    "NEW TEMPLATES COPIED: {}".format(
        len(copied_new_names)
    )
)

for name in copied_new_names:

    report_lines.append(
        "  • {}".format(name)
    )

report_lines.append("")


# -----------------------------------------------------------------------------
# REPLACED
# -----------------------------------------------------------------------------

report_lines.append(
    "EXISTING TEMPLATES REPLACED: {}".format(
        len(replaced_names)
    )
)

for name in replaced_names:

    affected_count = len(
        reassigned_views_by_template.get(
            name,
            []
        )
    )

    report_lines.append(
        "  • {}  [{} view(s) reassigned]".format(
            name,
            affected_count
        )
    )

report_lines.append("")


# -----------------------------------------------------------------------------
# SKIPPED
# -----------------------------------------------------------------------------

report_lines.append(
    "EXISTING TEMPLATES SKIPPED: {}".format(
        len(skipped_names)
    )
)

for name in skipped_names:

    report_lines.append(
        "  • {}".format(name)
    )

report_lines.append("")


# -----------------------------------------------------------------------------
# TOTAL VIEWS
# -----------------------------------------------------------------------------

report_lines.append(
    "DESTINATION VIEWS REASSIGNED: {}".format(
        reassigned_view_count
    )
)

report_lines.append("")

report_lines.append(
    "Transfer completed successfully."
)


forms.alert(
    "\n".join(report_lines),
    title=TITLE
)