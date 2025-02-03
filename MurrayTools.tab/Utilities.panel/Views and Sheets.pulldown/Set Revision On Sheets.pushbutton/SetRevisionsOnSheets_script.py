from pyrevit import revit, DB
from pyrevit import forms

# define a filterfunc to filter out issued revisions
def filterfunc(rev):
    return rev.Issued == False

revisions = forms.select_revisions(button_name='Select Revision',
                                   multiple=True,
                                   filterfunc=filterfunc)

if revisions:
    sheets = forms.select_sheets(button_name='Set Revision',
                                 include_placeholder=True)
    if sheets:
        with revit.Transaction('Set Revision on Sheets'):
            updated_sheets = revit.update.update_sheet_revisions(revisions,
                                                                 sheets)
        if updated_sheets:
            print('SELECTED REVISION ADDED TO THESE SHEETS:')
            print('-' * 100)
            for s in updated_sheets:
                snum = s.Parameter[DB.BuiltInParameter.SHEET_NUMBER]\
                        .AsString().rjust(10)
                sname = s.Parameter[DB.BuiltInParameter.SHEET_NAME]\
                         .AsString().ljust(50)
                print('NUMBER: {0}   NAME:{1}'.format(snum, sname))
