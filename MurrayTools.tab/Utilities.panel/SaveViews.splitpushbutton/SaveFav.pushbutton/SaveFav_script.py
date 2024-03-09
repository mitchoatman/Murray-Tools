from pyrevit import revit, DB, forms
import os


uidoc = revit.uidoc
doc = uidoc.Document
file_name = doc.Title

open_views = [doc.GetElement(view.ViewId) for view in uidoc.GetOpenUIViews() if not doc.GetElement(view.ViewId).IsTemplate]

if not open_views:
    print('There are no open views.')
    sys.exit()

if len(open_views) > 10:
    forms.alert(msg='You have more than ten open views.',
                title='Warning',
                sub_msg='Opening this many open views at once may take some time. Do you still wish to save these settings?',
                ok=False,
                yes=True,
                no=True,
                exitscript=True)

view_list = [view.Id for view in open_views]

folder_name = "c:\\Temp"
filepath = os.path.join(folder_name, 'Ribbon_OpenViews.txt')

# write values to a text file for future retrieval
with open((filepath), 'w') as the_file:
    line1 = (str(file_name) + '\n')
    line2 = (str(view_list) + '\n')
    the_file.writelines([line1, line2])

