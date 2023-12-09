from pyrevit import forms
import Autodesk
from Autodesk.Revit.DB import *
import System
import os

doc = __revit__.ActiveUIDocument.Document
file_path = doc.PathName
file_name = System.IO.Path.GetFileNameWithoutExtension(file_path)

favorite_views = forms.select_views(title='Select Favorite Views',
                                    button_name='Save Settings')

if not favorite_views:
    sys.exit()

if len(favorite_views) > 10:
    forms.alert(msg='You have selected more than ten views.',
                title='Warning',
                sub_msg='Opening this many favorited views at once may take some time. Do you still wish to save these settings?',
                ok=False,
                yes=True,
                no=True,
                exitscript=True)

ViewList = []

for view in favorite_views:
    ViewList.append(view.Id)

folder_name = "c:\\Temp"
filepath = os.path.join(folder_name, 'Ribbon_FavoriteViews.txt')

# write values to text file for future retrieval
with open((filepath), 'w') as the_file:
    line1 = (str(file_name) + '\n')
    line2 = (str(ViewList) + '\n')
    the_file.writelines([line1, line2])  