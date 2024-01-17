from pyrevit import forms
import os

path, filename = os.path.split(__file__)
NewFilename = '\Murray.ico'

forms.toast(
    "Congrats! Your Revit experience has been enhanced.",
    title="Murray Ribbon Loaded",
    appid="Murray Tools",
    icon= path+NewFilename,
    click="https://murraycompany.com",)