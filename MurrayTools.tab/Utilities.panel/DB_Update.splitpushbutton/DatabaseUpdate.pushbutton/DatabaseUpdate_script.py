from pyrevit import forms
import os

path, filename = os.path.split(__file__)
NewFilename = '\Murray.ico'

forms.toast(
    "The Database and Support files are being synced to your Hard Drive. You must reload the Fabrication database inside Revit to receive any database changes.",
    title="Database Update",
    appid="Murray Tools",
    icon= path+NewFilename,
    click="https://murraycompany.com",)
os.startfile (r"C:\\Egnyte\Shared\\Engineering\\031407MEP\\Sync-Database\\Update Database - notimer.lnk")