import os
from pyrevit import forms

forms.toast(
    "The Database and Support files are being synced to your Hard Drive. You must reload the Fabrication database inside Revit to receive any database changes.",
    title="Database Update",
    appid="Murray Tools",
    icon="C:\Egnyte\Shared\BIM\Murray CADetailing Dept\REVIT\MURRAY RIBBON\Murray.extension\Murray.ico",
    click="https://murraycompany.com",)
os.startfile (r"C:\\Egnyte\Shared\\Engineering\\031407MEP\\Sync-Database\\Update Database - notimer.lnk")