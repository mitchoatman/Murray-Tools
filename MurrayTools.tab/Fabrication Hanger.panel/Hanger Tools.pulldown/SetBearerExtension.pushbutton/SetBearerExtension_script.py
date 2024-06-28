
#Imports
from Autodesk.Revit import DB
from Autodesk.Revit.DB import FilteredElementCollector, Transaction, BuiltInCategory, FamilySymbol
from Autodesk.Revit.UI.Selection import *
import os

# -*- coding: utf-8 -*-

uidoc  = __revit__.ActiveUIDocument
doc    = __revit__.ActiveUIDocument.Document #type:Document

# .NET Imports
import clr
clr.AddReference('System')
clr.AddReference('System.Drawing')
clr.AddReference('System.Windows.Forms')
import System
from System.Windows.Forms import *
from System.Drawing import Point, Size, Font, FontStyle


class TXT_Form(Form):
    def __init__(self):
        self.Text          = 'Bearer Extension'
        self.Size          = Size(250,150)
        self.StartPosition = FormStartPosition.CenterScreen

        self.TopMost     = True  # Keeps the form on top of all other windows
        self.ShowIcon    = False  # Removes the icon from the title bar
        self.MaximizeBox = False
        self.MinimizeBox = False
        self.FormBorderStyle = FormBorderStyle.FixedDialog  # Disallows resizing

        # Initialize valuenum to None
        self.valuenum = None

        #Design of dialog

        #Label for TextBox
        self.label_textbox           = Label()
        self.label_textbox.Text      = 'Extension:'
        self.label_textbox.ForeColor = System.Drawing.Color.Black
        self.label_textbox.Font      = Font("Arial", 12, FontStyle.Bold)
        self.label_textbox.Location  = Point(20,20)
        self.label_textbox.Size      = Size(90,40)
        self.Controls.Add(self.label_textbox)

        #TextBox
        self.textBox          = TextBox()
        self.textBox.Text     = '1'
        self.textBox.Location = Point(110, 20)
        self.textBox.Size     = Size(100, 40)
        self.Controls.Add(self.textBox)

        #Button
        self.button          = Button()
        self.button.Text     = 'Modify Extension'
        self.button.Location = Point(68, 60)
        self.button.Size     = Size(100, 30)
        self.Controls.Add(self.button)

        self.button.Click      += self.on_click
        # self.button.MouseEnter += self.btn_hover

    def on_click(self, sender, event):
        try:
            self.value = self.textBox.Text
            self.valuenum = float(self.value) / 12
        except ValueError:
            MessageBox.Show("Please enter a valid number.")
        self.Close()

class MySelectionFilter(ISelectionFilter):
    def __init__(self):
        pass
    def AllowElement(self, element):
        if element.Category.Name == 'MEP Fabrication Hangers':
            return True
        else:
            return False
    def AllowReference(self, element):
        return False
selection_filter = MySelectionFilter()
Hanger = uidoc.Selection.PickElementsByRectangle(selection_filter)

if len(Hanger) > 0:
    #Show the Form
    form = TXT_Form()
    # form.Show()
    Application.Run(form)

    if form.valuenum is not None:
        t = Transaction(doc, 'Modify Bearer Extension')
        t.Start()

        for x in range(2):
            for e in Hanger:
                STName = e.GetRodInfo().RodCount
                STName1 = e.GetRodInfo()
                if STName > 1:
                    for n in range(STName):
                        STName1.SetBearerExtension(n, form.valuenum)

        t.Commit()
else:
    forms.alert('At least one fabrication hanger must be selected.')

