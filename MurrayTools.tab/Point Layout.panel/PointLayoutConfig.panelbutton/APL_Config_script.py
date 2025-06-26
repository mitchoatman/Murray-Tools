import os
import clr
clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')

from System.Windows.Forms import Application, Form, Label, TextBox, Button, DialogResult, FormStartPosition
from System.Drawing import Point, Size, Font

# --- Revit Document References ---
app = __revit__.Application
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
active_view = doc.ActiveView

# --- Get and Store Annular Space ---
folder_name = "c:\\Temp"
filepath = os.path.join(folder_name, 'Ribbon_Duct-Wall-Sleeve.txt')

if not os.path.exists(folder_name):
    os.makedirs(folder_name)

if not os.path.exists(filepath):
    with open(filepath, 'w') as f:
        f.write('1')

with open(filepath, 'r') as f:
    try:
        AnnularSpace = float(f.read()) / 2
    except:
        AnnularSpace = 0.5  # fallback

# --- Windows Form Dialog ---
class SleeveForm(Form):
    def __init__(self, initial_value):
        self.Text = "Sleeve Configuration"
        self.Size = Size(300, 140)
        self.StartPosition = FormStartPosition.CenterScreen

        self.label = Label()
        self.label.Text = "Sleeve Annular Space (Inches):"
        self.label.Location = Point(10, 10)
        self.label.Size = Size(260, 20)
        self.Controls.Add(self.label)

        self.textbox = TextBox()
        self.textbox.Text = str(initial_value)
        self.textbox.Location = Point(10, 35)
        self.textbox.Size = Size(260, 20)
        self.Controls.Add(self.textbox)

        self.ok_button = Button()
        self.ok_button.Text = "OK"
        self.ok_button.Location = Point(100, 70)
        self.ok_button.Click += self.ok_clicked
        self.AcceptButton = self.ok_button
        self.Controls.Add(self.ok_button)

        self.result_value = None

    def ok_clicked(self, sender, args):
        try:
            self.result_value = float(self.textbox.Text)
            self.DialogResult = DialogResult.OK
            self.Close()
        except:
            self.textbox.Text = "0.5"

# --- Show Form and Store Updated Value ---
form = SleeveForm(AnnularSpace)
if form.ShowDialog() == DialogResult.OK:
    updated_annular = form.result_value
    with open(filepath, 'w') as f:
        f.write(str(updated_annular * 2))
