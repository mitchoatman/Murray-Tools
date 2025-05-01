import Autodesk
from pyrevit import revit
from Autodesk.Revit.DB import Transaction, FabricationConfiguration
import os
from Parameters.Add_SharedParameters import Shared_Params
from Parameters.Get_Set_Params import set_parameter_by_name
Shared_Params()

# .NET Imports
import clr
clr.AddReference('System')
clr.AddReference('System.Drawing')
clr.AddReference('System.Windows.Forms')
import System
from System.Windows.Forms import *
from System.Drawing import Point, Size, Font, FontStyle

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
Config = FabricationConfiguration.GetFabricationConfiguration(doc)

folder_name = "c:\\Temp"
filepath = os.path.join(folder_name, 'Ribbon_StratusAssembly.txt')

if not os.path.exists(folder_name):
    os.makedirs(folder_name)
if not os.path.exists(filepath):
    with open(filepath, 'w') as the_file:
        line1 = 'L1-A1-CW-01\n'
        line2 = 'L1-A1-HGR-MAP\n'
        the_file.writelines([line1, line2])

with open(filepath, 'r') as file:
    lines = file.readlines()
    lines = [line.rstrip() for line in lines]

class TXT_Form(Form):
    def __init__(self):
        self.Text = 'Spool Data'
        self.Size = Size(275, 200)
        self.StartPosition = FormStartPosition.CenterScreen
        self.TopMost = True
        self.ShowIcon = False
        self.MaximizeBox = False
        self.MinimizeBox = False
        self.FormBorderStyle = FormBorderStyle.FixedDialog

        self.value = None

        self.label_textbox = Label()
        self.label_textbox.Text = 'Spool Name:'
        self.label_textbox.ForeColor = System.Drawing.Color.Black
        self.label_textbox.Font = Font("Arial", 12, FontStyle.Bold)
        self.label_textbox.Location = Point(5, 20)
        self.label_textbox.Size = Size(110, 40)
        self.Controls.Add(self.label_textbox)

        self.label_textbox2 = Label()
        self.label_textbox2.Text = 'Map Name:'
        self.label_textbox2.ForeColor = System.Drawing.Color.Black
        self.label_textbox2.Font = Font("Arial", 12, FontStyle.Bold)
        self.label_textbox2.Location = Point(5, 70)
        self.label_textbox2.Size = Size(110, 40)
        self.Controls.Add(self.label_textbox2)

        self.textBox1 = TextBox()
        self.textBox1.Text = lines[0]
        self.textBox1.Location = Point(125, 20)
        self.textBox1.Size = Size(125, 40)
        self.Controls.Add(self.textBox1)

        self.textBox2 = TextBox()
        self.textBox2.Text = lines[1]
        self.textBox2.Location = Point(125, 70)
        self.textBox2.Size = Size(125, 40)
        self.Controls.Add(self.textBox2)

        self.button = Button()
        self.button.Text = 'Set Spool Data'
        self.button.Location = Point(80, 120)
        self.button.Size = Size(100, 30)
        self.Controls.Add(self.button)

        self.button.Click += self.on_click

    def on_click(self, sender, event):
        try:
            self.value = self.textBox1.Text
            self.map_name = self.textBox2.Text
            
            # Write data to elements
            selection = revit.get_selection()
            if not selection:
                MessageBox.Show("No elements selected. Please select elements.")
                return
            
            t = Transaction(doc, 'Set Assembly Number')
            t.Start()
            for i in selection:
                param_exist = i.LookupParameter("STRATUS Assembly")
                isfabpart = i.LookupParameter("Fabrication Service")
                if isfabpart:
                    stat = i.PartStatus
                    STName = Config.GetPartStatusDescription(stat)
                    set_parameter_by_name(i, "STRATUS Assembly", self.value)
                    set_parameter_by_name(i, "FP_Spool Map", self.map_name)
                    set_parameter_by_name(i, "STRATUS Status", "Modeled")
                    i.SpoolName = self.value
                    i.PartStatus = 1
                    i.Pinned = True
                elif param_exist:
                    set_parameter_by_name(i, "STRATUS Assembly", self.value)
                    set_parameter_by_name(i, "FP_Spool Map", self.map_name)
                    set_parameter_by_name(i, "STRATUS Status", "Modeled")
                    i.Pinned = True
                else:
                    print('Elements missing parameters to modify, contact admin.')
            t.Commit()

            # Increment the spool name and map name
            valuesplit = self.value.rsplit('-', 1)
            spoolnumlength = len(valuesplit[-1])
            test = valuesplit[-1].isnumeric()

            if test:
                valuenum = int(valuesplit[-1])
                numincrement = valuenum + 1
                firstpart = valuesplit[0]
                lastpart = str(numincrement).zfill(spoolnumlength)
                newspoolname = firstpart + "-" + lastpart

                # Update the text file with the new values
                with open(filepath, 'w') as the_file:
                    line1 = newspoolname + '\n'
                    line2 = self.map_name + '\n'
                    the_file.writelines([line1, line2])

                # Update the text boxes with the new values
                self.textBox1.Text = newspoolname
                self.textBox2.Text = self.map_name
            else:
                with open(filepath, 'w') as the_file:
                    line1 = self.value + '\n'
                    line2 = self.map_name + '\n'
                    the_file.writelines([line1, line2])
        except ValueError:
            MessageBox.Show("Please enter valid data.")
        except Exception as e:
            MessageBox.Show('oof, not sure what happened! Contact admin.')

# Display the form modelessly
form = TXT_Form()
form.Show()

# Event loop to keep the form open
while form.Visible:
    System.Windows.Forms.Application.DoEvents()
