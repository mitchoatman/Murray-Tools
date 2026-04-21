# -*- coding: UTF-8 -*-
from Autodesk.Revit.DB import Transaction, FabricationConfiguration
import os
from Parameters.Add_SharedParameters import Shared_Params
from Parameters.Get_Set_Params import set_parameter_by_name
Shared_Params()

# .NET Imports
import clr
clr.AddReference('PresentationCore')
clr.AddReference('PresentationFramework')
clr.AddReference('WindowsBase')
clr.AddReference('System')
clr.AddReference('System.Windows.Forms')
import System
from System.Windows import Application, Window, Thickness, HorizontalAlignment, VerticalAlignment, WindowStartupLocation, ResizeMode, FontWeights
from System.Windows.Controls import Button, TextBox, Label, Grid, RowDefinition, ColumnDefinition
from System.Windows.Media import Brushes, FontFamily
from System.Windows import Size
from System.Windows.Threading import DispatcherFrame, Dispatcher

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

class TXT_Form(Window):
    def __init__(self):
        self.Title = 'Spool Data'
        self.Width = 275
        self.Height = 200
        self.MinWidth = 275
        self.MinHeight = 200
        self.WindowStartupLocation = WindowStartupLocation.CenterScreen
        self.Topmost = True
        self.ResizeMode = ResizeMode.NoResize

        # Create grid layout
        grid = Grid()
        for i in range(4):  # 4 rows: label1, textbox1, label2, textbox2, button
            grid.RowDefinitions.Add(RowDefinition())
        for i in range(2):  # 2 columns: labels, textboxes
            grid.ColumnDefinitions.Add(ColumnDefinition())

        # Label 1 (Spool Name)
        label1 = Label()
        label1.Content = 'Spool Name:'
        label1.FontFamily = FontFamily("Arial")
        label1.FontSize = 14
        label1.FontWeight = FontWeights.Bold
        label1.Margin = Thickness(5, 5, 0, 0)
        Grid.SetRow(label1, 0)
        Grid.SetColumn(label1, 0)
        grid.Children.Add(label1)

        # TextBox 1
        textbox1 = TextBox()
        textbox1.Text = lines[0]
        textbox1.Margin = Thickness(-10, 5, 5, 0)
        textbox1.FontFamily = FontFamily("Arial")
        textbox1.FontSize = 12
        textbox1.Height = 20
        textbox1.IsReadOnly = False
        textbox1.IsEnabled = True
        Grid.SetRow(textbox1, 0)
        Grid.SetColumn(textbox1, 1)
        grid.Children.Add(textbox1)

        # Label 2 (Map Name)
        label2 = Label()
        label2.Content = 'Map Name:'
        label2.FontFamily = FontFamily("Arial")
        label2.FontSize = 14
        label2.FontWeight = FontWeights.Bold
        label2.Margin = Thickness(5, 5, 0, 0)
        Grid.SetRow(label2, 1)
        Grid.SetColumn(label2, 0)
        grid.Children.Add(label2)

        # TextBox 2
        textbox2 = TextBox()
        textbox2.Text = lines[1]
        textbox2.Margin = Thickness(-10, 5, 5, 0)
        textbox2.FontFamily = FontFamily("Arial")
        textbox2.FontSize = 12
        textbox2.Height = 20
        textbox2.IsReadOnly = False
        textbox2.IsEnabled = True
        Grid.SetRow(textbox2, 1)
        Grid.SetColumn(textbox2, 1)
        grid.Children.Add(textbox2)

        # Button
        button = Button()
        button.Content = 'Set Spool Data'
        button.Margin = Thickness(80, -5, 80, 5)
        button.FontFamily = FontFamily("Arial")
        button.FontSize = 12
        button.Height = 25
        button.Click += self.on_click
        Grid.SetRow(button, 3)
        Grid.SetColumnSpan(button, 2)
        grid.Children.Add(button)

        self.Content = grid
        self.value = None
        self.map_name = None
        textbox1.Focus()  # Set focus to Spool Name textbox
        textbox1.SelectAll()  # Select all text in Spool Name textbox

    def on_click(self, sender, event):
        try:
            self.value = self.Content.Children[1].Text  # textbox1
            self.map_name = self.Content.Children[3].Text  # textbox2
            
            # Write data to elements
            selection = uidoc.Selection.GetElementIds()
            if not selection:
                from System.Windows.Forms import MessageBox
                MessageBox.Show("No elements selected. Please select elements.", "Selection Error")
                return
            
            t = Transaction(doc, 'Set Assembly Number')
            t.Start()
            for element_id in selection:
                i = doc.GetElement(element_id)  # Convert ElementId to Element
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
                self.Content.Children[1].Text = newspoolname
                self.Content.Children[3].Text = self.map_name
            else:
                with open(filepath, 'w') as the_file:
                    line1 = self.value + '\n'
                    line2 = self.map_name + '\n'
                    the_file.writelines([line1, line2])

        except ValueError:
            from System.Windows.Forms import MessageBox
            MessageBox.Show("Please enter valid data.", "Input Error")
        except Exception as e:
            from System.Windows.Forms import MessageBox
            MessageBox.Show('oof, not sure what happened! Contact admin.', "Error")

form = TXT_Form()

# Set up a WPF dispatcher frame
frame = DispatcherFrame()

# When form is closed, exit the frame cleanly
def exit_frame(sender, e):
    frame.Continue = False

form.Closed += exit_frame
form.Show()

# Keeps UI responsive, textboxes editable, and exits properly on close
Dispatcher.PushFrame(frame)