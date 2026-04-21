import Autodesk
import clr
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
clr.AddReference('PresentationFramework')
clr.AddReference('System.Windows.Forms')

from Autodesk.Revit import DB
from Autodesk.Revit.UI import ExternalEvent, IExternalEventHandler
from System.Windows import Window, Application
from System.Windows.Controls import TextBlock, StackPanel
from System import Windows
import System
import uuid
import traceback

DB = Autodesk.Revit.DB
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
app = __revit__.Application

# External event handler for updating the UI
class UpdateValveNumbersHandler(IExternalEventHandler):
    def __init__(self, update_callback):
        self.update_callback = update_callback
    
    def Execute(self, app):
        self.update_callback()
    
    def GetName(self):
        return "Update Valve Numbers"

# WPF Window for modeless display
class ValveNumberWindow(Window):
    def __init__(self, doc, uidoc, app):
        # Import DocumentChangedEventArgs here to ensure availability
        from Autodesk.Revit.DB.Events import DocumentChangedEventArgs
        
        self.doc = doc
        self.uidoc = uidoc
        self.app = app
        self.Title = "Valve Number Monitor"
        self.Width = 400
        self.Height = 300
        self.ResizeMode = Windows.ResizeMode.CanMinimize
        self.WindowStartupLocation = Windows.WindowStartupLocation.CenterScreen
        self.Topmost = True  # Make window always on top
        
        # Create UI elements
        self.stack_panel = StackPanel()
        self.text_block = TextBlock()
        self.text_block.Margin = Windows.Thickness(10)
        self.stack_panel.Children.Add(self.text_block)
        self.Content = self.stack_panel
        
        # Create external event
        self.update_handler = UpdateValveNumbersHandler(self.update_ui)
        self.external_event = ExternalEvent.Create(self.update_handler)
        
        # Subscribe to application-level DocumentChanged event
        try:
            self.app.DocumentChanged += self.on_document_changed
            self.text_block.Text = "Monitoring started..."

        except Exception as e:
            self.text_block.Text = "Failed to subscribe to DocumentChanged event: {}".format(str(e))
        
        # Initial UI update
        self.update_ui()
    
    def get_format_and_number(self, item_numbers):
        prefix_max_number = {}
        for item_number in item_numbers:
            if not item_number:  # Skip None or empty item numbers
                continue
            if item_number.isdigit():
                format_part = ''
            else:
                for i in range(len(item_number) - 1, -1, -1):
                    if not item_number[i].isdigit():
                        format_part = item_number[:i+1]
                        break
                else:
                    format_part = ''
            number_part = item_number[len(format_part):]
            if not number_part or not number_part.isdigit():
                continue
            if format_part not in prefix_max_number:
                prefix_max_number[format_part] = number_part
            else:
                prefix_max_number[format_part] = max(
                    prefix_max_number[format_part], 
                    number_part, 
                    key=lambda x: (len(x), x)
                )
        return prefix_max_number
    
    def update_ui(self):
        try:
            global doc, uidoc
            doc = __revit__.ActiveUIDocument.Document
            uidoc = __revit__.ActiveUIDocument
            
            # Verify active view is valid
            curview = doc.ActiveView
            if curview is None:
                self.text_block.Text = "Error: No active view available."
                return
            
            try:
                from Autodesk.Revit.DB import FilteredElementCollector#, FabricationPart
                part_collector = FilteredElementCollector(doc, curview.Id) \
                    .WhereElementIsNotElementType() \
                    .ToElements()
                
                # # Collect elements of FabricationPart, PipeAccessory, and DuctAccessory in the current view
                # part_collector = FilteredElementCollector(doc, curview.Id) \
                                # .WherePasses(category_filter) \
                                # .WhereElementIsNotElementType() \
                                # .ToElements()

            except Exception as e:
                self.text_block.Text = "Error in FilteredElementCollector or FabricationPart: {}".format(str(e))

                return
            
            fp_valve_numbers = []
            for element in part_collector:
                try:
                    param = element.LookupParameter('FP_Valve Number')
                    if param and param.HasValue:
                        valve_number = param.AsString()
                        if valve_number:
                            fp_valve_numbers.append(valve_number)
                except Exception as e:
                    self.text_block.Text = "Error accessing FP_Valve Number parameter: {}".format(str(e))

                    return
            
            result = self.get_format_and_number(fp_valve_numbers)
            result_text = ""
            if result:
                for prefix, max_number in result.items():
                    result_text += "Prefix '{}': Maximum Valve Number {}{}\n".format(prefix, prefix, max_number)
            else:
                result_text = "No valid FP_Valve Numbers found in the current view."
            
            self.text_block.Text = result_text
 
        except Exception as e:
            self.text_block.Text = "General error updating valve numbers: {}".format(str(e))
    
    def on_document_changed(self, sender, args):
        try:
            # Check if the changed document is the active document
            if args.GetDocument().Equals(self.doc):

                self.external_event.Raise()
        except Exception as e:
            pass
    
    def Close(self):
        # Clean up
        try:
            self.app.DocumentChanged -= self.on_document_changed

        except Exception as e:
            pass

        super(ValveNumberWindow, self).Close()

window = ValveNumberWindow(doc, uidoc, app)
window.Show()