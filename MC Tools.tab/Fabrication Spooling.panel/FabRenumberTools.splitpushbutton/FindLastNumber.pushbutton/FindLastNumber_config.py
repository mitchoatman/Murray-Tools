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
class UpdateItemNumbersHandler(IExternalEventHandler):
    def __init__(self, update_callback):
        self.update_callback = update_callback
    
    def Execute(self, app):
        self.update_callback()
    
    def GetName(self):
        return "Update Item Numbers"

# WPF Window for modeless display
class ValveNumberWindow(Window):
    def __init__(self, doc, uidoc, app):
        # Import DocumentChangedEventArgs here to ensure availability
        from Autodesk.Revit.DB.Events import DocumentChangedEventArgs
        
        self.doc = doc
        self.uidoc = uidoc
        self.app = app
        self.Title = "Item Number Monitor"
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
        self.update_handler = UpdateItemNumbersHandler(self.update_ui)
        self.external_event = ExternalEvent.Create(self.update_handler)
        
        # Subscribe to application-level DocumentChanged event
        try:
            self.app.DocumentChanged += self.on_document_changed
            self.text_block.Text = "Monitoring started..."
            # Log subscription success
            #try:
            #    log_dir = r"C:\Temp"
            #    if not os.path.exists(log_dir):
            #        os.makedirs(log_dir)
            #    log_path = os.path.join(log_dir, "ItemNumberDisplay_error.log")
            #    with open(log_path, "a") as log_file:
            #        log_file.write("Application.DocumentChanged event subscribed for Item Numbers at {}\n".format(time.strftime("%Y-%m-%d %H:%M:%S")))
            #except:
            #    pass  # Skip logging if path is inaccessible
        except Exception as e:
            self.text_block.Text = "Failed to subscribe to DocumentChanged event: {}".format(str(e))
            # Log error to C:\Temp
            #try:
            #    log_dir = r"C:\Temp"
            #    if not os.path.exists(log_dir):
            #        os.makedirs(log_dir)
            #    log_path = os.path.join(log_dir, "ItemNumberDisplay_error.log")
            #    with open(log_path, "a") as log_file:
            #        log_file.write("DocumentChanged subscription error: {}\n".format(str(e)))
            #        log_file.write(traceback.format_exc() + "\n")
            #except:
            #    pass
        
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
            # Reinitialize document and UI document to ensure valid context
            global doc, uidoc
            doc = __revit__.ActiveUIDocument.Document
            uidoc = __revit__.ActiveUIDocument
            
            # Verify active view is valid
            curview = doc.ActiveView
            if curview is None:
                self.text_block.Text = "Error: No active view available."
                return
            
            # Explicitly re-import FilteredElementCollector and FabricationPart
            try:
                from Autodesk.Revit.DB import FilteredElementCollector, FabricationPart
                part_collector = FilteredElementCollector(doc, curview.Id).OfClass(FabricationPart) \
                                .WhereElementIsNotElementType() \
                                .ToElements()
            except Exception as e:
                self.text_block.Text = "Error in FilteredElementCollector or FabricationPart: {}".format(str(e))
                # Log error to C:\Temp
                #try:
                #    log_dir = r"C:\Temp"
                #    if not os.path.exists(log_dir):
                #        os.makedirs(log_dir)
                #    log_path = os.path.join(log_dir, "ItemNumberDisplay_error.log")
                #    with open(log_path, "a") as log_file:
                #        log_file.write("FilteredElementCollector/FabricationPart error: {}\n".format(str(e)))
                #        log_file.write(traceback.format_exc() + "\n")
                #except:
                #    pass
                return
            
            fp_valve_numbers = []
            for element in part_collector:
                try:
                    param = element.LookupParameter('Item Number')
                    if param and param.HasValue:
                        valve_number = param.AsString()
                        if valve_number:
                            fp_valve_numbers.append(valve_number)
                except Exception as e:
                    self.text_block.Text = "Error accessing FP_Valve Number parameter: {}".format(str(e))
                    # Log error to C:\Temp
                    #try:
                    #    log_dir = r"C:\Temp"
                    #    if not os.path.exists(log_dir):
                    #        os.makedirs(log_dir)
                    #    log_path = os.path.join(log_dir, "ItemNumberDisplay_error.log")
                    #    with open(log_path, "a") as log_file:
                    #        log_file.write("FP_Valve Number parameter error: {}\n".format(str(e)))
                    #        log_file.write(traceback.format_exc() + "\n")
                    #except:
                    #    pass
                    return
            
            result = self.get_format_and_number(fp_valve_numbers)
            result_text = ""
            if result:
                for prefix, max_number in result.items():
                    result_text += "Prefix '{}': Maximum Item Number {}{}\n".format(prefix, prefix, max_number)
            else:
                result_text = "No valid FP_Valve Numbers found in the current view."
            
            self.text_block.Text = result_text
            # Log successful update
            #try:
            #    log_dir = r"C:\Temp"
            #    if not os.path.exists(log_dir):
            #        os.makedirs(log_dir)
            #    log_path = os.path.join(log_dir, "ItemNumberDisplay_error.log")
            #    with open(log_path, "a") as log_file:
            #        log_file.write("UI updated for Item Numbers at {}\n".format(time.strftime("%Y-%m-%d %H:%M:%S")))
            #except:
            #    pass
        except Exception as e:
            self.text_block.Text = "General error updating Item numbers: {}".format(str(e))
            # Log error to C:\Temp
            #try:
            #    log_dir = r"C:\Temp"
            #    if not os.path.exists(log_dir):
            #        os.makedirs(log_dir)
            #    log_path = os.path.join(log_dir, "ItemNumberDisplay_error.log")
            #    with open(log_path, "a") as log_file:
            #        log_file.write("General error: {}\n".format(str(e)))
            #        log_file.write(traceback.format_exc() + "\n")
            #except:
            #    pass
    
    def on_document_changed(self, sender, args):
        try:
            # Check if the changed document is the active document
            if args.GetDocument().Equals(self.doc):
                # Log event trigger
                #try:
                #    log_dir = r"C:\Temp"
                #    if not os.path.exists(log_dir):
                #        os.makedirs(log_dir)
                #    log_path = os.path.join(log_dir, "ItemNumberDisplay_error.log")
                #    with open(log_path, "a") as log_file:
                #        log_file.write("DocumentChanged event triggered for Item Numbers at {}\n".format(time.strftime("%Y-%m-%d %H:%M:%S")))
                #except:
                #    pass
                self.external_event.Raise()
        except Exception as e:
            pass
            # Log error to C:\Temp
            #try:
            #    log_dir = r"C:\Temp"
            #    if not os.path.exists(log_dir):
            #        os.makedirs(log_dir)
            #    log_path = os.path.join(log_dir, "ItemNumberDisplay_error.log")
            #    with open(log_path, "a") as log_file:
            #        log_file.write("DocumentChanged handler error: {}\n".format(str(e)))
            #        log_file.write(traceback.format_exc() + "\n")
            #except:
            #    pass
    
    def Close(self):
        # Clean up
        try:
            self.app.DocumentChanged -= self.on_document_changed
            # Log unsubscription
            #try:
            #    log_dir = r"C:\Temp"
            #    if not os.path.exists(log_dir):
            #        os.makedirs(log_dir)
            #    log_path = os.path.join(log_dir, "ItemNumberDisplay_error.log")
            #    with open(log_path, "a") as log_file:
            #        log_file.write("Application.DocumentChanged event unsubscribed for Item Numbers at {}\n".format(time.strftime("%Y-%m-%d %H:%M:%S")))
            #except:
            #    pass
        except Exception as e:
            pass
            # Log unsubscription error
            #try:
            #    log_dir = r"C:\Temp"
            #    if not os.path.exists(log_dir):
            #        os.makedirs(log_dir)
            #    log_path = os.path.join(log_dir, "ItemNumberDisplay_error.log")
            #    with open(log_path, "a") as log_file:
            #        log_file.write("DocumentChanged unsubscription error: {}\n".format(str(e)))
            #        log_file.write(traceback.format_exc() + "\n")
            #except:
            #    pass
        super(ValveNumberWindow, self).Close()

# Create and show the modeless window
window = ValveNumberWindow(doc, uidoc, app)
window.Show()