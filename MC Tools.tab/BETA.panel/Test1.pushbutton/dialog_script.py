# -*- coding: utf-8 -*-
# Keep the script engine alive after the script body ends. This window is modeless and its
# event handlers run later; without this, their globals are gone and Revit can crash.
__persistentengine__ = True

import os
import json
import clr

# Load .NET Framework & WPF Assemblies
clr.AddReference("System")
clr.AddReference("System.Drawing")
clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")
clr.AddReference("Microsoft.Web.WebView2.Wpf")
clr.AddReference("Microsoft.Web.WebView2.Core")

import System
from System.Windows import (
    Window, WindowStartupLocation, WindowStyle, ResizeMode,
    Thickness, FontWeights, HorizontalAlignment, VerticalAlignment,
    TextAlignment, GridLength, GridUnitType
)
from System.Windows.Controls import Grid, RowDefinition, Border, TextBlock, DockPanel, Dock
from System.Windows.Interop import WindowInteropHelper
from System.Windows.Media import SolidColorBrush, ColorConverter, FontFamily, Brushes
from Microsoft.Web.WebView2.Wpf import WebView2

from pyrevit import revit, script, forms
from Autodesk.Revit.DB import FilteredElementCollector, View, Transaction, ElementId
from Autodesk.Revit.UI import IExternalEventHandler, ExternalEvent

# Murray palette (black, white, yellow only). Yellow is fill/accent, never text.
MC_YELLOW = "#FFE716"
MC_BLACK = "#231F20"
MC_WHITE = "#FFFFFF"

# IronPython 2 (basestring) / CPython 3 (str) compatibility
try:
    STR_TYPES = basestring
except NameError:
    STR_TYPES = str


LOG_PATH = os.path.join(os.environ.get("APPDATA", "."), "pyRevitWebView2Data", "renamer.log")


def log(msg):
    # File log. Never print() from window/event callbacks: pyRevit's output stream
    # is closed once the script body finishes, and writing to it can fault Revit.
    try:
        folder = os.path.dirname(LOG_PATH)
        if not os.path.isdir(folder):
            os.makedirs(folder)
        with open(LOG_PATH, "a") as f:
            f.write(str(msg) + "\n")
    except Exception:
        pass


def brush(hex_color):
    return SolidColorBrush(ColorConverter.ConvertFromString(hex_color))


doc = revit.doc

# Excluded internal/system views
EXCLUDED_TYPES = ["ProjectBrowser", "SystemBrowser", "Undefined", "Internal"]

# 1. Collect non-template, non-browser views from active Revit model
views = [
    v for v in FilteredElementCollector(doc).OfClass(View).WhereElementIsNotElementType()
    if not v.IsTemplate and v.Name and str(v.ViewType) not in EXCLUDED_TYPES
]

view_data = []
for v in views:
    view_data.append({
        "id": v.Id.IntegerValue,
        "type": str(v.ViewType),
        "name": v.Name
    })

# Escape "</" so a view name can never close the script tag
json_views = json.dumps(view_data).replace("</", "<\\/")

# 2. Web Application UI (HTML / CSS / JavaScript)
html_content = """
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <title>Renamer</title>
    <style>
        :root { --mc-yellow: #FFE716; --mc-black: #231F20; --mc-white: #FFFFFF; }
        * { box-sizing: border-box; }
        html, body { height: 100%; }
        body {
            font-family: Arial, Helvetica, sans-serif;
            background: var(--mc-white);
            color: var(--mc-black);
            margin: 0;
            padding: 14px;
            display: flex;
            flex-direction: column;
            user-select: none;
        }

        .grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(160px, 1fr)); gap: 10px; margin-bottom: 12px; }
        .card { background: var(--mc-white); padding: 10px; border: 1px solid var(--mc-black); border-top: 5px solid var(--mc-yellow); }
        .card label { font-size: 10px; text-transform: uppercase; font-weight: bold; display: block; margin-bottom: 4px; letter-spacing: 0.5px; }
        input[type="text"] {
            width: 100%; padding: 6px; background: var(--mc-white); color: var(--mc-black);
            border: 1px solid var(--mc-black); font-family: Arial, Helvetica, sans-serif; font-size: 12px;
        }
        input[type="text"]:focus { outline: none; box-shadow: 0 0 0 3px var(--mc-yellow); }
        input[type="checkbox"] { accent-color: var(--mc-black); }

        .filter-bar { display: flex; gap: 6px; flex-wrap: wrap; margin-bottom: 12px; }
        .chip { background: var(--mc-white); padding: 4px 10px; font-size: 11px; cursor: pointer; border: 1px solid var(--mc-black); color: var(--mc-black); }
        .chip:hover { background: var(--mc-yellow); }
        .chip.active { background: var(--mc-yellow); font-weight: bold; box-shadow: inset 0 -3px 0 var(--mc-black); }

        .table-container { flex: 1; min-height: 120px; overflow-y: auto; border: 1px solid var(--mc-black); background: var(--mc-white); }
        table { width: 100%; border-collapse: collapse; }
        th, td { padding: 7px 10px; text-align: left; border-bottom: 1px solid var(--mc-black); font-size: 12px; }
        th { background: var(--mc-yellow); color: var(--mc-black); position: sticky; top: 0; z-index: 2; font-size: 11px; text-transform: uppercase; border-bottom: 2px solid var(--mc-black); }
        tr:hover td { background: var(--mc-yellow); }
        .new-name-cell { font-weight: bold; }

        .status { font-size: 10px; padding: 2px 6px; border: 1px solid var(--mc-black); background: var(--mc-white); color: var(--mc-black); font-weight: bold; }
        .status.changed { background: var(--mc-yellow); }
        .status.skip { border-style: dashed; }

        .action-btn {
            margin-top: 12px; width: 100%; padding: 12px; background: var(--mc-yellow); color: var(--mc-black);
            font-family: Arial, Helvetica, sans-serif; font-weight: bold; border: 2px solid var(--mc-black);
            cursor: pointer; font-size: 14px; text-transform: uppercase; letter-spacing: 0.5px;
        }
        .action-btn:hover { background: var(--mc-black); color: var(--mc-white); }
        .action-btn.done { background: var(--mc-black); color: var(--mc-white); }
    </style>
</head>
<body>
    <div class="grid">
        <div class="card">
            <label>Find / Replace</label>
            <input type="text" id="findVal" placeholder="Find..." oninput="updatePreview()">
            <input type="text" id="replaceVal" placeholder="Replace..." oninput="updatePreview()" style="margin-top:4px;">
        </div>
        <div class="card">
            <label>Prefix</label>
            <input type="text" id="prefixVal" placeholder="e.g. L1-" oninput="updatePreview()">
        </div>
        <div class="card">
            <label>Suffix</label>
            <input type="text" id="suffixVal" placeholder="e.g. -DRAFT" oninput="updatePreview()">
        </div>
        <div class="card">
            <label>Search Views</label>
            <input type="text" id="searchVal" placeholder="Filter view names..." oninput="filterTable()">
        </div>
    </div>

    <div class="filter-bar" id="typeFilters"></div>

    <div class="table-container">
        <table>
            <thead>
                <tr>
                    <th style="width:30px;"><input type="checkbox" id="selectAll" checked onclick="toggleAll(this)"></th>
                    <th>Type</th>
                    <th>Current Name</th>
                    <th>New Name Preview</th>
                    <th>Status</th>
                </tr>
            </thead>
            <tbody id="viewTableBody"></tbody>
        </table>
    </div>

    <button class="action-btn" id="submitBtn" onclick="submitRenames()">Rename views in Revit</button>

    <script>
        let rawViews = """ + json_views + """ ;
        let selectedType = "ALL";

        const preferredOrder = ["DrawingSheet", "Schedule", "FloorPlan", "ThreeD", "Section", "DraftingView"];

        function esc(s) {
            return String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
        }

        function post(obj) {
            if (window.chrome && window.chrome.webview) {
                window.chrome.webview.postMessage(JSON.stringify(obj));
            }
        }

        function init() {
            document.addEventListener("keydown", function(e) {
                if (e.key === "Escape") { post({ action: "close" }); }
            });

            const foundTypes = [...new Set(rawViews.map(v => v.type))].filter(
                t => t !== "ProjectBrowser" && t !== "SystemBrowser"
            );

            foundTypes.sort((a, b) => {
                let idxA = preferredOrder.indexOf(a);
                let idxB = preferredOrder.indexOf(b);
                if (idxA === -1) idxA = 999;
                if (idxB === -1) idxB = 999;
                if (idxA !== idxB) return idxA - idxB;
                return a.localeCompare(b);
            });

            const types = ["ALL", ...foundTypes];
            const filterBar = document.getElementById("typeFilters");
            filterBar.innerHTML = "";

            types.forEach(t => {
                const chip = document.createElement("div");
                chip.className = "chip" + (t === selectedType ? " active" : "");
                chip.innerText = t;
                chip.onclick = () => {
                    document.querySelectorAll(".chip").forEach(c => c.classList.remove("active"));
                    chip.classList.add("active");
                    selectedType = t;
                    filterTable();
                };
                filterBar.appendChild(chip);
            });
            renderRows(rawViews);
            updatePreview();
        }

        function filterTable() {
            const search = document.getElementById("searchVal").value.toLowerCase();
            const filtered = rawViews.filter(v => {
                const matchType = selectedType === "ALL" || v.type === selectedType;
                const matchSearch = v.name.toLowerCase().includes(search);
                return matchType && matchSearch;
            });
            renderRows(filtered);
            updatePreview();
        }

        function renderRows(views) {
            const tbody = document.getElementById("viewTableBody");
            tbody.innerHTML = "";
            views.forEach(v => {
                const tr = document.createElement("tr");
                tr.setAttribute("data-id", v.id);
                tr.setAttribute("data-orig", v.name);
                tr.innerHTML = `
                    <td><input type="checkbox" class="row-select" checked onchange="updatePreview()"></td>
                    <td>${esc(v.type)}</td>
                    <td class="current-name-cell">${esc(v.name)}</td>
                    <td class="new-name-cell">${esc(v.name)}</td>
                    <td><span class="status">UNCHANGED</span></td>
                `;
                tbody.appendChild(tr);
            });
        }

        function toggleAll(src) {
            document.querySelectorAll(".row-select").forEach(cb => cb.checked = src.checked);
            updatePreview();
        }

        function updatePreview() {
            const findVal = document.getElementById("findVal").value;
            const replaceVal = document.getElementById("replaceVal").value;
            const prefixVal = document.getElementById("prefixVal").value;
            const suffixVal = document.getElementById("suffixVal").value;

            let modifiedCount = 0;

            document.querySelectorAll("#viewTableBody tr").forEach(tr => {
                const cb = tr.querySelector(".row-select");
                const orig = tr.getAttribute("data-orig");
                const previewCell = tr.querySelector(".new-name-cell");
                const statusSpan = tr.querySelector(".status");

                if (cb.checked) {
                    let newName = orig;
                    if (findVal) newName = newName.split(findVal).join(replaceVal);
                    newName = prefixVal + newName + suffixVal;

                    previewCell.innerText = newName;
                    if (newName !== orig) {
                        statusSpan.innerText = "MODIFIED";
                        statusSpan.className = "status changed";
                        modifiedCount++;
                    } else {
                        statusSpan.innerText = "UNCHANGED";
                        statusSpan.className = "status";
                    }
                } else {
                    previewCell.innerText = orig;
                    statusSpan.innerText = "SKIP";
                    statusSpan.className = "status skip";
                }
            });

            const btn = document.getElementById("submitBtn");
            if (!btn.classList.contains("done")) {
                btn.innerText = `Rename ${modifiedCount} view(s) in Revit`;
            }
        }

        function submitRenames() {
            const updates = [];
            document.querySelectorAll("#viewTableBody tr").forEach(tr => {
                const cb = tr.querySelector(".row-select");
                if (cb && cb.checked) {
                    const id = parseInt(tr.getAttribute("data-id"));
                    const orig = tr.getAttribute("data-orig");
                    const newName = tr.querySelector(".new-name-cell").innerText;
                    if (newName !== orig) {
                        updates.push({ id: id, name: newName });
                    }
                }
            });

            if (updates.length === 0) {
                alert("No view names were modified.");
                return;
            }

            post({ action: "rename", data: updates });
        }

        window.onRenameSuccess = function(successList) {
            successList.forEach(s => {
                const item = rawViews.find(v => v.id === s.id);
                if (item) {
                    item.name = s.newName;
                }
            });

            document.getElementById("findVal").value = "";
            document.getElementById("replaceVal").value = "";
            document.getElementById("prefixVal").value = "";
            document.getElementById("suffixVal").value = "";

            const btn = document.getElementById("submitBtn");
            btn.classList.add("done");
            btn.innerText = `Renamed ${successList.length} view(s)`;

            filterTable();

            setTimeout(() => {
                btn.classList.remove("done");
                updatePreview();
            }, 2000);
        };

        window.onload = init;
    </script>
</body>
</html>
"""


# 3. External Event Handler
class BatchRenameHandler(IExternalEventHandler):
    def __init__(self):
        self.updates = []
        self.window = None

    def Execute(self, uiapp):
        import json
        from Autodesk.Revit.DB import Transaction, ElementId
        try:
            active_doc = uiapp.ActiveUIDocument.Document
            if not self.updates:
                return

            t = Transaction(active_doc, "WebView2 Batch View Rename")
            t.Start()
            successful_renames = []
            for item in self.updates:
                elem_id = ElementId(int(item["id"]))
                v_elem = active_doc.GetElement(elem_id)
                if v_elem:
                    try:
                        v_elem.Name = item["name"]
                        successful_renames.append({"id": item["id"], "newName": item["name"]})
                    except Exception as ex:
                        log("Error renaming '{}': {}".format(v_elem.Name, str(ex)))
            t.Commit()

            if self.window is not None and self.window.browser.CoreWebView2:
                js_callback = "if (window.onRenameSuccess) { window.onRenameSuccess(" + json.dumps(successful_renames) + "); }"
                self.window.browser.CoreWebView2.ExecuteScriptAsync(js_callback)

        except Exception as ex:
            log("Error in ExternalEvent: " + str(ex))

    def GetName(self):
        return "WebView2BatchRenameHandler"


# 4. Modeless Frameless WPF Host Window Class
class WebViewHostWindow(Window):
    def __init__(self, html, ext_handler, ext_event):
        self.Title = "Renamer"
        self.Width = 850
        self.Height = 620
        self.MinWidth = 520
        self.MinHeight = 400
        self.WindowStartupLocation = WindowStartupLocation.CenterScreen
        self.html_payload = html
        self.ext_handler = ext_handler
        self.ext_event = ext_event
        self.ext_handler.window = self

        self.WindowStyle = WindowStyle.None
        # CanResize keeps an invisible resize border (a grip would be hidden behind WebView2)
        self.ResizeMode = ResizeMode.CanResize
        self.Background = brush(MC_WHITE)

        helper = WindowInteropHelper(self)
        helper.Owner = revit.HOST_APP.proc.MainWindowHandle

        user_data = System.IO.Path.Combine(
            System.Environment.GetFolderPath(System.Environment.SpecialFolder.ApplicationData),
            "pyRevitWebView2Data"
        )
        if not System.IO.Directory.Exists(user_data):
            try:
                System.IO.Directory.CreateDirectory(user_data)
            except:
                pass
        System.Environment.SetEnvironmentVariable("WEBVIEW2_USER_DATA_FOLDER", user_data)

        # --- Layout: native WPF title bar (row 0) + WebView2 (row 1) ---
        root = Grid()
        row_title = RowDefinition()
        row_title.Height = GridLength(34)
        row_body = RowDefinition()
        row_body.Height = GridLength(1, GridUnitType.Star)
        root.RowDefinitions.Add(row_title)
        root.RowDefinitions.Add(row_body)

        # Title bar: Murray Yellow fill, black text, black rule underneath
        self.title_bar = Border()
        self.title_bar.Background = brush(MC_YELLOW)
        self.title_bar.BorderBrush = brush(MC_BLACK)
        self.title_bar.BorderThickness = Thickness(0, 0, 0, 2)
        Grid.SetRow(self.title_bar, 0)

        dock = DockPanel()
        dock.LastChildFill = True

        # Close button (native)
        self.close_box = Border()
        self.close_box.Width = 44
        self.close_box.Background = Brushes.Transparent
        self.close_glyph = TextBlock()
        self.close_glyph.Text = u"\u00D7"
        self.close_glyph.FontFamily = FontFamily("Arial")
        self.close_glyph.FontSize = 20
        self.close_glyph.FontWeight = FontWeights.Bold
        self.close_glyph.Foreground = brush(MC_BLACK)
        self.close_glyph.HorizontalAlignment = HorizontalAlignment.Center
        self.close_glyph.VerticalAlignment = VerticalAlignment.Center
        self.close_box.Child = self.close_glyph
        DockPanel.SetDock(self.close_box, Dock.Right)
        dock.Children.Add(self.close_box)

        # Title text
        title_text = TextBlock()
        title_text.Text = "RENAMER"
        title_text.FontFamily = FontFamily("Arial")
        title_text.FontSize = 13
        title_text.FontWeight = FontWeights.Bold
        title_text.Foreground = brush(MC_BLACK)
        title_text.VerticalAlignment = VerticalAlignment.Center
        title_text.Margin = Thickness(12, 0, 0, 0)
        dock.Children.Add(title_text)

        self.title_bar.Child = dock
        root.Children.Add(self.title_bar)

        # Browser
        self.browser = WebView2()
        self.browser.DefaultBackgroundColor = System.Drawing.Color.White
        Grid.SetRow(self.browser, 1)
        root.Children.Add(self.browser)

        self.Content = root

        # Brushes kept on the window so event handlers do not depend on module globals
        self.br_black = brush(MC_BLACK)
        self.br_white = brush(MC_WHITE)

        # Native events (these always fire on the UI thread, no Dispatcher needed)
        self.title_bar.MouseLeftButtonDown += self.on_title_mouse_down
        self.close_box.MouseLeftButtonDown += self.on_close_mouse_down
        self.close_box.MouseLeftButtonUp += self.on_close_mouse_up
        self.close_box.MouseEnter += self.on_close_enter
        self.close_box.MouseLeave += self.on_close_leave

        self.Loaded += self.on_window_loaded
        self.Closed += self.on_closed

    def on_closed(self, sender, args):
        try:
            # Detach so no late callback touches this window. No Dispose here:
            # disposing WebView2 inside Revit is a known source of crashes on close.
            self.ext_handler.window = None
            self.ext_handler.updates = []
        except Exception:
            pass
        try:
            import System as _System
            _System.AppDomain.CurrentDomain.SetData("MurrayRenamerWindow", None)
        except Exception:
            pass

    # ----- Title bar behavior (every handler is self-contained and cannot raise) -----
    def on_title_mouse_down(self, sender, args):
        try:
            self.DragMove()
        except Exception:
            pass

    def on_close_mouse_down(self, sender, args):
        try:
            args.Handled = True
        except Exception:
            pass

    def on_close_mouse_up(self, sender, args):
        try:
            self.Close()
        except Exception:
            pass

    def on_close_enter(self, sender, args):
        try:
            self.close_box.Background = self.br_black
            self.close_glyph.Foreground = self.br_white
        except Exception:
            pass

    def on_close_leave(self, sender, args):
        try:
            from System.Windows.Media import Brushes as _Brushes
            self.close_box.Background = _Brushes.Transparent
            self.close_glyph.Foreground = self.br_black
        except Exception:
            pass

    # ----- WebView2 -----
    def on_window_loaded(self, sender, args):
        try:
            self.browser.CoreWebView2InitializationCompleted += self.on_webview_initialized
            self.browser.EnsureCoreWebView2Async(None)
        except Exception as ex:
            log("Loaded error: " + str(ex))

    def on_webview_initialized(self, sender, args):
        try:
            if args.IsSuccess:
                self.browser.CoreWebView2.WebMessageReceived += self.on_web_message
                self.browser.CoreWebView2.NavigateToString(self.html_payload)
            else:
                log("WebView2 Initialization Failed: " + str(args.InitializationException))
        except Exception as ex:
            log("Init error: " + str(ex))

    def on_web_message(self, sender, args):
        import json
        try:
            data = json.loads(args.WebMessageAsJson)

            try:
                str_type = basestring
            except NameError:
                str_type = str
            payload = json.loads(data) if isinstance(data, str_type) else data

            if isinstance(payload, dict):
                msg_action = payload.get("action")
                if msg_action == "close":
                    self.Close()
                    return
                if msg_action == "rename":
                    self.ext_handler.updates = payload.get("data", [])
                    self.ext_event.Raise()
        except Exception as ex:
            try:
                log("Error processing web message: " + str(ex))
            except Exception:
                pass


# 5. Launch Modeless Window (single instance)
_existing = System.AppDomain.CurrentDomain.GetData("MurrayRenamerWindow")
_already_open = False
try:
    if _existing is not None and _existing.IsLoaded:
        _existing.Activate()
        _already_open = True
except Exception:
    _already_open = False

if not _already_open:
    rename_handler = BatchRenameHandler()
    ext_event = ExternalEvent.Create(rename_handler)

    window_inst = WebViewHostWindow(html_content, rename_handler, ext_event)
    System.AppDomain.CurrentDomain.SetData("MurrayRenamerWindow", window_inst)
    window_inst.Show()