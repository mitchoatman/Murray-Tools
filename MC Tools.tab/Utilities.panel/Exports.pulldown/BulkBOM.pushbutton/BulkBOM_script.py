#! python3
# -*- coding: utf-8 -*-


import os
import sys
import json
import time
import tempfile
import traceback

import clr
clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')
import System
from System import Uri, Environment
from System.Windows.Forms import (Form, DockStyle, FormStartPosition, FolderBrowserDialog, DialogResult,
                                  MessageBox, Application, Screen, Timer, NativeWindow)
from System.Drawing import Size

from System.Diagnostics import Process
from Autodesk.Revit.UI import IExternalEventHandler, ExternalEvent
from pyrevit import revit, forms

import murray_bom
from murray_bom import revit_read
from murray_bom.service import Service

HERE = os.path.dirname(os.path.abspath(murray_bom.__file__))
UI_PAGE = os.path.join(HERE, 'ui', 'index.html')
LOCAL_WEBVIEW2 = os.path.join(HERE, 'webview2')      # optional: drop the three WebView2 files here (see README)


def load_webview2():
    """Loads the WebView2 WinForms control. Tries what pyRevit/Revit already has, then the local webview2 folder."""
    def attempt():
        clr.AddReference('Microsoft.Web.WebView2.Core')
        clr.AddReference('Microsoft.Web.WebView2.WinForms')
    try:
        attempt()
        return True
    except Exception:
        pass
    # An assembly with this name may already be loaded from somewhere pyRevit knows about.
    for asm in System.AppDomain.CurrentDomain.GetAssemblies():
        if asm.GetName().Name.startswith('Microsoft.Web.WebView2'):
            d = os.path.dirname(asm.Location)
            if d and d not in sys.path:
                sys.path.append(d)
    try:
        attempt()
        return True
    except Exception:
        pass
    if os.path.isdir(LOCAL_WEBVIEW2):
        sys.path.append(LOCAL_WEBVIEW2)
        os.environ['PATH'] = LOCAL_WEBVIEW2 + os.pathsep + os.environ.get('PATH', '')
        try:
            clr.AddReference(os.path.join(LOCAL_WEBVIEW2, 'Microsoft.Web.WebView2.Core.dll'))
            clr.AddReference(os.path.join(LOCAL_WEBVIEW2, 'Microsoft.Web.WebView2.WinForms.dll'))
            return True
        except Exception:
            pass
    return False


def webview2_versions():
    out = []
    for asm in System.AppDomain.CurrentDomain.GetAssemblies():
        n = asm.GetName()
        if n.Name.startswith('Microsoft.Web.WebView2'):
            try:
                loc = asm.Location
            except Exception:
                loc = ''
            out.append('%s %s (%s)' % (n.Name, n.Version, loc))
    return '\n'.join(out) or 'no WebView2 assemblies loaded'


def make_environment(user_dir, timeout=30):
    """Creates the WebView2 environment through the Core assembly directly. The WinForms control's own
    CreationProperties path calls a newer Core constructor and fails when Revit has loaded an older Core."""
    from Microsoft.Web.WebView2.Core import CoreWebView2Environment
    task = CoreWebView2Environment.CreateAsync(None, user_dir, None)
    t0 = time.time()
    while not task.IsCompleted:
        Application.DoEvents()          # the task completes through the UI message loop
        time.sleep(0.01)
        if time.time() - t0 > timeout:
            raise Exception('WebView2 environment creation timed out')
    if task.IsFaulted:
        raise task.Exception
    return task.Result


KEEP = '_murray_bulk_bom_window'      # one window at a time; also keeps every object alive after the command returns
# Reading the parts comes first and everything waits for it (exports need it, the 3D view is only extra context), so while it
# runs Revit gives it nearly all of its time. Once the model is read the 3D view is a background task and takes much less,
# so Revit stays comfortable to use. Raise the LOAD numbers for a faster read, lower them for a smoother Revit.
LOAD_SLICE_SECONDS = 0.25             # Revit time per slice while the parts are being read
LOAD_GAP_MS = 15                      # gap between those slices (the next one is also queued at once)
VIEW_SLICE_SECONDS = 0.06             # Revit time per slice while the 3D view is being built
VIEW_GAP_MS = 120                     # gap between those slices: Revit handles clicks and keys in between


def handler_class():
    """The ExternalEvent handler class, made once per Revit session (a .NET type cannot be defined twice)."""
    holder = sys.modules.setdefault(KEEP + '_cls', type(sys)(KEEP + '_cls'))
    if not hasattr(holder, 'PumpHandler'):
        class PumpHandler(IExternalEventHandler):
            __namespace__ = 'MurrayBulkBom'

            def __init__(self, fn):
                self.fn = fn

            def Execute(self, uiapp):
                self.fn()

            def GetName(self):
                return 'Murray Bulk BOM'
        holder.PumpHandler = PumpHandler
    return holder.PumpHandler


def main():
    old = sys.modules.get(KEEP)
    if old is not None:
        try:
            if not old.form.IsDisposed:
                old.form.Activate()          # already open: bring it forward
                return
        except Exception:
            pass

    doc = revit.doc
    uidoc = revit.uidoc
    if doc is None or doc.IsFamilyDocument:
        forms.alert('Open a Revit project first.', exitscript=True)

    if not load_webview2():
        forms.alert(
            'The WebView2 control could not be loaded.\n\n'
            'Copy Microsoft.Web.WebView2.Core.dll, Microsoft.Web.WebView2.WinForms.dll and WebView2Loader.dll '
            'into:\n' + LOCAL_WEBVIEW2 + '\n\nThe WebView2 Runtime must also be installed on this PC.',
            exitscript=True)

    from Microsoft.Web.WebView2.WinForms import WebView2

    pi = doc.ProjectInformation
    svc = Service(revit_read, doc, uidoc,
                  project_name=getattr(pi, 'Name', '') or '',
                  project_number=getattr(pi, 'Number', '') or '')

    # Exports go to the user's Desktop until another folder is picked (the Windows Desktop folder, which is also right when it is
    # redirected to OneDrive).
    try:
        desk = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
    except Exception:
        desk = ''
    svc.default_folder = desk if desk and os.path.isdir(desk) else os.path.join(os.path.expanduser('~'), 'Desktop')

    form = Form()
    form.Text = 'Murray Company | Bulk BOM Export'
    form.StartPosition = FormStartPosition.CenterScreen
    wa = Screen.PrimaryScreen.WorkingArea
    form.Size = Size(min(1100, wa.Width), min(1100, wa.Height))      # 1100 x 1100, or the screen if it is smaller
    form.MinimumSize = Size(900, 560)

    wv = WebView2()
    wv.Dock = DockStyle.Fill
    form.Controls.Add(wv)
    user_dir = os.path.join(tempfile.gettempdir(), 'MurrayBulkBOM_WebView2')   # Revit's folder is not writable
    try:
        env = make_environment(user_dir)
    except Exception as ex:
        forms.alert('WebView2 could not create its environment:\n%s\n\nLoaded:\n%s' % (ex, webview2_versions()),
                    exitscript=True)

    state = {'pending': False, 'ext': None, 'timer': None, 'owner': None, 'blocking': False}

    def reply(obj):
        try:
            wv.CoreWebView2.PostWebMessageAsString(json.dumps(obj))
        except Exception:
            pass

    def pump_once():
        """Runs inside a Revit API context: the ExternalEvent, or the blocking loop in fallback mode."""
        state['pending'] = False
        try:
            reading = not svc.loaded                           # the parts first, the 3D view after
            for m in svc.pump(LOAD_SLICE_SECONDS if reading else VIEW_SLICE_SECONDS):
                reply(m)
            if reading and svc.busy:
                kick()                                         # no gap: queue the next slice of the read at once
        except Exception as ex:
            reply({'type': 'error', 'message': '%s' % ex})
            sys.stderr.write(traceback.format_exc())

    def kick():
        if state['ext'] is not None and svc.busy and not state['pending']:
            state['pending'] = True
            state['ext'].Raise()

    def on_message(sender, args):
        # Window thread, no Revit API context: only svc.handle() (which never calls the API) is allowed here.
        try:
            msg = json.loads(args.TryGetWebMessageAsString())
            cmd = msg.get('cmd')
            if cmd == 'chooseFolder':
                dlg = FolderBrowserDialog()
                dlg.Description = 'Choose the folder for the MR forms'
                start = msg.get('folder') or ''
                if start and os.path.isdir(start):
                    dlg.SelectedPath = start                     # open where the user already is
                reply({'type': 'folder', 'path': dlg.SelectedPath if dlg.ShowDialog() == DialogResult.OK else ''})
                return
            out = svc.handle(msg)
            for m in out:
                reply(m)
                if m.get('type') == 'done' and m.get('files'):
                    try:
                        os.startfile(os.path.dirname(m['files'][0]))
                    except Exception:
                        pass
            kick()
        except Exception as ex:
            reply({'type': 'error', 'message': '%s' % ex})
            sys.stderr.write(traceback.format_exc())

    # Model changes: Revit calls this after every transaction in any open document, so it must stay cheap. It only does
    # bookkeeping (no geometry, no reading) and tells the page that the model changed since it was read.
    app = doc.Application

    def on_doc_changed(sender, args):
        # Edits during a read are recorded too; the tag is shown when the read finishes (see Service._finish).
        m = revit_read.on_document_changed(doc, args, svc)      # never raises, never prints
        if m and svc.loaded:
            reply(m)

    try:
        app.DocumentChanged += on_doc_changed
        svc.tracking = True
    except Exception as ex:
        sys.stderr.write('Model change tracking is not available, Refresh model re-reads everything: %s\n' % ex)

    def on_init_done(sender, args):
        if not args.IsSuccess:
            MessageBox.Show('WebView2 failed to start: %s\n\nLoaded:\n%s' % (args.InitializationException, webview2_versions()))
            return
        wv.CoreWebView2.WebMessageReceived += on_message
        wv.CoreWebView2.Navigate(Uri(UI_PAGE).AbsoluteUri)

    wv.CoreWebView2InitializationCompleted += on_init_done
    wv.EnsureCoreWebView2Async(env)

    # Modeless: the window stays open while Revit is used. Revit API work is done in slices from an ExternalEvent.
    try:
        ext = ExternalEvent.Create(handler_class()(pump_once))
        state['ext'] = ext
    except Exception as ex:
        state['ext'] = None
        sys.stderr.write('ExternalEvent could not be created, using the blocking window instead: %s\n' % ex)

    def on_tick(sender, args):
        sender.Interval = VIEW_GAP_MS if svc.loaded else LOAD_GAP_MS
        kick()

    def on_closed(sender, args):
        try:
            if svc.tracking:
                app.DocumentChanged -= on_doc_changed
            if state['timer'] is not None:
                state['timer'].Stop()
            if state['ext'] is not None:
                state['ext'].Dispose()
            if state['owner'] is not None:
                state['owner'].ReleaseHandle()
        except Exception:
            pass
        sys.modules.pop(KEEP, None)

    form.FormClosed += on_closed
    keep = type(sys)(KEEP)
    keep.form, keep.svc, keep.wv, keep.state, keep.handlers = form, svc, wv, state, (on_message, on_init_done, on_tick, on_closed, pump_once, on_doc_changed)
    sys.modules[KEEP] = keep

    if state['ext'] is not None:
        timer = Timer()
        timer.Interval = LOAD_GAP_MS
        timer.Tick += on_tick
        timer.Start()
        state['timer'] = timer
        owner = NativeWindow()
        owner.AssignHandle(Process.GetCurrentProcess().MainWindowHandle)    # keeps the window above Revit
        state['owner'] = owner
        keep.timer, keep.owner = timer, owner
        form.Show(owner)
    else:
        # Fallback: stay inside this command (so the API is available) and pump between window events. Revit is
        # blocked until the window is closed, but progress, Cancel and slicing all work.
        form.Show()
        while form.Visible:
            Application.DoEvents()
            if svc.busy:
                pump_once()
            else:
                time.sleep(0.02)


main()
