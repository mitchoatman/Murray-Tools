# -*- coding: utf-8 -*-
import os
import re
import shutil
import clr
import System

clr.AddReference("System")
clr.AddReference("System.Core")
clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")
clr.AddReference("System.Windows.Forms")
clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")

from pyrevit import forms
import Autodesk.Revit.UI as UI

from System.Collections.Generic import List
from System.Windows.Forms import FolderBrowserDialog, DialogResult
from Autodesk.Revit import DB
from Autodesk.Revit.DB import (
    Transaction,
    TransactionGroup,
    FilteredElementCollector,
    ElementId,
    Family,
    FamilySymbol,
    FabricationPart,
    ConnectorType,
    XYZ,
    Plane,
    SketchPlane,
    ViewType,
    ViewFamilyType,
    ViewFamily,
    View3D,
    ViewPlan,
    ViewSection,
    ReferencePlane,
    LocationPoint,
    LocationCurve,
    FamilyInstance,
    AssemblyInstance,
    ImageExportOptions,
    ImageFileType,
    ImageResolution,
    ZoomFitType
)
from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter
from Autodesk.Revit.UI import (
    IExternalEventHandler,
    ExternalEvent,
    TaskDialog,
    TaskDialogCommandLinkId,
    TaskDialogCommonButtons,
    TaskDialogResult
)
from Autodesk.Revit.Exceptions import OperationCanceledException

PANE_XAML = """
<Page
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Background="#FFF5F5F5">

    <Grid Margin="10">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <!-- Folder Path Selection -->
        <StackPanel Grid.Row="0" Margin="0,0,0,8">
            <TextBlock Text="Families Folder Path:" FontWeight="SemiBold" Margin="0,0,0,2"/>
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="70"/>
                    <ColumnDefinition Width="70"/>
                </Grid.ColumnDefinitions>
                <TextBox x:Name="folder_path_tb" Grid.Column="0" Height="24" VerticalContentAlignment="Center"/>
                <Button x:Name="browse_btn" Grid.Column="1" Content="Browse" Height="24" Margin="4,0,0,0"/>
                <Button x:Name="default_path_btn" Grid.Column="2" Content="Default" Height="24" Margin="4,0,0,0"/>
            </Grid>
        </StackPanel>

        <!-- Search Filter -->
        <StackPanel Grid.Row="1" Margin="0,0,0,8">
            <TextBlock Text="Search Families:" FontWeight="SemiBold" Margin="0,0,0,2"/>
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="24"/>
                </Grid.ColumnDefinitions>
                <TextBox x:Name="search_tb" Grid.Column="0" Height="24" VerticalContentAlignment="Center"/>
                <Button x:Name="clear_search_btn" Grid.Column="1" Content="X" Height="24" Margin="4,0,0,0"/>
            </Grid>
        </StackPanel>

        <!-- Work Plane Tools Section -->
        <Grid Grid.Row="2" Margin="0,0,0,8">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="4"/>
                <ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>
            <Button x:Name="set_workplane_btn" Grid.Column="0" Content="Set Work Plane" Height="26"/>
            <Button x:Name="hide_workplane_btn" Grid.Column="2" Content="Hide Work Plane" Height="26"/>
        </Grid>

        <!-- Section Tools Section -->
        <Grid Grid.Row="3" Margin="0,0,0,8">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="4"/>
                <ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>
            <Button x:Name="create_section_btn" Grid.Column="0" Content="Create Section" Height="26"/>
        </Grid>

        <TextBlock Grid.Row="4"
                   Text="Single Click to Place in View"
                   FontWeight="SemiBold"
                   Margin="0,0,0,4"/>

        <!-- Image Grid ListBox -->
        <ScrollViewer Grid.Row="5"
                      VerticalScrollBarVisibility="Auto"
                      HorizontalScrollBarVisibility="Disabled"
                      BorderBrush="#FFD0D0D0"
                      BorderThickness="1"
                      Background="White">
            <ListBox x:Name="families_lb"
                     BorderThickness="0"
                     ScrollViewer.HorizontalScrollBarVisibility="Disabled">
                <ListBox.ItemsPanel>
                    <ItemsPanelTemplate>
                        <WrapPanel IsItemsHost="True" Orientation="Horizontal"/>
                    </ItemsPanelTemplate>
                </ListBox.ItemsPanel>
                <ListBox.ItemTemplate>
                    <DataTemplate>
                        <Border Width="95" Height="110" Margin="4" Padding="4" BorderBrush="#FFDDDDDD" BorderThickness="1" Background="White" CornerRadius="3">
                            <StackPanel Orientation="Vertical" HorizontalAlignment="Center">
                                <Image Source="{Binding ImagePath}" Width="75" Height="75" Stretch="Uniform"/>
                                <TextBlock Text="{Binding DisplayName}" TextTrimming="CharacterEllipsis" TextAlignment="Center" FontSize="10" Margin="0,4,0,0" Width="85"/>
                            </StackPanel>
                        </Border>
                    </DataTemplate>
                </ListBox.ItemTemplate>
            </ListBox>
        </ScrollViewer>

        <!-- Generate Button at Bottom -->
        <Button x:Name="generate_images_btn"
                Grid.Row="6"
                Content="Generate Missing Images"
                Height="28"
                Margin="0,8,0,0"/>

        <TextBlock x:Name="status_tb"
                   Grid.Row="7"
                   Margin="0,6,0,0"
                   Foreground="#666666"
                   TextWrapping="Wrap"
                   Text="Ready."/>
    </Grid>
</Page>
"""

DEFAULT_FAMILY_FOLDER = r"C:\Egnyte\Shared\BIM\Murray CADetailing Dept\REVIT\FAMILIES\Generic Models\Atkore"

# Allowed straight fabrication CIDs
ALLOWED_CIDS = set([
    2041, 866, 40,
])

state = UI.DockablePaneState()
state.DockPosition = UI.DockPosition.Right


class FamilyItem(object):
    def __init__(self, name, rfa_path, image_path):
        self.Name = name
        self.DisplayName = name
        self.RfaPath = rfa_path
        if image_path and os.path.exists(image_path):
            self.ImagePath = image_path
        else:
            self.ImagePath = None


def natural_sort_key(text):
    parts = re.split(r'(\d+)', text or "")
    key = []
    for part in parts:
        if part.isdigit():
            key.append((0, int(part)))
        else:
            key.append((1, part.lower()))
    return key


def family_sort_key(item):
    name = item.Name or ""
    channel_first = 0 if "channel" in name.lower() else 1
    return (channel_first, natural_sort_key(name))


class FamilyLoaderOptionsHandler(System.Object, System.IDisposable, DB.IFamilyLoadOptions):
    def OnFamilyFound(self, familyInUse, overwriteParameterValues):
        overwriteParameterValues.Value = False
        return True

    def OnSharedFamilyFound(self, sharedFamily, familyInUse, source, overwriteParameterValues):
        source.Value = DB.FamilySource.Family
        overwriteParameterValues.Value = False
        return True

    def Dispose(self):
        pass


class DialogSuppressor(object):
    def __init__(self):
        self.active = False

    def handler(self, sender, args):
        if not self.active:
            return
        if args.DialogId == "TaskDialog_Opening_File_From_Later_Version":
            args.OverrideResult(1)


class TopLevelSelectionFilter(ISelectionFilter):
    def AllowElement(self, elem):
        return (
            isinstance(elem, FabricationPart) or
            isinstance(elem, FamilyInstance) or
            isinstance(elem, AssemblyInstance)
        )

    def AllowReference(self, reference, position):
        return True


class AssemblyMemberSelectionFilter(ISelectionFilter):
    def __init__(self, allowed_ids):
        self.allowed_ids = set([eid.IntegerValue for eid in allowed_ids])

    def AllowElement(self, elem):
        if elem.Id.IntegerValue not in self.allowed_ids:
            return False
        return isinstance(elem, FabricationPart) or isinstance(elem, FamilyInstance)

    def AllowReference(self, reference, position):
        return True


class StraightFabOrLineBasedGenericFilter(ISelectionFilter):
    def __init__(self, doc):
        self.doc = doc

    def _is_straight_fab(self, element):
        try:
            if not isinstance(element, FabricationPart):
                return False

            cid = element.ItemCustomId
            if cid not in ALLOWED_CIDS:
                return False

            loc = element.Location
            if not isinstance(loc, LocationCurve):
                return False

            if not isinstance(loc.Curve, DB.Line):
                return False

            return True
        except:
            return False

    def _is_line_based_generic(self, element):
        try:
            if not isinstance(element, FamilyInstance):
                return False

            if not element.Category:
                return False

            if element.Category.Id.IntegerValue != int(DB.BuiltInCategory.OST_GenericModel):
                return False

            fam = element.Symbol.Family
            if fam.FamilyPlacementType != DB.FamilyPlacementType.CurveBased:
                return False

            loc = element.Location
            if not isinstance(loc, LocationCurve):
                return False

            if not isinstance(loc.Curve, DB.Line):
                return False

            return True
        except:
            return False

    def is_allowed(self, element):
        return self._is_straight_fab(element) or self._is_line_based_generic(element)

    def AllowElement(self, element):
        return self.is_allowed(element)

    def AllowReference(self, reference, point):
        try:
            element = self.doc.GetElement(reference.ElementId)
            return self.is_allowed(element)
        except:
            return False


def midpoint(p1, p2):
    return XYZ((p1.X + p2.X) / 2.0, (p1.Y + p2.Y) / 2.0, (p1.Z + p2.Z) / 2.0)


def is_mostly_vertical(vec):
    return abs(vec.Z) > max(abs(vec.X), abs(vec.Y))


def get_horizontal_direction(vec):
    horiz = XYZ(vec.X, vec.Y, 0.0)
    if horiz.GetLength() == 0:
        return None
    return horiz.Normalize()


class _FamilyViewerRequestHandler(IExternalEventHandler):

    def _create_temporary_project_doc(self, app):
        try:
            default_template = app.DefaultProjectTemplate
            if default_template and os.path.exists(default_template):
                return app.NewProjectDocument(default_template)
        except:
            pass
        return None

    def __init__(self, pane_instance):
        self.request_name = None
        self.selected_family = None
        self.pane = pane_instance
        self.dialog_suppressor = DialogSuppressor()
        self.uiapp = None

    def Execute(self, uiapp):
        self.uiapp = uiapp
        if self.pane is None:
            return

        uidoc = uiapp.ActiveUIDocument
        if uidoc is None:
            self.pane.set_status("No active document.")
            return

        doc = uidoc.Document
        view = doc.ActiveView

        try:
            if self.request_name == "place_family":
                if view and view.ViewType in [ViewType.Schedule, ViewType.DrawingSheet, ViewType.Report]:
                    self.pane.set_status("Cannot place elements in schedules, sheets, or reports.")
                    return
                self._do_place_family(uidoc, doc, self.selected_family)

            elif self.request_name == "set_work_plane":
                if view.ViewType not in (ViewType.ThreeD, ViewType.Section):
                    self.pane.set_status("Run Work Plane tool from a 3D view or Section view.")
                    return
                self._do_set_work_plane(doc, uidoc, view)

            elif self.request_name == "hide_work_plane":
                self._do_hide_work_plane(doc, view)

            elif self.request_name == "create_section":
                self._do_create_section(doc, uidoc, view)

            elif self.request_name == "generate_images":
                self._do_generate_images(uiapp, self.pane.folder_path_tb.Text)

        except OperationCanceledException:
            self.pane.set_status("Operation cancelled.")
        except Exception as ex:
            self.pane.set_status("Error: {}".format(str(ex)))
        finally:
            self.request_name = None
            self.selected_family = None

    def GetName(self):
        return "Family Viewer Action Handler"

    def _do_place_family(self, uidoc, doc, fam_item):
        if not fam_item or not os.path.exists(fam_item.RfaPath):
            self.pane.set_status("Family file path invalid.")
            return

        family = None
        for fam in FilteredElementCollector(doc).OfClass(Family):
            if fam.Name == fam_item.Name:
                family = fam
                break

        if not family:
            with Transaction(doc, "Load Family: {}".format(fam_item.Name)) as t:
                t.Start()
                fload_handler = FamilyLoaderOptionsHandler()
                doc.LoadFamily(fam_item.RfaPath, fload_handler)
                t.Commit()

            for fam in FilteredElementCollector(doc).OfClass(Family):
                if fam.Name == fam_item.Name:
                    family = fam
                    break

        if not family:
            self.pane.set_status("Failed to load family '{}'.".format(fam_item.Name))
            return

        symbol_ids = family.GetFamilySymbolIds()
        if not symbol_ids:
            self.pane.set_status("No types found in family '{}'.".format(fam_item.Name))
            return

        first_symbol_id = next(iter(symbol_ids), None)
        if not first_symbol_id:
            self.pane.set_status("No valid symbol ID found.")
            return

        symbol = doc.GetElement(first_symbol_id)
        if not symbol:
            self.pane.set_status("Could not retrieve family symbol.")
            return

        if not symbol.IsActive:
            with Transaction(doc, "Activate Symbol") as t:
                t.Start()
                symbol.Activate()
                t.Commit()

        try:
            uidoc.PostRequestForElementTypePlacement(symbol)
            self.pane.set_status("Ready to place '{}'. Click in model view.".format(fam_item.Name))
        except Exception as e:
            self.pane.set_status("Placement error: {}".format(str(e)))

    def _do_generate_images(self, uiapp, folder_path):
        if not os.path.exists(folder_path):
            self.pane.set_status("Folder path does not exist.")
            return

        rfa_files = []
        for root, dirs, files in os.walk(folder_path):
            for file in files:
                if file.lower().endswith(".rfa"):
                    rfa_files.append(os.path.join(root, file))

        if not rfa_files:
            self.pane.set_status("No RFA files found in folder.")
            return

        app = uiapp.Application
        processed = 0
        failed = 0
        fail_msgs = []

        self.dialog_suppressor.active = True
        uiapp.DialogBoxShowing += self.dialog_suppressor.handler

        try:
            for idx, file_path in enumerate(rfa_files):
                rfa_name = os.path.splitext(os.path.basename(file_path))[0]
                img_path = os.path.join(os.path.dirname(file_path), rfa_name + ".png")

                if os.path.exists(img_path):
                    continue

                self.pane.set_status("Generating preview [{}/{}]: {}...".format(idx + 1, len(rfa_files), rfa_name))
                fam_doc = None

                try:
                    open_opts = DB.OpenOptions()
                    open_opts.Audit = False

                    model_path = DB.ModelPathUtils.ConvertUserVisiblePathToModelPath(file_path)
                    fam_doc = app.OpenDocumentFile(model_path, open_opts)

                    result = self._export_family_preview(fam_doc, app, file_path, img_path)
                    if result is True:
                        processed += 1
                    else:
                        failed += 1
                        fail_msgs.append("{}: {}".format(rfa_name, result or "Unknown export failure"))

                except Exception as ex:
                    failed += 1
                    fail_msgs.append("{}: {}".format(rfa_name, str(ex)))
                    print("Preview generation failed for {}: {}".format(file_path, ex))

                finally:
                    if fam_doc:
                        try:
                            fam_doc.Close(False)
                        except:
                            pass

        finally:
            self.dialog_suppressor.active = False
            try:
                uiapp.DialogBoxShowing -= self.dialog_suppressor.handler
            except:
                pass

        self.pane.load_families_from_folder(folder_path)

        if fail_msgs:
            self.pane.set_status(
                "Generated {} image(s). Failed: {}. First error: {}".format(
                    processed,
                    failed,
                    fail_msgs[0]
                )
            )
        else:
            self.pane.set_status("Generated {} image(s). Failed: {}.".format(processed, failed))

    def _export_family_preview(self, fam_doc, app, rfa_path, img_path):
        temp_root = os.environ.get("TEMP") or os.environ.get("TMP") or r"C:\Temp"
        temp_dir = os.path.join(temp_root, "FamilyPreviewTemp")

        if not os.path.exists(temp_dir):
            os.makedirs(temp_dir)

        for f in os.listdir(temp_dir):
            if f.lower().endswith(".png"):
                try:
                    os.remove(os.path.join(temp_dir, f))
                except:
                    pass

        export_err = None

        view3d = None
        try:
            for v in FilteredElementCollector(fam_doc).OfClass(View3D):
                if not v.IsTemplate:
                    view3d = v
                    break
        except Exception:
            view3d = None

        if view3d is None:
            try:
                vft = None
                for x in FilteredElementCollector(fam_doc).OfClass(ViewFamilyType):
                    if x.ViewFamily == ViewFamily.ThreeDimensional:
                        vft = x
                        break

                if vft is not None:
                    t = Transaction(fam_doc, "Create Preview 3D View")
                    t.Start()
                    view3d = View3D.CreateIsometric(fam_doc, vft.Id)
                    fam_doc.Regenerate()
                    t.Commit()
            except Exception as ex:
                export_err = "3D view creation failed: {}".format(str(ex))
                view3d = None

        if view3d is not None:
            try:
                opts = ImageExportOptions()
                opts.ExportRange = DB.ExportRange.SetOfViews
                opts.FilePath = os.path.join(temp_dir, "preview")
                opts.HLRandWFViewsFileType = ImageFileType.PNG
                opts.ImageResolution = ImageResolution.DPI_150
                opts.ZoomType = ZoomFitType.FitToPage
                opts.PixelSize = 512

                view_ids = List[ElementId]()
                view_ids.Add(view3d.Id)
                opts.SetViewsAndSheets(view_ids)

                fam_doc.ExportImage(opts)

                pngs = [os.path.join(temp_dir, f) for f in os.listdir(temp_dir) if f.lower().endswith(".png")]
                if pngs:
                    exported = max(pngs, key=os.path.getmtime)
                    shutil.copyfile(exported, img_path)
                    return True
                else:
                    export_err = "ExportImage created no PNG"

            except Exception as ex:
                export_err = "ExportImage failed: {}".format(str(ex))
        elif export_err is None:
            export_err = "No usable 3D view found"

        try:
            clr.AddReference("System.Drawing")
            import System.Drawing as Drawing

            temp_proj = self._create_temporary_project_doc(app)
            if temp_proj is None:
                return "{}; could not create temporary project document".format(export_err)

            try:
                loaded_family = None

                t = Transaction(temp_proj, "Temp Load Family For Preview")
                t.Start()
                fload_handler = FamilyLoaderOptionsHandler()
                load_ok = temp_proj.LoadFamily(rfa_path, fload_handler)
                t.Commit()

                if not load_ok:
                    return "{}; temporary project failed to load family".format(export_err)

                fam_name = os.path.splitext(os.path.basename(rfa_path))[0]
                for fam in FilteredElementCollector(temp_proj).OfClass(Family):
                    if fam.Name == fam_name:
                        loaded_family = fam
                        break

                if loaded_family is None:
                    return "{}; temp project could not find loaded family".format(export_err)

                symbol_ids = loaded_family.GetFamilySymbolIds()
                if not symbol_ids:
                    return "{}; temp project found family but no symbols".format(export_err)

                first_symbol = None
                for sid in symbol_ids:
                    first_symbol = temp_proj.GetElement(sid)
                    if first_symbol is not None:
                        break

                if first_symbol is None:
                    return "{}; temp project symbol was None".format(export_err)

                bmp = first_symbol.GetPreviewImage(Drawing.Size(256, 256))
                if bmp is None:
                    return "{}; temp project GetPreviewImage returned None".format(export_err)

                bmp.Save(img_path, Drawing.Imaging.ImageFormat.Png)
                return True

            finally:
                try:
                    temp_proj.Close(False)
                except:
                    pass

        except Exception as ex:
            return "{}; temp project fallback failed: {}".format(export_err, str(ex))

    def _do_set_work_plane(self, doc, uidoc, curview):
        ref = uidoc.Selection.PickObject(
            ObjectType.Element,
            TopLevelSelectionFilter(),
            "Select a fabrication part, family instance, or assembly"
        )
        elem = doc.GetElement(ref.ElementId)

        if isinstance(elem, AssemblyInstance):
            member_ids = list(elem.GetMemberIds())
            valid_member_ids = [mid for mid in member_ids if isinstance(doc.GetElement(mid), (FabricationPart, FamilyInstance))]
            if not valid_member_ids:
                raise Exception("The selected assembly has no valid members.")
            if len(valid_member_ids) == 1:
                elem = doc.GetElement(valid_member_ids[0])
            else:
                member_ref = uidoc.Selection.PickObject(
                    ObjectType.Element,
                    AssemblyMemberSelectionFilter(valid_member_ids),
                    "Pick part inside assembly"
                )
                elem = doc.GetElement(member_ref.ElementId)

        plane_origin, x_vector, y_vector = None, None, None

        if isinstance(elem, FabricationPart):
            connectors = [conn for conn in elem.ConnectorManager.Connectors if conn.ConnectorType == ConnectorType.End]
            if len(connectors) < 2:
                raise Exception("Fabrication part needs at least two end connectors.")
            p1, p2 = connectors[0].Origin, connectors[1].Origin
            vec = p2 - p1
            if vec.GetLength() == 0:
                raise Exception("Invalid connector geometry.")
            plane_origin, line_dir = midpoint(p1, p2), vec.Normalize()

            if curview.ViewType == ViewType.Section:
                x_vector, y_vector = curview.RightDirection.Normalize(), curview.UpDirection.Normalize()
            elif is_mostly_vertical(line_dir):
                axis = self._choose_dialog("Select Axis", "Vertical part. Choose alignment axis:", "X-axis", "Y-axis")
                if not axis:
                    return
                x_vector, y_vector = (XYZ.BasisX, line_dir) if axis == 1 else (line_dir, XYZ.BasisY)
            else:
                horiz = get_horizontal_direction(line_dir)
                orientation = self._choose_dialog("Select Orientation", "Choose work plane orientation:", "Vertical", "Horizontal")
                if not orientation:
                    return
                if orientation == 1:
                    x_vector, y_vector = horiz, XYZ.BasisZ
                else:
                    perp = XYZ.BasisZ.CrossProduct(horiz)
                    x_vector, y_vector = horiz, perp.Normalize()

        elif isinstance(elem, FamilyInstance):
            loc = elem.Location
            if isinstance(loc, LocationPoint):
                plane_origin = loc.Point
                if curview.ViewType == ViewType.Section:
                    x_vector, y_vector = curview.RightDirection.Normalize(), curview.UpDirection.Normalize()
                else:
                    orientation = self._choose_dialog("Select Orientation", "Choose orientation:", "Horizontal", "Vertical")
                    if not orientation:
                        return
                    if orientation == 1:
                        x_vector, y_vector = XYZ.BasisX, XYZ.BasisY
                    else:
                        axis = self._choose_dialog("Select Axis", "Choose vertical axis:", "X-axis", "Y-axis")
                        if not axis:
                            return
                        x_vector, y_vector = (XYZ.BasisX, XYZ.BasisZ) if axis == 1 else (XYZ.BasisY, XYZ.BasisZ)
            elif isinstance(loc, LocationCurve):
                curve = loc.Curve
                vec = curve.GetEndPoint(1) - curve.GetEndPoint(0)
                plane_origin, line_dir = curve.Evaluate(0.5, True), vec.Normalize()
                if curview.ViewType == ViewType.Section:
                    x_vector, y_vector = curview.RightDirection.Normalize(), curview.UpDirection.Normalize()
                else:
                    horiz = get_horizontal_direction(line_dir)
                    x_vector, y_vector = horiz, XYZ.BasisZ

        plane = Plane.CreateByOriginAndBasis(plane_origin, x_vector, y_vector)

        with TransactionGroup(doc, "Create Work Plane") as tg:
            tg.Start()
            with Transaction(doc, "Set Work Plane") as t1:
                t1.Start()
                sp = SketchPlane.Create(doc, plane)
                try:
                    sp.Name = "TEMPORARY"
                except:
                    pass
                curview.SketchPlane = sp
                curview.ShowActiveWorkPlane()
                t1.Commit()
            with Transaction(doc, "Clean Reference Planes") as t2:
                t2.Start()

                temp_refplane_ids = [
                    rp.Id for rp in FilteredElementCollector(doc)
                    .OfClass(ReferencePlane)
                    .WhereElementIsNotElementType()
                    if rp.Name == "TEMPORARY"
                ]

                for rp_id in temp_refplane_ids:
                    doc.Delete(rp_id)

                ref_plane = doc.Create.NewReferencePlane(
                    plane.Origin,
                    plane.Origin + plane.XVec,
                    plane.YVec,
                    curview
                )
                ref_plane.Name = "TEMPORARY"

                t2.Commit()
            tg.Assimilate()

        self.pane.set_status("Work plane set successfully.")

    def _do_hide_work_plane(self, doc, curview):
        with Transaction(doc, "Disable Workplane Display") as t:
            t.Start()
            curview.HideActiveWorkPlane()
            t.Commit()
        self.pane.set_status("Work plane hidden.")

    def _get_section_view_type(self, document):
        for vft in DB.FilteredElementCollector(document).OfClass(ViewFamilyType):
            if vft.ViewFamily == ViewFamily.Section:
                return vft
        return None

    def _require_supported_section_view(self, view):
        if not isinstance(view, ViewPlan) and not isinstance(view, ViewSection):
            raise Exception("Run this command from a floor plan or section view.")

    def _pick_element_and_point_for_section(self, doc, uidoc):
        ref = uidoc.Selection.PickObject(
            ObjectType.PointOnElement,
            StraightFabOrLineBasedGenericFilter(doc),
            "Pick a straight fabrication part or line-based Generic Model on the side where you want the section marker"
        )
        elem = doc.GetElement(ref.ElementId)
        pt = ref.GlobalPoint
        return elem, pt

    def _get_linear_curve(self, elem):
        loc = elem.Location
        if isinstance(loc, LocationCurve):
            crv = loc.Curve
            if isinstance(crv, DB.Line):
                return crv
        return None

    def _bbox_corners(self, bbox):
        mn = bbox.Min
        mx = bbox.Max
        pts = []
        for x in [mn.X, mx.X]:
            for y in [mn.Y, mx.Y]:
                for z in [mn.Z, mx.Z]:
                    pts.append(XYZ(x, y, z))
        return pts

    def _get_points_for_bounds(self, elem, curve):
        pts = [curve.GetEndPoint(0), curve.GetEndPoint(1)]
        bbox = elem.get_BoundingBox(None)
        if bbox:
            pts.extend(self._bbox_corners(bbox))
        return pts

    # PLAN LOGIC: kept to match your working behavior
    def _build_section_transform_from_plan(self, curve, view, pick_point):
        p0 = curve.GetEndPoint(0)
        p1 = curve.GetEndPoint(1)
        origin = curve.Evaluate(0.5, True)

        line_dir = (p1 - p0)
        if line_dir.GetLength() < 1e-9:
            raise Exception("Selected element is too short.")

        line_dir = line_dir.Normalize()

        plan_view_dir = view.ViewDirection.Normalize()

        x_guess = line_dir - plan_view_dir.Multiply(line_dir.DotProduct(plan_view_dir))
        if x_guess.GetLength() < 1e-6:
            raise Exception("Selected element is vertical or nearly vertical in this plan.")

        x_guess = x_guess.Normalize()
        y_axis = XYZ.BasisZ

        z_guess = x_guess.CrossProduct(y_axis)
        if z_guess.GetLength() < 1e-6:
            raise Exception("Could not determine section direction.")
        z_guess = z_guess.Normalize()

        proj = curve.Project(pick_point)
        if proj:
            on_curve = proj.XYZPoint
        else:
            on_curve = origin

        side_vec = pick_point - on_curve
        side_vec = side_vec - y_axis.Multiply(side_vec.DotProduct(y_axis))

        # keep plan behavior exactly as your working version
        if side_vec.GetLength() > 1e-6 and side_vec.DotProduct(z_guess) > 0:
            z_axis = z_guess.Negate()
        else:
            z_axis = z_guess

        x_axis = y_axis.CrossProduct(z_axis).Normalize()

        tf = DB.Transform.Identity
        tf.Origin = origin
        tf.BasisX = x_axis
        tf.BasisY = y_axis
        tf.BasisZ = z_axis

        return tf

    # SECTION LOGIC: separate branch for active section view
    def _build_section_transform_from_section(self, curve, view, pick_point):
        p0 = curve.GetEndPoint(0)
        p1 = curve.GetEndPoint(1)
        origin = curve.Evaluate(0.5, True)

        line_dir = p1 - p0
        if line_dir.GetLength() < 1e-9:
            raise Exception("Selected element is too short.")
        line_dir = line_dir.Normalize()

        view_dir = view.ViewDirection.Normalize()

        x_guess = line_dir - view_dir.Multiply(line_dir.DotProduct(view_dir))
        if x_guess.GetLength() < 1e-6:
            raise Exception("Selected element is perpendicular to the active section view.")
        x_guess = x_guess.Normalize()

        y_guess = view.UpDirection.Normalize()

        z_guess = x_guess.CrossProduct(y_guess)
        if z_guess.GetLength() < 1e-6:
            raise Exception("Could not determine section direction.")
        z_guess = z_guess.Normalize()

        proj = curve.Project(pick_point)
        if proj:
            on_curve = proj.XYZPoint
        else:
            on_curve = origin

        side_vec = pick_point - on_curve
        side_vec = side_vec - view_dir.Multiply(side_vec.DotProduct(view_dir))

        if side_vec.GetLength() > 1e-6 and side_vec.DotProduct(z_guess) > 0:
            z_axis = z_guess.Negate()
        else:
            z_axis = z_guess

        x_axis = y_guess.CrossProduct(z_axis).Normalize()
        y_axis = z_axis.CrossProduct(x_axis).Normalize()

        tf = DB.Transform.Identity
        tf.Origin = origin
        tf.BasisX = x_axis
        tf.BasisY = y_axis
        tf.BasisZ = z_axis

        return tf

    def _get_local_bounds(self, points, transform):
        inv = transform.Inverse
        xs, ys, zs = [], [], []

        for p in points:
            lp = inv.OfPoint(p)
            xs.append(lp.X)
            ys.append(lp.Y)
            zs.append(lp.Z)

        return XYZ(min(xs), min(ys), min(zs)), XYZ(max(xs), max(ys), max(zs))

    def _do_create_section(self, doc, uidoc, active_view):
        self._require_supported_section_view(active_view)

        elem, pick_point = self._pick_element_and_point_for_section(doc, uidoc)
        curve = self._get_linear_curve(elem)
        if not curve:
            raise Exception("Selected element must have a straight line location curve.")

        section_type = self._get_section_view_type(doc)
        if not section_type:
            raise Exception("No Section ViewFamilyType found.")

        if isinstance(active_view, ViewPlan):
            tf = self._build_section_transform_from_plan(curve, active_view, pick_point)
        elif isinstance(active_view, ViewSection):
            tf = self._build_section_transform_from_section(curve, active_view, pick_point)
        else:
            raise Exception("Run this command from a floor plan or section view.")

        pts = self._get_points_for_bounds(elem, curve)
        local_min, local_max = self._get_local_bounds(pts, tf)

        pad_x = 1.0
        pad_y = 1.0
        pad_z = 0.5

        min_z = local_min.Z - pad_z
        max_z = local_max.Z + pad_z
        if (max_z - min_z) < 0.25:
            max_z = min_z + 0.25

        box = DB.BoundingBoxXYZ()
        box.Transform = tf
        box.Min = XYZ(local_min.X - pad_x, local_min.Y - pad_y, min_z)
        box.Max = XYZ(local_max.X + pad_x, local_max.Y + pad_y, max_z)

        with Transaction(doc, "Create Section Along Element") as t:
            t.Start()
            section_view = DB.ViewSection.CreateSection(doc, section_type.Id, box)
            section_view.DetailLevel = DB.ViewDetailLevel.Fine
            t.Commit()

        self.pane.set_status("Section created successfully.")

    def _choose_dialog(self, title, instruction, opt1, opt2):
        td = TaskDialog(title)
        td.MainInstruction = instruction
        td.AddCommandLink(TaskDialogCommandLinkId.CommandLink1, opt1)
        td.AddCommandLink(TaskDialogCommandLinkId.CommandLink2, opt2)
        td.CommonButtons = TaskDialogCommonButtons.Cancel
        res = td.Show()
        if res == TaskDialogResult.CommandLink1:
            return 1
        if res == TaskDialogResult.CommandLink2:
            return 2
        return None


py_state = UI.DockablePaneState()
py_state.DockPosition = UI.DockPosition.Right


class FamilyViewerPane(forms.WPFPanel):
    panel_id = "C3F8A2D1-4E6B-4182-9B7C-123456789ABC"
    panel_title = "Family Navigator"
    panel_source = "inline"
    initial_state = py_state

    def __init__(self):
        self.load_xaml(PANE_XAML, literal_string=True)

        self._handler = _FamilyViewerRequestHandler(self)
        self._ext_event = ExternalEvent.Create(self._handler)

        self._all_families = []
        self._is_updating_selection = False

        self.folder_path_tb.Text = DEFAULT_FAMILY_FOLDER

        self.browse_btn.Click += self.on_browse_clicked
        self.default_path_btn.Click += self.on_default_path_clicked
        self.clear_search_btn.Click += lambda s, e: setattr(self.search_tb, 'Text', '')
        self.search_tb.TextChanged += self.on_search_changed
        self.families_lb.SelectionChanged += self.on_family_selected

        self.set_workplane_btn.Click += self.on_set_workplane_clicked
        self.hide_workplane_btn.Click += self.on_hide_workplane_clicked
        self.create_section_btn.Click += self.on_create_section_clicked
        self.generate_images_btn.Click += self.on_generate_images_clicked

        self.Loaded += self.on_loaded

    def on_loaded(self, sender, args):
        self.load_families_from_folder(self.folder_path_tb.Text)

    def on_browse_clicked(self, sender, args):
        dlg = FolderBrowserDialog()
        current_path = (self.folder_path_tb.Text or "").strip()

        if current_path and os.path.exists(current_path):
            dlg.SelectedPath = current_path
        elif os.path.exists(DEFAULT_FAMILY_FOLDER):
            dlg.SelectedPath = DEFAULT_FAMILY_FOLDER

        if dlg.ShowDialog() == DialogResult.OK:
            selected_folder = dlg.SelectedPath
            if selected_folder:
                self.folder_path_tb.Text = selected_folder
                self.load_families_from_folder(selected_folder)

    def on_default_path_clicked(self, sender, args):
        if os.path.exists(DEFAULT_FAMILY_FOLDER):
            self.folder_path_tb.Text = DEFAULT_FAMILY_FOLDER
            self.load_families_from_folder(DEFAULT_FAMILY_FOLDER)
        else:
            self.set_status("Default folder path does not exist.")

    def load_families_from_folder(self, folder_path):
        self._is_updating_selection = True
        self._all_families = []
        if not os.path.exists(folder_path):
            self.set_status("Folder path does not exist: {}".format(folder_path))
            self.apply_search_filter()
            self._is_updating_selection = False
            return

        try:
            for root, dirs, files in os.walk(folder_path):
                for file in files:
                    if file.lower().endswith(".rfa"):
                        rfa_name = os.path.splitext(file)[0]
                        rfa_path = os.path.join(root, file)

                        image_path = None
                        for ext in [".png", ".jpg", ".jpeg", ".bmp"]:
                            potential_img = os.path.join(root, rfa_name + ext)
                            if os.path.exists(potential_img):
                                image_path = potential_img
                                break

                        self._all_families.append(FamilyItem(rfa_name, rfa_path, image_path))

            self._all_families = sorted(self._all_families, key=family_sort_key)
            self.apply_search_filter()
            self.set_status("Loaded {} family/families from folder.".format(len(self._all_families)))
        except Exception as ex:
            self.set_status("Error loading folder: {}".format(str(ex)))
        finally:
            self._is_updating_selection = False

    def on_search_changed(self, sender, args):
        self.apply_search_filter()

    def apply_search_filter(self):
        self._is_updating_selection = True
        search_text = (self.search_tb.Text or "").strip().lower()

        if search_text:
            terms = [t for t in re.split(r"\s+", search_text) if t]

            filtered = []
            for x in self._all_families:
                name = (x.Name or "").lower()
                if all(term in name for term in terms):
                    filtered.append(x)
        else:
            filtered = list(self._all_families)

        filtered = sorted(filtered, key=family_sort_key)
        self.families_lb.ItemsSource = filtered
        self.families_lb.SelectedItem = None
        self._is_updating_selection = False

    def on_family_selected(self, sender, args):
        if self._is_updating_selection:
            return

        selected = self.families_lb.SelectedItem
        if not selected:
            return

        self.set_status("Triggering placement for: {}".format(selected.Name))
        self._handler.request_name = "place_family"
        self._handler.selected_family = selected
        self._ext_event.Raise()

        self._is_updating_selection = True
        self.families_lb.SelectedItem = None
        self._is_updating_selection = False

    def on_set_workplane_clicked(self, sender, args):
        self._handler.request_name = "set_work_plane"
        self._ext_event.Raise()

    def on_hide_workplane_clicked(self, sender, args):
        self._handler.request_name = "hide_work_plane"
        self._ext_event.Raise()

    def on_create_section_clicked(self, sender, args):
        self.set_status("Pick a straight fabrication part or line-based generic model to create a section.")
        self._handler.request_name = "create_section"
        self._ext_event.Raise()

    def on_generate_images_clicked(self, sender, args):
        self.set_status("Starting image generation...")
        self._handler.request_name = "generate_images"
        self._ext_event.Raise()

    def set_status(self, message):
        self.status_tb.Text = message or ""