# Family Navigator Pane

## Summary

The Family Navigator is a dockable Revit pane for browsing the shared family library as a thumbnail grid and placing families in the active view. It also includes work plane and section tools, and a utility that generates the thumbnails.

### Quick reference

- **Click a tile:** Loads the family if needed and starts placement. Click in the view to place.
- **ALT + click a tile:** Opens that family's folder in Windows Explorer. Does not place the family.
- **Hover over a tile:** Shows the full file name in the status bar.
- **Search box:** Filters tiles as you type. Every word entered must appear in the name.
- **Browse / Default:** Choose any folder or return to the main FAMILIES folder.
- **Quick folder buttons:** Atkore Strut, Plumbing Fixtures, Mechanical Equipment, Pipe Accessories.
- **Set Work Plane:** 3D or section view only. Pick a part, then choose orientation if asked.
- **Hide Work Plane:** Turns off work plane display in the active view.
- **Create Section:** Plan or section view only. Pick a point on a straight pipe, rod, or line-based Generic Model.
- **Generate Missing Images:** Creates transparent PNG thumbnails for families that have none. Never overwrites.

### Good to know

- The last folder used reopens automatically next time.
- Thumbnails are PNG, JPG, JPEG, or BMP files saved next to the family with the same file name.
- To regenerate a thumbnail, delete its PNG, then run Generate Missing Images again.
    - Cannot remove or update existing images while Revit is running.
- The status bar at the bottom shows instructions, progress, and error messages. Check it first if something does not work.

## 1. Choosing a Folder

- **Families Folder Path:** shows the folder being browsed. Subfolders are included.
- **Browse:** pick any folder.
- **Default:** returns to the main FAMILIES folder on the C: Egnyte BIM share.
- **Quick buttons:** Atkore Strut, Plumbing Fixtures, Mechanical Equipment, Pipe Accessories.
- The last folder used reopens automatically next time.

## 2. Searching

- Type in Search Families to filter the grid as you type.
- Not case sensitive. Multiple words are allowed, and a family must contain every word.
- The X button clears the search.

## 3. The Family Grid

- Each tile shows a thumbnail and a short name (the text before the first underscore in the file name). A tile with no image shows only the name.
- Families with "channel" in the name sort first. Everything else sorts in natural number order (2 before 10).
- **Hover** over a tile to see the full file name in the status bar.
- **Click** a tile to place it:
    - If the family is not in the project, it is loaded first. Families already in the project are not overwritten.
    - Placement starts with the first type in the family. Click in the view to place.
    - Placement is blocked in schedules, sheets, and reports.
- **ALT + click** a tile to open that family's folder in Windows Explorer. This does not place the family.

## 4. Set Work Plane and Hide Work Plane

- Run from a 3D view or section view only.
- Click Set Work Plane, then pick a fabrication part, family instance, or assembly. For an assembly, you are then asked to pick the part inside it.
- Depending on the element and view, you may be asked to choose orientation (vertical or horizontal) or axis (X or Y).
- In a section view, the plane matches the view automatically.
- The plane is shown in the view and a reference plane named TEMPORARY is created. Setting a new plane replaces the previous TEMPORARY reference plane.
- Hide Work Plane turns off the work plane display in the active view.

## 5. Create Section

- Run from a floor plan or section view.
- **Create Element Section**, pick a point on a straight pipe, hanger rod, or line-based Generic Model. Pick on the side where you want the section marker.
- Supported fabrication parts are limited to straight parts with the configured item IDs.
- The section runs along the element with 1 ft of padding on each side and above and below, and 0.5 ft in depth. Detail level is set to Fine.
- **Create Section of Elements**, select elements in view, then pick section marker location.
- All element categories are supported.
- The section runs along elements with 1 ft of padding on each side and above and below, and 0.5 ft in depth. Detail level is set to Fine.

## 6. Generate Missing Images (Admin)

- Scans the current folder and subfolders for families that have no matching image.
- For each one, the family is opened in the background, a 3D view is exported, and the background is made transparent.
- If the export fails, a Revit type preview image is used instead.
- The PNG is saved next to the RFA with the same name. Existing images are never overwritten.
- To regenerate an image, delete its PNG first, then run the tool again.
- Requires write access to the folder. Revit may be unresponsive while a large folder processes. Progress and failures show in the status bar.

## 7. Status Bar

Shows instructions, hover names, progress, and error messages. If something does not work, read this line first.