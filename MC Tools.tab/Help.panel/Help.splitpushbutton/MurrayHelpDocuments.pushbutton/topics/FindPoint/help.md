# TS Point Number Locator

Find, zoom to, and edit Total Station (TS) points in the active Revit view.

## What it does

Lists every element in the active view that has a value in the `TS_Point_Number` parameter, along with its `TS_Point_Description`. Pick a row to select the element in the model and zoom to it. Edit the description right in the list and it is written back to the element.

## How to use

1. Open the view that contains your TS points (plan, section, 3D, etc.).
2. Click the **TS Point Number Locator** button. The window opens with a table of Point Number and Description.
3. Type in the **Find Point Data** box to filter the list.
4. Click a row. The element is selected in Revit and the view zooms to it.
5. To change a description, double-click the Description cell, type the new text, and press Enter or click away.
6. Click **Close** or press **Esc** when done.

## Searching

- The search matches against both Point Number and Description.
- It matches partial text anywhere in the value, and ignores upper/lower case. Example: `120` finds `P-1205` and `Anchor 120 east`.
- The list filters as you type.
- Your last search text is saved when you close the window and is restored the next time you open it. Clear the box to see all points.

## Editing descriptions

- Only the **Description** column can be edited. **Point Number** is read-only.
- Each edit is saved to Revit immediately as its own transaction named **Update Point Description**, so you can undo it with Ctrl+Z.
- If the element's `TS_Point_Description` parameter is read-only (for example, locked by a workset or the parameter setup), the change will not be saved to the model even though the table shows it. Reopen the tool to see the true values.

## Sort order

Points are sorted in natural order, so `P-2` comes before `P-10`.

## Things to know

- **Active view only.** The list is built from the view that is active when you open the tool. If you switch views, close and reopen the tool to refresh the list.
- **Only elements with a Point Number are listed.** Elements where `TS_Point_Number` is empty or missing are skipped.
- **The window stays open** while you work in Revit and stays on top of other windows, so you can click through points one after another.
- **Shared parameters.** The tool loads its shared parameters (`TS_Point_Number`, `TS_Point_Description`) each time it starts. If the points are not showing up, confirm those parameters are on the elements and have values.

## Files used

| File | Purpose |
|------|---------|
| `C:\Temp\Ribbon_TSPointNumber.txt` | Stores your last search text. Created automatically. Safe to delete. |

## Troubleshooting

| Problem | Likely cause and fix |
|---------|----------------------|
| Table is empty | No elements in the active view have a `TS_Point_Number` value. Check the active view and the parameter values. |
| A point I expect is missing | It is not visible in the active view, or its Point Number is blank. |
| Clicking a row does nothing | The element may be hidden or in a view where it cannot be shown. Try a different view. |
| Description edit did not stick | The `TS_Point_Description` parameter is read-only for that element. |
| Old search text keeps appearing | Clear the Find box and close the window, or delete `C:\Temp\Ribbon_TSPointNumber.txt`. |
