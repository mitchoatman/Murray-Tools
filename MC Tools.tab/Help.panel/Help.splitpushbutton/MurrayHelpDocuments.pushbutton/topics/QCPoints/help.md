# QC Points

Exports QC points from the active Revit view as AutoCAD `POINT` commands. Paste them into a DWG command line to drop the points at the correct shared coordinates.

## What it exports

Run it in the view that shows the hangers you want. Only elements visible in the active view are included.

| Element | Points created |
|---|---|
| Fabrication hanger, single rod | 1 point at the rod end |
| Fabrication hanger, trapeze (2+ rods) | 1 point per rod end |
| GTP generic model family | 1 point at the family location |

Elements are only exported if they (or the parent family that owns them) have a value in the **TS_Point_Number** parameter. Elements with no TS_Point_Number are skipped.

## Before you run

1. Open the view with the hangers and GTPs you want to export.
2. Confirm TS_Point_Number is filled in on the hangers and on the GTP parent families.
3. Confirm the model has the correct shared coordinates set (the tool reads shared coordinates, not internal).

## Two versions

### Quick version (QCPoints_script)

1. Click the button.
2. A Windows notification appears: "N point(s) copied to clipboard".
3. Open the DWG, click the command line, and paste.

Notes:
- Beam hangers are always excluded.
- Finds GTPs nested inside other families (sub-components) as well as standalone GTPs.
- Shows the Murray icon in the notification if `Murray.ico` is in the same folder as the script.

### Window version (QCPoints_config)

1. Click the button.
2. A window lists the commands and the point count.
3. Click **Copy**. The window closes with the text on the clipboard.
4. Open the DWG, click the command line, and paste.

Notes:
- Beam hanger exclusion is controlled by a settings file (see below).
- Finds hangers and standalone GTP generic models in the view. GTPs nested inside other families are not searched for.
- If no hangers or GTPs are in the view, a "No Elements" message appears.
- If elements are found but none have a TS_Point_Number, a "No Points" message appears.

## Beam hanger setting (window version only)

Settings file: `C:\Temp\ExportPoints.txt`

```
exclude_beam_hangers=true
```

- `true` (default, also used if the file is missing): hangers with **FP_Beam Hanger** = Yes are not exported.
- `false`: beam hangers are exported like any other hanger.

## Output format

The first line is `SETAPLCOLOR`. Each point after that is one line:

```
POINT East,North,Elevation
```

- Values come from the project's shared coordinates.
- Values are converted from feet to inches (multiplied by 12) and rounded to 8 decimals.
- Very small elevation values (under 1e-9) are written as 0.

## Troubleshooting

| Problem | Likely cause |
|---|---|
| "No Points" or nothing copied | No TS_Point_Number on the elements in the view |
| A hanger is missing | It is a beam hanger and exclusion is on, or it is not visible in the active view |
| A GTP is missing (window version) | The GTP is nested inside another family; use the quick version |
| Points land in the wrong place in CAD | Shared coordinates in the model are not set correctly, or the DWG is not on the same coordinate system |
| Notification does not appear | Windows notifications are blocked; the clipboard is still filled, so paste as normal |

## Checks

These tools are a drafting aid. Spot check a few pasted points against the Revit model before using them for layout or QC.
