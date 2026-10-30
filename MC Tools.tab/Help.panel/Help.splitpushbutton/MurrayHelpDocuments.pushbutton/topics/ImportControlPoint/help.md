# Import Control Points

**Why and when do we import control points?**

**Why** - Places Control Point families in the model from a CSV file of coordinates, and fills in each point's Point Number and Point Description. This replaces placing and numbering each point by hand.

**When** - As requested, when you have a list of points (for example from a survey or layout file) that needs to go into the model.

This tool places points from a file. To place points by hand or at grid intersections, use the **Control Point** tool instead.

## **1. Prepare the CSV file:**

The CSV needs a header row and one row per point. The first row is always treated as the header.

* Click **Make Template** in the tool window to create a blank CSV on your Desktop.

  * The file is named `CONTROL_POINT_TEMPLATE.csv`.
  * It has these columns: POINT NUMBER, Y, X, Z, DESCRIPTION.
  * The tool offers to open it for you.
* Fill in one row per point.
* Coordinates must be numbers. Use decimal feet. [CONFIRM: units of the survey data used by the team]
* Rows with a missing or non-numeric X, Y, or Z are skipped and reported in the result.
* The file can be separated by commas, semicolons, tabs, or pipes.

**Columns and accepted header names (Auto order):**

* **X:** X, E, EAST, EASTING
* **Y:** Y, N, NORTH, NORTHING
* **Z:** Z, H, EL, ELEV, ELEVATION, HEIGHT
* **Point number:** POINT NUMBER, POINT NO, POINT NUM, POINT ID, NUMBER, PT NO, PT NUM, TS_Point_Number
* **Description:** DESCRIPTION, DESC, POINT DESCRIPTION, POINT DESC, TS_Point_Description

Capital letters, spaces, and underscores in the header do not matter. Point number and description columns are optional. If they are missing, those fields are left blank.

---

## **2. Open the tool and choose the settings:**

* Click the **Import Control Points** button on the Murray Ribbon.
* The **Import Control Points from CSV** window opens. It remembers the settings from last time.
* **Control Point Type:** the type of the Control Point family to use. The default is **CP**. If the type does not exist in the project, the tool creates it. Letters are changed to capitals as you type.
* **CSV Coordinate System:** choose one.

  * **Shared** (default) means the CSV values are in the project's shared coordinates (easting, northing, elevation). The tool converts them to the model's position.
  * **Project Internal** means the CSV values are in the model's own coordinates and are used as they are.
* **Import Order:** choose one.

  * **Auto** (default) finds the X, Y, and Z columns by their header names (section 1).
  * **(Y,X,Z : N,E,H)** ignores the header names and reads the columns by position, with Y first.
  * **(X,Y,Z : E,N,H)** ignores the header names and reads the columns by position, with X first.
* **CSV File:** type the path or click **Browse...** to find the file.

**Column positions when you choose Y,X,Z or X,Y,Z:**

* **4 or 5 columns:** column 1 is the Point Number, columns 2 to 4 are the coordinates in the order you chose, and column 5 (if present) is the Description.
* **3 columns:** the columns are the coordinates only, in the order you chose.
* **Fewer than 3 columns:** the tool stops with a message.

---

## **3. Import:**

* Click **Import**.
* If the Control Point family is not in the project, the tool loads it for you.
* The tool places one Control Point for every valid row and writes the Point Number to `TS_Point_Number` and the Description to `TS_Point_Description`.
* If a point cannot be placed in 3D, the tool places it on the nearest level instead.
* All the points go in as a single step, so one Undo removes them all.
* Click **Cancel** to close the window without importing.

When it finishes, a **CSV Import Result** message shows:

* How many points were placed and how many failed.
* How many CSV rows were skipped before placement.
* The coordinate system, import order, family, type, and file that were used.
* A parameter warning, if the Point Number or Description could not be written.
* The first 10 skipped rows and the first 10 placement failures, with their row numbers.

**CHECK THE RESULT MESSAGE EVERY TIME. ONE BAD ROW DOES NOT STOP THE REST, SO A PARTIAL IMPORT IS POSSIBLE.**

**CHECK THE PLACED POINTS AGAINST THE SURVEY OR LAYOUT DATA BEFORE THEY ARE USED IN THE FIELD OR SENT TO ANOTHER TRADE.**

---

## **If something goes wrong:**

* **"Please select a CSV file."** Browse to the file, or type its path.
* **"File Not Found."** The path in the CSV File box is wrong. Browse to the file again.
* **"Please enter a control point type."** The Control Point Type box is empty. Enter CP or your type name.
* **"CSV file is empty" or "CSV header row is empty."** The file has no header row. Add one.
* **"Could not detect X/Y/Z columns."** The header names are not on the accepted list in section 1. Rename the headers, or choose Y,X,Z or X,Y,Z so the tool reads by position.
* **"CSV must contain at least 3 columns for hard-set import order."** The file has fewer than 3 columns.
* **"No valid point rows found in CSV."** Every row was skipped. The message lists the first issues. Check that X, Y, and Z are numbers.
* **"Family file not found."** The Control Point family file is missing from the tool folder. Tell the tool owner. [CONFIRM: tool owner]
* **"Could not evaluate shared coordinates."** The project's shared coordinates are not set up correctly. Check the project location, or use Project Internal.
* **A Parameter warning.** The points were placed, but `TS_Point_Number` or `TS_Point_Description` was not written on some of them. Check that both parameters exist on the Control Point family as text parameters and are not read-only.
* **Points land in the wrong place.** Check the coordinate system choice (Shared or Project Internal) and the import order (X first or Y first). Undo, then run it again with the right settings.

[CONFIRM: Egnyte link for the training video or the family file, if one exists]

<!--
OPEN ITEMS (remove before posting to the Help Hub)
1. [CONFIRM: tool title, tooltip, and author from bundle.yaml. Not provided.]
2. [CONFIRM: exact Murray Ribbon tab and panel and the button name.]
3. [CONFIRM: units of the survey CSVs. The script reads values as decimal feet. Decimal commas are not supported.]
4. [CONFIRM: tool owner name for the support lines.]
5. [CONFIRM: add the buttons: line in this help.md, using the same format as the other topics. Format not provided.]
6. Topic folder and topic.txt name not provided. Other topics use names like MCPointloadDuct or ChangeRodSize.
7. No images are referenced in this file by request. The page is text only.
-->
