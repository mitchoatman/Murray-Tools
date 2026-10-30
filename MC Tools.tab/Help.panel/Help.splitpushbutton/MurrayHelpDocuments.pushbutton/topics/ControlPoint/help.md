# Control Point

**Why and when do we place control points?**

**Why** - Places a Control Point family in the model and fills in its Point Number and Point Description parameters, so the points can be scheduled, tagged, and exported for layout.

**When** - As requested, when control points are needed in a plan view. Place them one at a time by clicking in the view, or all at once at grid intersections.

There are 2 ways to place control points:

* **Click** the button to place points by hand (section 1).
* **Shift+Click** the button to place points at grid intersections (section 2). \[CONFIRM: Shift+Click runs the grid method]

**The tool does not work in a 3D view.** If you try it there, a message says "Cannot use in 3D view." Open a plan view first.

!\[](images/img-01.png)

## **1. Click: place points by hand:**

This method is best for a few points at chosen locations.

* Open a plan view.
* Click the **Control Point** button on the Murray Ribbon.
* A **Set Control Point Data** window opens.

  * **Point Number:** the number to write to the points.
  * **Point Description:** a description of the points.
  * Both boxes are required. The window remembers what you typed last time.
!\[](images/img-02.png)
* Click **Place**.
* Click in the view to place each point.
* Press **Esc** when you are done placing.
* The tool then writes the Point Number and Point Description to every point you placed in that run.
* If you place no points, nothing is written.

**ONE THING TO NOTE: EVERY POINT PLACED IN ONE RUN GETS THE SAME POINT NUMBER.**

**TO GIVE POINTS DIFFERENT NUMBERS, RUN THE TOOL AGAIN FOR EACH NUMBER.**

\---

## **2. Shift+Click: place points at grid intersections:**

This method places a point at every intersection of the grid lines you select, and numbers them from the grid names.

* Open a plan view.
* **Shift+Click** the **Control Point** button.
* Select 2 or more grid lines.

  * Use a window or click each grid, then finish the selection.
  * The tool finds every place the selected grids cross. If none cross, a message says no intersections were found.
!\[](images/img-03.png)
* A **Set Control Point Data** window opens. Choose the **Point Number Mode**:

  * **Append Grid Names** (default) writes the number for you as `CP\_` plus the two grid names. The first name is the more vertical grid and the second is the more horizontal grid. Example: `CP\_A-3`. The Point Number box is not used in this mode.
  * **Custom** uses the number you type. Every point gets that same number.
* Type the **Point Description**. It is required, and every point gets the same description.
* Click **Place**.
* A message shows how many points were placed, for example: Placed 12 control point(s).


All the points go in as a single step, so one Undo removes them all.

\---

## **What the tool does for you:**

* **Loads the family.** If the Control Point family is not in the project, the tool loads it for you. It uses the **CP** type.
* **Writes the parameters.** Point Number goes to `TS\_Point\_Number` and Point Description goes to `TS\_Point\_Description`.
* **Remembers your entries.** The last Point Number, Point Description, and mode are filled in the next time you open the window.

\---

## **If something goes wrong:**

* **"Cannot use in 3D view."** Open a plan view and run the tool again.
* **"Please enter a Point Number."** Custom mode and the click method need a Point Number.
* **"Please enter a Point Description."** Type a description before you click Place.
* **"Please select at least 2 grid lines."** Shift+Click method only. Select two or more grids.
* **"No intersections were found between the selected grid lines."** The grids you picked do not cross. Pick grids that cross each other.
* **"Family file not found."** The Control Point family file is missing from the tool folder. Tell the tool owner. \[CONFIRM: tool owner]
* **"Type 'CP' was not found in family 'Control Point'."** The family in the project does not have a CP type. Tell the tool owner.
* **A Warning that a parameter was not found, is read-only, or is not a text parameter.** The points were placed, but the number or description was not written. Check that `TS\_Point\_Number` and `TS\_Point\_Description` exist on the Control Point family as text parameters.
* **Revit 2022 only: "Family loaded successfully. Run the command again to place the family."** The first run after loading only loads the family. Run the tool again.

**CHECK THE POINT LOCATIONS AGAINST THE CONTROL OR LAYOUT DATA BEFORE THEY ARE USED IN THE FIELD OR SENT TO ANOTHER TRADE.**

\[CONFIRM: Egnyte link for the training video or the family file, if one exists]

<!--
OPEN ITEMS (remove before posting to the Help Hub)
1. \[CONFIRM: Shift+Click runs the grid script (ControlPoint\_config.py). The click script is ControlPoint\_script.py. File names suggest this but the tool bundle.yaml was not provided.]
2. \[CONFIRM: tool title, tooltip, and author from bundle.yaml. Not provided.]
3. \[CONFIRM: exact Murray Ribbon tab and panel for the button.]
4. \[CONFIRM: which grid direction (letters or numbers) the owner considers "Y". The tool uses the more vertical grid name first.]
5. \[CONFIRM: Revit versions supported. The script has a Revit 2022 only workaround.]
6. \[CONFIRM: tool owner name for the support lines.]
7. \[CONFIRM: add the buttons: line in this help.md, using the same format as the other topics. Format not provided.]
8. Topic folder and topic.txt name not provided. Other topics use names like MCPointloadDuct or ChangeRodSize.
9. Screenshots needed, saved in images/ with these names:
   img-01.png  Control Point family placed in a plan view
   img-02.png  the Set Control Point Data window (click method)
   img-03.png  grid lines selected in a plan view
   img-04.png  the Set Control Point Data window with the Point Number Mode options

