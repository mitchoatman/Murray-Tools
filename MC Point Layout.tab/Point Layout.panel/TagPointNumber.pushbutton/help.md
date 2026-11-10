# Tag Point Number

**Why and when do we tag point numbers?**

**Why** - Adds a tag showing the `TS_Point_Number` value to the points and fixtures in the active view, so the point numbers show on the drawing without tagging each one by hand. The tool also spaces the tags so they do not sit on top of each other.

**When** - As requested, after the points have point numbers and before the view is checked or placed on a sheet.

**The tool does not work in a 3D view.** If you try it there, a message says to run it from a 2D view (plan, section, or elevation).

## **1. Open the view:**

* Open the 2D view you want to tag, for example a plan view.
* The tool looks only at what is visible in the active view.
* It looks at these categories: Pipe Accessories, Plumbing Fixtures, Structural Stiffeners, Generic Models, and Duct Accessories.
* An element is included only if its `TS_Point_Number` parameter has a value. Nested family parts are not included, only the parent family.

---

## **2. Click the Tag Point Number button:**

* Click the **Tag Point Number** button on the Murray Ribbon. [CONFIRM: ribbon tab and panel and button name]
* A **Select TS Point Numbers to Tag** window opens. It lists every point number found in the view.
* If no element in the view has a `TS_Point_Number` value, a message says none were found and the tool stops.

---

## **3. Choose what to tag:**

Use one of these two ways.

**A. Pick point numbers from the list:**

* To narrow the list, type in the **Search Point Number** box. The list shows only numbers that contain what you typed.
* In the list, select the numbers you want. Hold **Ctrl** or **Shift** to select more than one.
* Click **Select All** to select every number in the list.
* Tick or clear **Include Leader**. It is ticked by default.
* Click **Tag Selected**.
* Every element in the view with a selected point number is tagged.
* If you click Tag Selected with nothing selected, nothing happens.

**B. Fence: pick elements in the view:**

* Tick or clear **Include Leader**.
* Click **Fence**. The window closes.
* If you already had elements selected in Revit, the tool uses them and does not ask you to pick.
* Otherwise, select the elements in the view, then click **Finish**. Use a window or click each one.
* Only elements that are in the view and have a `TS_Point_Number` value are tagged. Everything else you select is ignored.
* Press **Esc** to cancel.

Click **Cancel** in the window to close it without tagging.

---

## **4. Check the result:**

* If the tag family is not in the project, the tool loads it for you.
* Each element gets a horizontal tag showing its point number.
* The tag starts at the element and is then moved to a nearby open spot so it does not overlap other tags. If it cannot find an open spot, it stays at the original position.
* A Windows notification near the system clock says "Tagging Complete" and shows how many elements were tagged.
* All the tags go in as a single step, so one Undo removes them all.
* Move or adjust tags by hand if any still overlap other items.

**ONE THING TO NOTE: ELEMENTS THAT ALREADY HAVE A TAG IN THE VIEW ARE SKIPPED, SO RUNNING THE TOOL AGAIN DOES NOT ADD DUPLICATE TAGS.**

**A COUNT OF 0 MEANS EVERYTHING YOU CHOSE WAS ALREADY TAGGED.**

---

## **What the tool does for you:**

* **Finds the points.** It collects the elements in the view with a point number and lists the numbers in natural order (1, 2, 10, not 1, 10, 2).
* **Loads the tag family.** The Multi Category Tag -TS_Point_Number family is loaded if it is missing.
* **Skips tagged elements.** No duplicate tags.
* **Spaces the tags.** It tries nearby spots above, below, and to the sides to avoid overlaps. The overlap check is an estimate based on the tag text length.
* **Tags in one step.** One Undo removes everything.

---

## **If something goes wrong:**

* **"Please run this tool from a 2D view (plan, section, or elevation)."** Open a 2D view and run the tool again.
* **"No elements with a valid 'TS_Point_Number' parameter were found in the active view."** Nothing in the view has a point number, or the elements are not in the categories listed in step 1. Check the view and the parameter.
* **The window closes and nothing happens.** No point numbers were selected, nothing was picked in Fence mode, or Cancel was clicked.
* **No notification appeared.** The Windows notification can be turned off. Check the view for the tags.
* **"Failed to load family: Family file not found."** The tag family file is missing from the tool folder. Tell the tool owner. [CONFIRM: tool owner]
* **"Could not find tag type ..."** The tag family does not have the expected type. Tell the tool owner.
* **"Failed while tagging elements."** The tags could not be placed and nothing was changed. Try again, and tell the tool owner if it repeats.
* **Some tags still overlap.** The spacing is an estimate. Move those tags by hand.
* **An element was not tagged.** It has no point number, it is not a parent family, or it already has a tag in the view.

**CHECK EVERY TAG AGAINST THE POINT NUMBERS IN THE MODEL OR LAYOUT DATA BEFORE THE VIEW IS PLACED ON A SHEET OR SENT TO ANOTHER TRADE.**

[CONFIRM: Egnyte link for the training video or the tag family, if one exists]

<!--
OPEN ITEMS (remove before posting to the Help Hub)
1. [CONFIRM: tool title, tooltip, and author from bundle.yaml. Not provided.]
2. [CONFIRM: exact Murray Ribbon tab and panel and the button name.]
3. [CONFIRM: tool owner name for the support lines.]
4. [CONFIRM: add the buttons: line in this help.md, using the same format as the other topics. Format not provided.]
5. Topic folder and topic.txt name not provided. Other topics use names like MCPointloadDuct or ChangeRodSize.
6. No images are referenced in this file by request. The page is text only.
-->
