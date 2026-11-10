# Move Control Point

**Why and when do we move control points?**

**Why** - Moves a group of control points by the same X and Y distance in one step, for example to shift a set of points that were placed or imported off by a known amount. [CONFIRM: purpose and typical use]

**When** - As requested, after control points are in the model and you know how far they need to move.

**The tool moves Generic Model elements, not only Control Point families.** Only elements in the Generic Models category can be selected.

## **1. Select the elements:**

* Open the view that shows the points you want to move.
* Click the **Move Control Point** button on the Murray Ribbon.
* Select the points by dragging a window around them, from left to right.

  * Only Generic Model elements can be selected. Other elements will not highlight.
  * Finish the selection when you are done.
* If you select nothing, or cancel, a message appears and the tool stops.

---

## **2. Enter the distances:**

A **Move Generic Model Elements** window opens.

* **X Distance (ft):** positive moves right, negative moves left.
* **Y Distance (ft):** positive moves up, negative moves down.
* Enter the distances as decimal feet. Examples: `2.5`, `-0.25`, `10`.

  * Do not use inches, fractions, or units. Example: for 6 inches, enter `0.5`.
* The window remembers the last distances you used. It starts at 0 and 0 the first time.
* Click **OK**.
* If you close the window without clicking OK, nothing is moved.

---

## **3. Check the result:**

* Every selected element moves by the same distance, in X and Y only. The height does not change.
* All the points move in a single step, so one Undo puts them all back.
* The tool does not show a confirmation message. Check the points in the view.

**ONE THING TO NOTE: THE DISTANCES ARE SAVED, SO THE NEXT RUN STARTS WITH THE SAME VALUES.**

**CHECK THE X AND Y VALUES BEFORE YOU CLICK OK EVERY TIME, OR POINTS MAY MOVE THE WRONG WAY.**

---

## **What the tool does for you:**

* **Limits the selection.** Only Generic Model elements can be picked, so other elements cannot be moved by mistake.
* **Moves as a group.** Each element moves by the same distance from where it is now.
* **Remembers your values.** The last X and Y distances are filled in the next time.

---

## **If something goes wrong:**

* **"No generic model elements selected. Script will exit."** Nothing was selected. Run the tool again and select the points.
* **"Selection cancelled or failed. Script will exit."** The selection was cancelled. Run the tool again.
* **"Please enter valid numeric values for distances. Script will exit."** One of the boxes has text that is not a number. Use decimal feet, with no units or fractions.
* **The points moved the wrong way.** Press Undo, then run the tool again with the signs reversed.
* **The points moved the wrong distance.** The distances are in feet. Press Undo and run the tool again.
* **Some points were not selected.** The elements are not in the Generic Models category, or they were not fully inside your window. Select them again.

**CHECK THE NEW POINT LOCATIONS AGAINST THE CONTROL OR LAYOUT DATA BEFORE THEY ARE USED IN THE FIELD OR SENT TO ANOTHER TRADE.**

[CONFIRM: Egnyte link for the training video, if one exists]

<!--
OPEN ITEMS (remove before posting to the Help Hub)
1. [CONFIRM: tool title, tooltip, and author from bundle.yaml. Not provided.]
2. [CONFIRM: exact Murray Ribbon tab and panel and the button name.]
3. [CONFIRM: purpose and typical use for the Why line.]
4. [CONFIRM: directions. The script moves along the model's X and Y axes. Right and up match a plan view set to Project North. Check a view set to True North.]
5. [CONFIRM: whether Control Point families are in the Generic Models category in this firm's family.]
6. [CONFIRM: tool owner name for the support lines.]
7. [CONFIRM: add the buttons: line in this help.md, using the same format as the other topics. Format not provided.]
8. Topic folder and topic.txt name not provided. Other topics use names like MCPointloadDuct or ChangeRodSize.
9. No images are referenced in this file by request. The page is text only.
-->
