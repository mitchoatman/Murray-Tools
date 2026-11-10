# Pipe Riser

**Why and when do we place pipe riser families?**

**Why** - Places a Pipe Riser family on vertical fabrication pipe, sized to the pipe, so riser locations can be seen, scheduled, and coordinated.

**When** - As requested, once the vertical fabrication pipe is modeled and the levels are set.

## **What it does:**

* Click the **Pipe Riser** button on the Murray Ribbon. [CONFIRM: ribbon tab and panel]
* Works in a **Floor Plan** or a **3D view**. No window opens for this method.
* Uses the fabrication pipe you have selected. If nothing is selected, it uses all fabrication pipe in the view.
* Only **vertical** pipe is used. Other parts are counted as skipped.
* Places a Pipe Riser at each level the pipe passes through.

  * Floor Plan: that plan's level only.
  * 3D view: every level. If a section box is on, only levels and pipe inside the section box.
* Loads the Pipe Riser family if it is not in the project.
* Sets the riser Diameter and Schedule Level.
* Fills `FP_Product Entry`, `FP_Service Name`, and `FP_Service Abbreviation` from the pipe.
* Does not place a second riser where one is already at the same spot. It does not update the old one.
* One Undo removes everything placed in one run.
* A **Pipe Riser** message shows how many were placed, how many non-vertical parts were skipped, and any warnings.

## **Shift+Click:** [CONFIRM: Shift+Click runs the pick method]

* Plan views only. Not available in a 3D view.
* Select the vertical, straight fabrication pipe when prompted, then finish the selection.
* Places one riser on each selected pipe at the current view's level.
* Updates the diameter of a riser already at that spot.
* Does not check that the pipe passes through the level. Check the levels.
* A **Result** message shows the counts.

## **Notes:**

* Check the message every time. A count of 0 usually means the pipe is not vertical, does not cross a level in the view, or a riser is already in place.
* Check riser locations, levels, and sizes against the model before use for coordination or before sending to another trade.
* If something fails, a message names the problem. Common ones: not in a Floor Plan or 3D view, Pipe Riser family file missing (tell the tool owner), or a parameter that could not be set.

<!--
OPEN ITEMS (remove before posting to the Help Hub)
1. These uploads (Piperiser_script_1.py, Piperiser_config_1.py) are identical to the earlier Piperiser files. Content matches the full Pipe Riser help page, in short form.
2. [CONFIRM: tool title, tooltip, and author from bundle.yaml. Not provided.]
3. [CONFIRM: ribbon tab and panel, and the button name.]
4. [CONFIRM: Shift+Click runs the pick method (config file).]
5. [CONFIRM: tool owner name.]
6. [CONFIRM: add the buttons: line, same format as the other topics. Format not provided.]
7. Topic folder and topic.txt name not provided.
-->
