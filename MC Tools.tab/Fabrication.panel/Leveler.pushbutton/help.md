# Leveler

**Why and when do we set levels on fabrication parts?**

**Why** - Sets the **Reference Level** of fabrication pipework and ductwork from where each part sits in height, so the parts report the right level for schedules, tags, filters, and exports. [CONFIRM: purpose and typical use]

**When** - As requested, after the fabrication parts are modeled and the levels are set, and again after parts or levels move.

There are 2 ways to set levels:

* **Click** the button and choose a Top and Bottom level (sections 1 to 4).
* **Shift+Click** the button to set levels automatically with no window (section 5). [CONFIRM: Shift+Click runs the automatic method]

**Both methods work only on fabrication pipework and ductwork that is visible in the active view.** Parts that are hidden or filtered out of the view are not changed. The height used is the **bottom** of each part.

## **1. Open the view:**

* Open the view that shows the parts you want to set. Make sure the parts are visible.
* Use view filters, a section box, or a view range to limit which parts are in the view.
![](images/img-01.png)

---

## **2. Click the Leveler button:**

* Click the **Leveler** button on the Murray Ribbon. [CONFIRM: ribbon tab and panel]
* A **Select Levels** window opens.
![](images/img-02.png)

---

## **3. Choose the Top and Bottom levels:**

The window reads "Set zone of elements to set level on." Choose the two levels that bound the parts.

* **Choose Top Level** and **Choose Bottom Level** list every level in the project, plus **(None)**.
* The window starts with the lowest level as Bottom and the next level up as Top.
* Parts whose bottom is between the Bottom and Top elevations (including the elevations themselves) get the **Bottom** level.

**Special cases:**

* **Roof:** set Bottom to the roof level and Top to **(None)**. Parts at or above the Bottom level get the Bottom level.
* **Underground:** set Top to the level and Bottom to **(None)**. Parts at or below the Top level get the Top level.
* **Both set to (None):** nothing is changed.
* **Bottom higher than Top:** nothing is changed.
![](images/img-03.png)

---

## **4. Click OK and check the result:**

* Click **Ok** to set the levels. Click **Cancel** to close without changing anything.
* The tool does not show a confirmation message.
* Select a part and check **Reference Level** in Properties.
* All the changes go in as a single step, so one Undo puts them all back.
* Run it again for the next zone, for example the next floor.
![](images/img-04.png)

**ONE THING TO NOTE: THE TOOL ONLY CHANGES PARTS BETWEEN THE LEVELS YOU CHOSE. PARTS OUTSIDE THAT ZONE KEEP THEIR OLD LEVEL.**

**CHECK THE TOP AND BOTTOM LEVELS IN THE WINDOW BEFORE YOU CLICK OK EVERY TIME.**

---

## **5. Shift+Click: set levels automatically:**

This method needs no window. It gives every part in the view the closest level at or below the bottom of the part.

* Open the view with the parts you want to set.
* **Shift+Click** the **Leveler** button.
* The tool uses the levels that are visible in the active view. If none are visible, it stops with an error.
* Parts below the lowest visible level get the lowest level.
* Only parts that need a change are changed. Parts with no size, or a Reference Level that cannot be edited, are skipped.
* In a workshared model, the tool checks the parts out first. **Parts owned by another user are skipped and no message tells you.** Ask the owner to relinquish them, then run the tool again.
* The tool does not show a confirmation message. One Undo puts all changes back.
![](images/img-05.png)

---

## **What the tool does for you:**

* **Finds the parts.** Fabrication pipework and ductwork in the active view.
* **Uses the bottom of each part.** Height is measured at the bottom of the part.
* **Sets Reference Level.** It writes the level for you, so you do not edit parts one at a time.
* **Changes in one step.** One Undo puts everything back.

---

## **If something goes wrong:**

* **Nothing changed.** The parts are not in the active view, the levels you chose do not bound them, or both levels are (None).
* **"No levels visible in the active view."** Shift+Click method only. Show the levels in the view, or use the Click method.
* **"Error processing element ..."** A part could not be changed, often because its Reference Level cannot be edited. Check the part, and tell the tool owner if it repeats. [CONFIRM: tool owner]
* **"Script Failed."** The tool could not finish. Note the message and tell the tool owner.
* **Some parts did not change (workshared model).** They are owned by another user. Ask them to relinquish the parts, then run the tool again.
* **Parts got the wrong level.** Press Undo and run it again with the right levels.

**CHECK THE REFERENCE LEVELS AGAINST THE MODEL AND THE LEVEL SCHEDULE BEFORE THEY ARE USED FOR SCHEDULES, EXPORTS, OR SENT TO ANOTHER TRADE.**

[CONFIRM: Egnyte link for the training video, if one exists]

<!--
OPEN ITEMS (remove before posting to the Help Hub)
1. [CONFIRM: tool title, tooltip, and author from bundle.yaml. Not provided.]
2. [CONFIRM: exact Murray Ribbon tab and panel and the button name.]
3. [CONFIRM: Shift+Click runs the automatic method (Leveler_config.py). The click script is Leveler_script.py. File names suggest this but the bundle.yaml was not provided.]
4. [CONFIRM: purpose and typical use for the Why line.]
5. [CONFIRM: tool owner name for the support lines.]
6. [CONFIRM: add the buttons: line in this help.md, using the same format as the other topics. Format not provided.]
7. Topic folder and topic.txt name not provided. Other topics use names like MCPointloadDuct or ChangeRodSize.
8. Placeholder images are in images/. Replace each with a real screenshot using the same file name:
   img-01.png  a view with fabrication parts visible
   img-02.png  the Leveler button on the Murray Ribbon
   img-03.png  the Select Levels window with levels chosen
   img-04.png  a part selected with Reference Level in Properties
   img-05.png  Shift+Click on the Leveler button
-->
