# Tag Sleeve Size

**Why and when do we tag sleeve sizes?**

**Why** - Adds a size tag to every sleeve and pipe riser family in the active view in one click, so sleeve sizes show on the drawing without tagging each one by hand.

**When** - As requested, after the sleeves and risers are placed and before the view is checked or placed on a sheet.

**No window opens for this tool.** Click the button and it tags everything it finds in the active view.

## **1. Open the view:**

* Open the view you want to tag, for example a plan view.
* The tool works only on what is visible in the active view. Sleeves hidden in the view are not tagged.
* Make sure the sleeves and risers are in the view and not hidden by view settings.
![](images/img-01.png)

---

## **2. Click the Tag Sleeve Size button:**

* Click the **Tag Sleeve Size** button on the Murray Ribbon. [CONFIRM: ribbon tab and panel]
* If the sleeve size tag family is not in the project, the tool loads it for you.
* The tool then tags every matching element in the view. It tags pipe accessories whose type name contains any of these:

  * Metal Sleeve
  * Plastic Sleeve
  * Round Floor Sleeve
  * Cast Iron Sleeve
  * Pipe Riser
![](images/img-02.png)

---

## **3. Check the tags:**

* A horizontal tag is placed at the location of each sleeve or riser. The tags have no leader.
* The tool does not show a confirmation message. Look over the view to confirm the tags are there.
* All the tags go in as a single step, so one Undo removes them all.
* Move or adjust tags by hand if they overlap other items.
![](images/img-03.png)

**ONE THING TO NOTE: THE TOOL DOES NOT CHECK FOR EXISTING TAGS. EVERY RUN ADDS A NEW TAG TO EVERY SLEEVE, EVEN ONES ALREADY TAGGED.**

**RUN IT ONCE PER VIEW. IF YOU RUN IT TWICE, PRESS UNDO.**

---

## **What the tool does for you:**

* **Loads the tag family.** The Pipe Accessory - Trimble Sleeve Size Tag family is loaded if it is missing.
* **Finds the sleeves.** It looks only at pipe accessories in the active view whose type name matches the list in step 2.
* **Tags in one step.** All tags are placed together, so one Undo removes them.
* **Works with the Pipe Riser tool.** Risers placed by the Pipe Riser tool are included.

---

## **If something goes wrong:**

* **Nothing was tagged.** No pipe accessory in the view matches the list in step 2, or the sleeves are hidden in the view. Check the type names and the view's visibility settings.
* **"Failed to load family: Family file not found."** The tag family file is missing from the tool folder. Tell the tool owner. [CONFIRM: tool owner]
* **"Failed to collect family types" or "Failed to tag sleeves."** The tags could not be placed and nothing was changed. Try again in a different view, and tell the tool owner if it repeats.
* **Tags show the wrong tag type.** The tool places tags by category, so Revit uses the default Pipe Accessory tag type in the project. Check that the Sleeve Size tag (type Pipe Accessory Tag - SLV Tag) is the default, or change the type of the placed tags. [CONFIRM: behavior]
* **A sleeve has no tag.** Its type name is not on the list in step 2. Tag it by hand, or tell the tool owner to add the name.
* **Tags are doubled.** The tool was run more than once in the view. Press Undo, or delete the extra tags.

**CHECK EVERY TAG AND SLEEVE SIZE AGAINST THE MODEL AND THE SLEEVE SCHEDULE BEFORE THE VIEW IS PLACED ON A SHEET OR SENT TO ANOTHER TRADE.**

[CONFIRM: Egnyte link for the training video or the tag family, if one exists]

<!--
OPEN ITEMS (remove before posting to the Help Hub)
1. [CONFIRM: tool title, tooltip, and author from bundle.yaml. Not provided.]
2. [CONFIRM: exact Murray Ribbon tab and panel and the button name.]
3. [CONFIRM: tag type behavior. The script activates the type Pipe Accessory Tag - SLV Tag, but places tags by category, which uses the project's default Pipe Accessory tag type.]
4. [CONFIRM: view types supported. The script has no view type check. Behavior in 3D views, sections, and sheets is untested.]
5. [CONFIRM: tool owner name for the support lines.]
6. [CONFIRM: add the buttons: line in this help.md, using the same format as the other topics. Format not provided.]
7. Topic folder and topic.txt name not provided. Other topics use names like MCPointloadDuct or ChangeRodSize.
8. Placeholder images are in images/. Replace each with a real screenshot using the same file name:
   img-01.png  a view with sleeves and risers visible
   img-02.png  the Tag Sleeve Size button on the Murray Ribbon
   img-03.png  the same view with size tags placed
-->
