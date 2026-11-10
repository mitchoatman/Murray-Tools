# Deck Sleeve Schedules

**Why and when do we create the deck sleeve schedule?**

**Why** - Creates the **DECK SLEEVE SCHEDULE** for you, with the columns, sorting, and filters already set, so you do not build the schedule by hand on every project.

**When** - As requested, after the Round Floor Sleeve families are placed and numbered, and before the schedule is placed on a sheet.

**No window opens for this tool.** Click the button and the schedule is created or updated, then opened.

## **1. Check the sleeves:**

* The sleeves should be in the project before you run the tool. The schedule is built from Pipe Accessories.
* Each sleeve should have a value in `TS_Point_Number`. It becomes the ITEM NO column.
* The tool filters the schedule to the Round Floor Sleeve family. This only works if at least one Round Floor Sleeve is already in the project.

---

## **2. Click the Deck Sleeve Schedules button:**

* Click the **Deck Sleeve Schedules** button on the Murray Ribbon. [CONFIRM: ribbon tab and panel and button name]
* You can run it from any view.

---

## **3. Check the schedule:**

The tool then does one of two things.

**If there is no DECK SLEEVE SCHEDULE yet**, it creates one and opens it:

* **Name:** DECK SLEEVE SCHEDULE
* **Category:** Pipe Accessories
* **Columns:**

  * **ITEM NO** from `TS_Point_Number`
  * **SIZE** from `Diameter`
  * **LENGTH** from `Length`
  * **NAME** from Family
* **Sorting:** by ITEM NO, ascending.
* **Filter (Revit 2022 and newer):** it shows only the Round Floor Sleeve type, found from the first Round Floor Sleeve in the project.
* **Filter by Sheet (Revit 2023 and newer):** turned on, so the schedule lists only the sleeves in views on the sheet where it is placed.

**If a DECK SLEEVE SCHEDULE already exists**, it updates it and opens it:

* The sorting by ITEM NO and, in Revit 2023 and newer, Filter by Sheet are set again.
* Columns, headings, and the filter are not changed. The tool does not add anything that is missing.

The tool does not show a confirmation message. The schedule opens as the active view.

---

## **4. Place the schedule on a sheet:**

* The tool does not place the schedule on a sheet.
* In the Project Browser, drag **DECK SLEEVE SCHEDULE** onto the sheet.
* Check the schedule on the sheet. With Filter by Sheet on, it should list only the sleeves in the views on that sheet.

**ONE THING TO NOTE: IF THE SCHEDULE ALREADY EXISTS, THE TOOL DOES NOT FIX ITS COLUMNS OR FILTER.**

**TO REBUILD IT FROM SCRATCH, DELETE THE OLD DECK SLEEVE SCHEDULE AND RUN THE TOOL AGAIN.**

---

## **What the tool does for you:**

* **Builds the schedule.** Fields, column headings, and sort order are set for you.
* **Finds the sleeve type.** It looks up the Round Floor Sleeve type for the filter.
* **Reuses the schedule.** If one exists, it re-applies the sort and the sheet filter and opens it.
* **Opens it.** The schedule becomes the active view.

---

## **If something goes wrong:**

* **A column is missing.** The tool adds a column only if the parameter is in the project. Check that `TS_Point_Number`, `Diameter`, and `Length` exist as project or shared parameters on the Pipe Accessories category.
* **The schedule lists the wrong items.** No Round Floor Sleeve was in the project when the schedule was created, so no filter was added. Or the project is Revit 2021 or older, which gets no filter. Place a Round Floor Sleeve, delete the schedule, and run the tool again.
* **Only one sleeve type shows.** The filter uses the type of the first Round Floor Sleeve found. Other types of the family are left out. Edit the filter by hand if you need them. [CONFIRM: whether the project uses more than one Round Floor Sleeve type]
* **The schedule is empty.** Check that sleeves exist, that `TS_Point_Number` has values if you need them, and that the sheet filter matches the views on the sheet. [CONFIRM: what the schedule shows before it is placed on a sheet]
* **Items are not in the right order.** The sort uses the ITEM NO column. If that column is missing, the sort is not set.
* **"Cannot open a view template" or "Cannot open this internal schedule view."** The schedule could not be opened. Open it from the Project Browser.
* **"Could not open view ..."** Revit could not open the schedule. Open it from the Project Browser, and tell the tool owner if it repeats. [CONFIRM: tool owner]

**CHECK THE SCHEDULE AGAINST THE MODEL AND THE SLEEVE LOCATIONS BEFORE IT IS PLACED ON A SHEET OR SENT TO ANOTHER TRADE.**

[CONFIRM: Egnyte link for the training video, if one exists]

<!--
OPEN ITEMS (remove before posting to the Help Hub)
1. [CONFIRM: tool title, tooltip, and author from bundle.yaml. Not provided.]
2. [CONFIRM: exact Murray Ribbon tab and panel and the button name.]
3. [CONFIRM: tool owner name for the support lines.]
4. [CONFIRM: Revit versions supported. The script treats 2022 and newer as filter-capable and 2023 and newer for Filter by Sheet.]
5. [CONFIRM: behavior of a sheet-filtered schedule that is not yet on a sheet.]
6. [CONFIRM: add the buttons: line in this help.md, using the same format as the other topics. Format not provided.]
7. Remember the topic.txt file and the topic folder name. Without topic.txt, Build Help does not create the stub help.html.
8. No images are referenced in this file. Copied the text-only layout of the last page. Say if you want placeholder PNGs.
-->
