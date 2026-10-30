# Make Schedule

1. **Know your category.**

   1. Select the elements in Revit to query their category.
2. **Make a new schedule.**

   1. Right click on Schedules in project browser.
   2. Select New Schedule/Quantities
   3. Choose the category of the elements you wish to schedule.
   4. Name the Schedule
3. **Choose the data fields you wish to see.**

   1. Count = **Quantity**
   2. Family = **Name**
   3. Level or Reference Level (Depending on category chosen) = **Level**
   4. Size or Type or Overall size or Size of Primary End = **Size**
   5. Length = **Length**
4. **Filter your data (if needed).**

   1. **Revit schedules will always search the entire model. You must filter the data to restrict it to a level or subset of elements.**
   2. Filtering by Family is available on Revit 2022+
   3. Use negate filtering to have multiple filters (does NOT contain, does NOT equal). Usually you are filtering schedules from what you **DON'T** want to allow you to get what you **DO** want.
5. **Sort and group the data to get data to run in numerical / alphabetical order. (Important step)**

   1. Turn of Itemize every instance
   2. Sorting and grouping by field in order of importance to you.
   3. Enable Headers and Footers for more grouping
   4. If you see varies or no data in any cell, you have sorted/grouped improperly.

      1. Usually adding more sort by fields will resolve the issue
6. **Formatting**

   1. Use to hide fields from resulting schedule.

      1. Why add a field u don’t want to see (good question).

         1. Sometimes you need to add fields to have access to more data to run filters against but you don’t need to show end user that data.
   2. Use to have length fields calculate a total
   3. Use to change units in schedule to report differently than units in the model



https://murraycompany.egnyte.com/dl/kJQwxCK7D4XF
