# Change Rod Size

Changes the rod size on fabrication hangers you select. One choice is applied to all selected hangers.

**USING IT:**

1. Click the **Change Rod Size** button.

2. Select the fabrication hangers you want to change.
    1. Only hangers in the MEP Fabrication Hangers category can be picked.
    2. Click the finish button in the top left corner when you are done selecting.

3. Pick a rod size in the **Select Rod Size** window and click OK.
    1. Choices are 3/8, 1/2, 5/8, 3/4, 7/8, 1, 1-1/4 and 1-1/2.
    2. The window opens on 3/8, so check it before you click OK.

**WHAT HAPPENS:**

1. Every selected hanger is switched to the rod size you picked.

2. The `FP_Rod Size` parameter is then updated from the hanger's new rod.
    1. A hanger that reports no rod diameter keeps its existing `FP_Rod Size` value.

3. The whole change runs as one step, so a single undo reverses it.

**GOOD TO KNOW:**

- If you cancel the selection or close the rod size window, nothing is changed.
- Clicking finish with no hangers selected shows a message and changes nothing.
- If you see "Operation cancelled by user." and did not cancel, something failed. Let the admin know.

> This tool only edits the model. Confirm the required rod size for the support before you change it.
