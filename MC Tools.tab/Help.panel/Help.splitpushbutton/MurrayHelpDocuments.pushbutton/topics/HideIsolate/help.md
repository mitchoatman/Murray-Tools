# Hide / Isolate Service

Temporarily hides or isolates fabrication services in the active view.

**USING IT:**

1. Open the view you want to work in.
    1. The list only shows services that have fabrication hangers, pipe or duct in the current view.

2. Click the **Hide / Isolate** button.
    1. If you select elements first, their services start out checked in the list.

3. In the **Service Visibility** window, check the services you want.
    1. Type in the search box to narrow the list.
    2. Hold [SHIFT] and click a checkbox to check or uncheck every box between it and the last one you clicked.
    3. **All / None** checks every service showing in the list. Click it again to uncheck them.

4. Click **Hide** or **Isolate**.
    1. **Hide** hides the checked services in the view.
    2. **Isolate** shows only the checked services and hides everything else.
    3. The window closes and the view is in Revit's temporary hide/isolate mode.

5. To get back to the normal view, run the button again and click **Reset View**.
    1. Reset View ends temporary hide/isolate in the active view, including anything hidden or isolated another way.

**GOOD TO KNOW:**

- Hide and Isolate are temporary and only affect the active view.
- To keep the result, use Revit's **Apply Hide/Isolate to View** option from the temporary hide/isolate menu.
- Hide and Isolate use the services showing in the list at the moment you click. A service you checked and then filtered out with the search box may be left out. Clear the search box and check the list again before you click.
- If no service is checked, Hide and Isolate do nothing and the window stays open.
- If the view has no fabrication hangers, pipe or duct, a message says no fabrication services were found.
