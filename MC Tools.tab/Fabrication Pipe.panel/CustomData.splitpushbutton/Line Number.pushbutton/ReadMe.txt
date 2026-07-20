This button opens a dockable pane.

What makes the pane work?

startup.py  - there is code in there to register pane with unique GUID
lib folder -  \LineNumberPane that is where the pane functionality and design is stored. pane.py  config.py
pushbutton folder - that is the button in Revit ribbon that toggles the display of the panel feeding from startup and lib files.
