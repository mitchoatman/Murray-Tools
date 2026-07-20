This 'hook' works because of the following:

The script in this folder writes a txt file to c:\temp indicating on\off.
The startup script for MC ribbon has a trigger to read that status and set icon on revit open.
The view-activated file in \hooks folder fires upon view change (IF) the txt file is set to true or on.
only works with floorplan views and if the project is workshared and if there are worksets named the same as levels.