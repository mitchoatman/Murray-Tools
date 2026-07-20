This 'hook' works because of the following:

The script in this folder writes a txt file to c:\temp indicating on\off.
The startup script for MC ribbon has a trigger to read that status and set icon on revit open.
The doc-updater file in \hooks folder fires upon every doc update (IF) the txt file is set to true or on.