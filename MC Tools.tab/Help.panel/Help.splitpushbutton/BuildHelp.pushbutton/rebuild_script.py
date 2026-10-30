import os
import clr

# Load .NET libraries required for Windows Forms and drawing
clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')

from System.Windows.Forms import Form, PictureBox, PictureBoxSizeMode, FormBorderStyle, FormStartPosition, DockStyle
from System.Drawing import Size, Image
from System.IO import File, MemoryStream

# Get the directory of the currently running script
script_dir = os.path.dirname(__file__)
gif_path = os.path.join(script_dir, "smh-prince.gif")

# Print text to the pyRevit output window
# print("What are YOU doing?")

if os.path.exists(gif_path):
    # Read file into memory stream to prevent file locking
    byte_data = File.ReadAllBytes(gif_path)
    ms = MemoryStream(byte_data)
    img = Image.FromStream(ms)

    # Create a custom dialog window
    form = Form()
    form.Text = "What are YOU doing?"
    form.FormBorderStyle = FormBorderStyle.FixedDialog
    form.MaximizeBox = False
    form.MinimizeBox = False
    form.StartPosition = FormStartPosition.CenterScreen

    # Set the dialog size to match the image dimensions
    form.ClientSize = Size(img.Width, img.Height)

    # Create a PictureBox to host and animate the GIF
    pb = PictureBox()
    pb.Image = img
    pb.SizeMode = PictureBoxSizeMode.Zoom
    pb.Dock = DockStyle.Fill

    # Add the PictureBox to the form and display it modally
    form.Controls.Add(pb)
    form.ShowDialog()
else:
    print("GIF file not found in the script folder.")