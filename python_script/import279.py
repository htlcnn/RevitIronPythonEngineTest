import clr
clr.AddReference("RevitAPIUI")
from Autodesk.Revit.UI import TaskDialog
import os
dll_path = os.path.join(os.path.dirname(__file__), "htllib279.dll")
clr.AddReferenceToFileAndPath(dll_path)
import htl
TaskDialog.Show("title1", str(htl.clipboard))