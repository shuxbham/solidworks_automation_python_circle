import win32com.client
import pythoncom
swApp = win32com.client.Dispatch("SldWorks.Application")
swApp.Visible=True

template=r"C:\ProgramData\SolidWorks\SOLIDWORKS 2024\templates\Part.prtdot"
Part=swApp.NewDocument(template,0,0,0)

boolstatus = Part.Extension.SelectByID2("Front Plane", "PLANE", 0, 0, 0, False, 0,pythoncom.Nothing, 0)
Part.SketchManager.InsertSketch(True)
Part.ClearSelection2(True)

skSegment = Part.SketchManager.CreateCircle(0, 0, 0, 0.05, 0, 0)
Part.ClearSelection2(True)
Part.SketchManager.InsertSketch(True)
Part.ClearSelection2(True)

