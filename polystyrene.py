import win32com.client
import win32gui
import win32process

acad = win32com.client.Dispatch("AutoCAD.Application")

# shell = win32com.client.Dispatch("WScript.Shell") #windows scripts
doc = acad.ActiveDocument

# shell.AppActivate(acad.Caption) #change focus to autocad
# shell.SendKeys("{ENTER}") #press enter key

doc.Utility.Prompt("Press enter!")
doc.SendCommand("BOUNDARY")

# entity = doc.Utility.GetEntity()

# print(entity[0].Area)
