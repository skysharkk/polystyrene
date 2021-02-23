import win32com.client
import win32gui
import win32process
import logging

acad = win32com.client.Dispatch("AutoCAD.Application")
shell = win32com.client.Dispatch("WScript.Shell")  # windows scripts
doc = acad.ActiveDocument
logger = logging.getLogger(__name__)


# def create_boundary():
#     doc.Utility.Prompt(
#         "Нажмите ENTER или правую кнопку мыши для образования контура\n")
#     # doc.SendCommand("BOUNDARY") #for english version
#     doc.SendCommand("КОНТУР")
#     # shell.SendKeys("{ENTER}")  # press enter key


def get_selection(doc, text="Выберете объект"):
    doc.Utility.prompt(text)
    try:
        doc.SelectionSets.Item("SS1").Delete()
    except Exception:
        logger.debug('Delete selection failed')

    selection = doc.SelectionSets.Add('SS1')
    selection.SelectOnScreen()
    return selection


def main():
    shell.AppActivate(acad.Caption)  # change focus to autocad
    # create_boundary()
    doc.Utility.Prompt("Выберите образованный контур\n")
    selection = get_selection(doc, text="Выберете объект")
    # selection = doc.SelectionSets.Item('SS1')
    # entity = doc.Utility.GetEntity()
    # print(entity[0].Area)
    print(selection.Item(0).Area)


main()
