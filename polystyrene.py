import win32com.client
import logging
import math

acad = win32com.client.Dispatch("AutoCAD.Application")
shell = win32com.client.Dispatch("WScript.Shell")  # windows scripts
doc = acad.ActiveDocument
logger = logging.getLogger(__name__)


def create_boundary():
    doc.Utility.Prompt(
        "Нажмите ENTER или правую кнопку мыши для образования контура\n(Важно: тип объекта должен быть полилиния!!!)\n")
    # doc.SendCommand("BOUNDARY") #for english version
    doc.SendCommand("КОНТУР")
    # shell.SendKeys("{ENTER}")  # press enter key


def get_selection(doc, text="Выберете объект"):
    doc.Utility.prompt(text)
    try:
        doc.SelectionSets.Item("SS1").Delete()
    except Exception:
        logger.debug('Delete selection failed')

    selection = doc.SelectionSets.Add('SS1')
    selection.SelectOnScreen()
    return selection


def get_width_and_height_of_rectangle(item):
    width = item.Coordinates[0] - item.Coordinates[2]
    height = item.Coordinates[1] - item.Coordinates[5]
    return [width, height]


def get_and_create_position_discription(item, scale):
    poly_type = {
        1: "ППТ-15-А-Р",
        2: "Эффективный утеплитель"
    }
    returned_dict = {
        "count": 1
    }
    width_height_list = get_width_and_height_of_rectangle(item)
    print(width_height_list)
    thikness = doc.Utility.GetInteger(
        "Введите толщину пакета утеплителя и нажмите Enter\n")
    returned_dict["dimensions"] = str(
        math.trunc(width_height_list[0] * scale)) + "x" + str(math.trunc(width_height_list[1] * scale)) + "x" + str(thikness)
    received_type = doc.Utility.GetInteger(
        "Выберите вид утеплителя и нажмите Enter\n(если ППТ-15-А-Р - введите 1, если Эффективный утеплитель - введите 2)\n")
    returned_dict["type"] = poly_type[received_type]
    returned_dict["volume"] = (
        item.Area * (scale ** 2) / 1000000) * (thikness / 1000)
    return returned_dict


def main():
    shell.AppActivate(acad.Caption)  # change focus to autocad
    scale = doc.Utility.GetInteger(
        "Введите масштаб\n(Пример: если масштаб 1:40, то введите 40)\n")
    object_collection = {}
    is_continued = 1
    while is_continued == 1:
        create_boundary()
        selection = get_selection(
            doc, text="Выберите образованный контур (не более одного!)")
        if selection.Count > 1 and selection.Item(0).ObjectName.lower() != "acdbpolyline":
            doc.Utility.Prompt(
                "Ошибка: выбрано недопустимое количество объектов или невеный тип объекта\n")
            continue
        item = selection.Item(0)
        item.Color = 1
        position = doc.Utility.GetInteger(
            "Введите позицию пакета утеплителя и нажмите Enter\n")
        if position in object_collection:
            object_collection[position]["count"] += 1
        else:
            object_collection[position] = get_and_create_position_discription(
                item, scale)
        is_continued = doc.Utility.GetInteger(
            "Для того чтобы продолжить введите '1', для завершения введите '0'\n(для ввода доступны только вышеуказанные числа, иначе будет ошибка)\n")

    print(object_collection)


main()
