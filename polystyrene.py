import win32com.client
import logging
import math
import pandas
import numpy as np

acad = win32com.client.Dispatch("AutoCAD.Application")
shell = win32com.client.Dispatch("WScript.Shell")  # windows scripts
logger = logging.getLogger(__name__)


# def create_boundary():
#     doc.Utility.Prompt(
#         "Нажмите ENTER или правую кнопку мыши для образования контура\n(Важно: тип объекта должен быть полилиния!!!)\n")
#     # doc.SendCommand("BOUNDARY") #for english version
#     doc.SendCommand("КОНТУР")
#     # shell.SendKeys("{ENTER}")  # press enter key


def get_selection(doc, text="Выберете объект"):
    doc.Utility.prompt(text)
    try:
        doc.SelectionSets.Item("SS1").Delete()
    except Exception:
        logger.debug('Delete selection failed')

    selection = doc.SelectionSets.Add("SS1")
    selection.SelectOnScreen()
    return selection


def get_width_and_heights(tu, scale):
    x_coordinates_list = []
    y_coordinates_list = []
    for index in range(len(tu)):
        if index % 2 == 0:
            x_coordinates_list.append(abs(tu[index]))
        else:
            y_coordinates_list.append(abs(tu[index]))
    max_x = max(x_coordinates_list)
    min_x = min(x_coordinates_list)
    max_y = max(y_coordinates_list)
    min_y = min(y_coordinates_list)

    width = math.trunc((max_x - min_x) * scale)
    height = math.trunc((max_y - min_y) * scale)
    return [width, height]


def get_and_create_position_discription(doc, item, scale, number):
    poly_type = {
        1: {
            "type": "ППТ-15-А-Р-",
            "norm": "СТБ 1437"
        },
        2: {
            "type": "",
            "norm": "Эффективный утеплитель"
        }
    }
    START_AMOUNT = 1
    width_height_list = get_width_and_heights(item, scale)
    thikness = doc.Utility.GetInteger(
        "Введите толщину пакета утеплителя и нажмите Enter\n")
    received_type = doc.Utility.GetInteger(
        "Выберите вид утеплителя и нажмите Enter\n(если ППТ-15-А-Р - введите 1, если Эффективный утеплитель - введите 2)\n")
    dimensions = poly_type[received_type]["type"] + str(
        width_height_list[0]) + "x" + str(width_height_list[1]) + "x" + str(thikness)
    volume = (
        item.Area * (scale ** 2) / 1000000) * (thikness / 1000)
    return [number, poly_type[received_type]["norm"], dimensions, START_AMOUNT, volume, ""]


def main():
    doc = acad.ActiveDocument
    shell.AppActivate(acad.Caption)  # change focus to autocad
    scale = doc.Utility.GetInteger(
        "Введите масштаб\n(Пример: если масштаб 1:40, то введите 40)\n")
    object_collection = []
    is_continued = 1
    is_exist = False
    AMOUNT_LIST_POSITION = 3
    while is_continued == 1:
        selection = get_selection(
            doc, text="Выберите образованный контур (не более одного!)")
        if selection.Count > 1 and selection.Item(0).ObjectName.lower() != "acdbpolyline":
            doc.Utility.Prompt(
                "Ошибка: выбрано недопустимое количество объектов или неверный тип объекта\n")
            continue
        item = selection.Item(0)

        position = doc.Utility.GetInteger(
            "Введите позицию пакета утеплителя и нажмите Enter\n")
        for element in object_collection:
            if element[0] == position:
                element[AMOUNT_LIST_POSITION] += 1
                is_exist = True
                break

        if is_exist == False:
            object_collection.append(
                get_and_create_position_discription(doc, item, scale, position))
        else:
            is_exist = False

        is_continued = doc.Utility.GetInteger(
            "Для того чтобы продолжить введите '1', для завершения введите '0'\n(для ввода доступны только вышеуказанные числа, иначе будет ошибка)\n")
    print(object_collection)
    df = pandas.DataFrame(np.array(object_collection), columns=[
                          "Поз.", "Обозначение", "Наименование", "Кол.", "Объем ед. м3", "Примечание"])
    df.to_excel('./ppt_excel_template.xlsx',
                sheet_name='Расскладка', index=False)


main()
