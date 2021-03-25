import win32com.client
import logging
import math
import pandas
import numpy as np
import array
from pyautocad import Autocad

acad = Autocad(create_if_not_exists=True)
shell = win32com.client.Dispatch("WScript.Shell")  # windows scripts
logger = logging.getLogger(__name__)


def write_to_excel(collection):
    df = pandas.DataFrame(np.array(collection), columns=[
        "Поз.", "Обозначение", "Наименование", "Кол.", "Объем ед. м3", "Примечание"])

    df.to_excel('./ppt_excel_template.xlsx',
                sheet_name='Расскладка', index=False)


def round_half_up(n, decimals=0):
    multiplier = 10 ** decimals
    return int(math.floor(n*multiplier + 0.5) / multiplier)


def get_selected(doc, text="Выберете объект"):
    doc.Utility.prompt(text)
    try:
        doc.SelectionSets.Item("SS1").Delete()
    except Exception:
        logger.debug('Delete selection failed')

    selected = doc.SelectionSets.Add("SS1")
    selected.SelectOnScreen()
    return selected


def get_initial_data(doc):
    poly_type = {
        1: {
            "type": "ППТ-15-А-Р ",
            "norm": "СТБ 1437"
        },
        2: {
            "type": "",
            "norm": "Эффективный утеплитель λ ≤ 0,034 Вт/(м·°C),"
        }
    }
    thikness = doc.Utility.GetInteger(
        "Введите толщину пакета утеплителя и нажмите Enter\n")
    received_type = doc.Utility.GetInteger(
        "Выберите вид утеплителя и нажмите Enter\n(если ППТ-15-А-Р - введите 1, если Эффективный утеплитель - введите 2)\n")
    p_type = poly_type[received_type]["type"]
    norm = poly_type[received_type]["norm"]
    selected = get_selected(
        doc, "Выберите" + poly_type[received_type]["type"] + "толщиной " + str(thikness))
    return {
        "selected": selected,
        "thikness": thikness,
        "type": p_type,
        "norm": norm,
    }


def get_coordinates_of_item(coordinates_tuple):
    x_coordinates_list = []
    y_coordinates_list = []
    for index in range(len(coordinates_tuple)):
        if index % 2 == 0:
            x_coordinates_list.append(coordinates_tuple[index])
        else:
            y_coordinates_list.append(coordinates_tuple[index])
    max_x = max(x_coordinates_list)
    min_x = min(x_coordinates_list)
    max_y = max(y_coordinates_list)
    min_y = min(y_coordinates_list)
    return {
        "min_x": min_x,
        "max_x": max_x,
        "min_y": min_y,
        "max_y": max_y,
    }


def get_width_and_heights(coordinates_tuple, scale):
    coordinates = get_coordinates_of_item(coordinates_tuple)
    width = round_half_up(
        (abs(coordinates["max_x"]) - abs(coordinates["min_x"])) * scale)
    height = round_half_up(
        (abs(coordinates["max_y"]) - abs(coordinates["min_y"])) * scale)
    return [width, height]


def add_name_item_to_model(doc, coordinates_tuple, name):
    TEXT_HEIGHT = 3
    coordinates = get_coordinates_of_item(coordinates_tuple)
    x = (coordinates["min_x"] + coordinates["max_x"]) / 2
    y = (coordinates["min_y"] + coordinates["max_y"]) / 2
    insertion_point = array.array('d', [x, y, 0.0])
    doc.ModelSpace.AddText(name, insertion_point, TEXT_HEIGHT)


def compare_sizes(first_size_list, second_size_list):
    return set(first_size_list) == set(second_size_list)


def create_item_dict(item, scale, thikness, type, norm, number):
    START_AMOUNT = 1
    coordinates_tuple = item.Coordinates
    width_height_list = get_width_and_heights(coordinates_tuple, scale)
    dimensions = type + str(
        width_height_list[0]) + "x" + str(width_height_list[1]) + "x" + str(thikness)
    volume = (
        item.Area * (scale ** 2) / 1000000) * (thikness / 1000)
    return {
        "discription": [number, norm, dimensions, START_AMOUNT, volume, ""],
        "sizes": width_height_list,
    }


def check_size_match(items_data, item_sizes_list):
    if len(items_data) == 0:
        return 0
    for item in items_data:
        if compare_sizes(items_data[item]["sizes"], item_sizes_list):
            return item
    return 0


def format_data(collection):
    result = []
    for name in collection:
        result.append(collection[name]["discription"])
    return result


def main():
    AMOUNT_INDEX = 3
    is_continued = 1
    items_data = {}
    doc = acad.doc
    shell.AppActivate(acad.app.Caption)
    scale = doc.Utility.GetInteger(
        "Введите масштаб\n(Пример: если масштаб 1:40, то введите 40)\n")
    while is_continued == 1:
        initial_data = get_initial_data(doc)
        items = initial_data["selected"]
        for index in range(items.Count):
            item = items.Item(index)
            coordinates = item.Coordinates
            item_sizes_list = get_width_and_heights(coordinates, scale)
            is_exist = check_size_match(items_data, item_sizes_list)
            if is_exist > 0:
                items_data[is_exist]["discription"][AMOUNT_INDEX] += 1
                add_name_item_to_model(doc, coordinates, str(is_exist))
            else:
                pos = len(items_data) + 1
                item_dict = create_item_dict(
                    item, scale, initial_data["thikness"], initial_data["type"], initial_data["norm"], pos)
                items_data[pos] = item_dict
                add_name_item_to_model(doc, coordinates, str(pos))
        is_continued = doc.Utility.GetInteger(
            "Для того чтобы продолжить введите '1', для завершения введите '0'\n(для ввода доступны только вышеуказанные числа, иначе будет ошибка)\n")

    write_to_excel(format_data(items_data))


# def main():
#     doc = acad.ActiveDocument
#     shell.AppActivate(acad.Caption)  # change focus to autocad
#     scale = doc.Utility.GetInteger(
#         "Введите масштаб\n(Пример: если масштаб 1:40, то введите 40)\n")
#     object_collection = []
#     is_continued = 1
#     is_exist = False
#     AMOUNT_LIST_POSITION = 3
#     while is_continued == 1:
#         selection = get_selected(
#             doc, text="Выберите образованный контур (не более одного!)")
#         if selection.Count > 1 and selection.Item(0).ObjectName.lower() != "acdbpolyline":
#             doc.Utility.Prompt(
#                 "Ошибка: выбрано недопустимое количество объектов или неверный тип объекта\n")
#             continue
#         item = selection.Item(0)
#         position = doc.Utility.GetInteger(
#             "Введите позицию пакета утеплителя и нажмите Enter\n")
#         for element in object_collection:
#             if element[0] == position:
#                 element[AMOUNT_LIST_POSITION] += 1
#                 is_exist = True
#                 break
#         if is_exist == False:
#             object_collection.append(
#                 get_and_create_position_discription(doc, item, scale, position))
#         else:
#             is_exist = False
#         is_continued = doc.Utility.GetInteger(
#             "Для того чтобы продолжить введите '1', для завершения введите '0'\n(для ввода доступны только вышеуказанные числа, иначе будет ошибка)\n")
#     write_to_excel(object_collection)
main()
