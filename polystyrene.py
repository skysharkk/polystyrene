import win32com.client
import logging
import math
import pandas
import numpy as np
import array
import ctypes
import json

from pyautocad import Autocad

acad = Autocad(create_if_not_exists=True)
shell = win32com.client.Dispatch("WScript.Shell")  # windows scripts
logger = logging.getLogger(__name__)


def show_error_window(error_message, window_name=u"Ошибка"):
    ctypes.windll.user32.MessageBoxW(
        0, error_message, window_name, 0)


def write_to_excel(collection):
    df = pandas.DataFrame(np.array(collection), columns=[
        "Поз.", "Обозначение", "Наименование", "Кол.", "Объем ед. м3", "Примечание"])
    try:
        df.to_excel('./ppt_excel_template.xlsx',
                    sheet_name='Расскладка', index=False)
    except Exception:
        show_error_window(
            u"Ошибка при записи данных в Excel, возможно вы забыли закрыть рабочий файл.")
        logger.debug(
            "Ошибка при записи данных в Excel, возможно вы забыли закрыть рабочий файл.")


def round_half_up(n, decimals=0, int_result=True):
    multiplier = 10 ** decimals
    result = math.floor(n*multiplier + 0.5) / multiplier
    if int_result:
        return int(result)
    else:
        return result


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
    width = abs(round_half_up(
        (coordinates["max_x"] - coordinates["min_x"]) * scale))
    height = abs(round_half_up(
        (coordinates["max_y"] - coordinates["min_y"]) * scale))
    return [width, height]


def add_name_item_to_model(doc, coordinates_tuple, name, text_height=3):
    coordinates = get_coordinates_of_item(coordinates_tuple)
    x = ((coordinates["min_x"] + coordinates["max_x"]) / 2) - (text_height / 2)
    y = ((coordinates["min_y"] + coordinates["max_y"]) / 2) - (text_height / 2)
    insertion_point = array.array('d', [x, y, 0.0])
    doc.ModelSpace.AddText(name, insertion_point, text_height)


def compare_sizes(first_size_list, second_size_list):
    return set(first_size_list) == set(second_size_list)


def create_item_dict(item, scale, thikness, type, norm, number):
    START_AMOUNT = 1
    coordinates_tuple = item.Coordinates
    width_height_list = get_width_and_heights(coordinates_tuple, scale)
    dimensions = type + str(
        width_height_list[0]) + "x" + str(width_height_list[1]) + "x" + str(thikness)
    volume = round_half_up((
        item.Area * (scale ** 2) / 1000000) * (thikness / 1000), 2, False)
    return {
        "discription": [number, norm, dimensions, START_AMOUNT, volume, ""],
        "sizes": width_height_list + [thikness],
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
            item_sizes_list = get_width_and_heights(
                coordinates, scale) + [initial_data["thikness"]]
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
    with open("data.json", "w", encoding="utf8") as write_file:
        json.dump(items_data, write_file, ensure_ascii=False)
    write_to_excel(format_data(items_data))


main()
