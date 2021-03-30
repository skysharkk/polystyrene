import ctypes
import pandas
import logging
import numpy as np
import array

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


def get_selected(doc, text="Выберете объект"):
    doc.Utility.prompt(text)
    try:
        doc.SelectionSets.Item("SS1").Delete()
    except Exception:
        logger.debug('Delete selection failed')

    selected = doc.SelectionSets.Add("SS1")
    selected.SelectOnScreen()
    return selected


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


def add_name_item_to_model(doc, coordinates_tuple, name, text_height=3):
    coordinates = get_coordinates_of_item(coordinates_tuple)
    x = ((coordinates["min_x"] + coordinates["max_x"]) / 2) - (text_height / 2)
    y = ((coordinates["min_y"] + coordinates["max_y"]) / 2) - (text_height / 2)
    insertion_point = array.array('d', [x, y, 0.0])
    doc.ModelSpace.AddText(name, insertion_point, text_height)
