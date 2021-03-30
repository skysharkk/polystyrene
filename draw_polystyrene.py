from array import array
import json
from copy import copy
from pyautocad import Autocad
from p_utility.p_modules import add_name_item_to_model

acad = Autocad(create_if_not_exists=True)
doc = acad.doc


# scale = doc.Utility.GetInteger(
#     "Введите масштаб\n(Пример: если масштаб 1:40, то введите 40)\n")

def create_array_of_double(list):
    return array("d", list)


def two_demention_list_to_list(list):
    formatted_list = []
    for item in list:
        formatted_list.extend(item)
    return formatted_list


def create_point(coordinates, size, coordinate_position):
    copy_coordinates = copy(coordinates)
    copy_coordinates[coordinate_position] += size
    return copy_coordinates


def format_data(data):
    result = {}
    for item in data:
        thikness = data[item]["sizes"][2]
        if result.get(thikness):
            result[thikness].append(
                [data[item]["sizes"], data[item]["discription"][3]])
        else:
            result[thikness] = [
                [data[item]["sizes"], data[item]["discription"][3]]]
    return result


def get_point(message_text="Выберете точку"):
    base_point = create_array_of_double([0, 0, 0])
    return doc.Utility.GetPoint(base_point, message_text)


def get_json_data():
    with open("data.json", "r") as load_file:
        data = json.load(load_file)
    return data


def draw_rectangle(doc, initial_point, width, height):
    ms = doc.ModelSpace
    points = [list(initial_point)]
    sizes = [width, -height, -width, height]
    AMOUNT_ITERATES = 4
    for i in range(AMOUNT_ITERATES):
        if i % 2 == 0:
            new_point = create_point(points[i], sizes[i], 0)
            points.append(new_point)
        else:
            new_point = create_point(points[i], sizes[i], 1)
            points.append(new_point)
    points = create_array_of_double(two_demention_list_to_list(points))
    ms.AddPolyline(points)
    return points


def format_coordinates_from_name(coordinates):
    result = []
    coordinates_list = coordinates.tolist()
    for i in range(len(coordinates_list)):
        if (i + 1) % 3 != 0:
            result.append(coordinates_list[i])
    return result


def main(data):
    initial_point = get_point()
    for item in data:
        for el in data[item]:
            for i in range(el[1]):
                item_coordinates = draw_rectangle(
                    doc, initial_point, el[0][0], el[0][1])
                add_name_item_to_model(
                    doc, format_coordinates_from_name(item_coordinates), item, 100)
                initial_point = (
                    item_coordinates[3], item_coordinates[4], item_coordinates[5])


main(format_data(get_json_data()))
