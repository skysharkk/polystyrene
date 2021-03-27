from array import array
import json
from pyautocad import Autocad

acad = Autocad(create_if_not_exists=True)
doc = acad.doc


# scale = doc.Utility.GetInteger(
#     "Введите масштаб\n(Пример: если масштаб 1:40, то введите 40)\n")


def draw_rectangle(doc, initial_point, width, height):
    ms = doc.ModelSpace
    points = [create_array_of_double(initial_point)]
    AMOUNT_ITERATES = 4
    for i in range(AMOUNT_ITERATES):
        if i % 2 == 0:
            x = points[i][0] + height

        else:
            y = points[i][1] + width


def create_array_of_double(list):
    return array("d", list)


def format_data(data):
    result = {}
    for item in data:
        thikness = data[item]["sizes"][2]
        if result.get(thikness):
            result[thikness].append(data[item]["sizes"])
        else:
            result[thikness] = [data[item]["sizes"]]
    print(result)
    return result


def get_point():
    base_point = create_array_of_double([0, 0, 0])
    return doc.Utility.GetPoint(base_point, "Выберете точку")


def get_json_data():
    with open("data.json", "r") as load_file:
        data = json.load(load_file)
    return data
