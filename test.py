import greedypacker
from pyautocad import Autocad
from array import array
from copy import copy


acad = Autocad(create_if_not_exists=True)
doc = acad.doc


def create_point(coordinates, size, coordinate_position):
    copy_coordinates = copy(coordinates)
    copy_coordinates[coordinate_position] += size
    return copy_coordinates


def create_array_of_double(list):
    return array("d", list)


def two_demention_list_to_list(list):
    formatted_list = []
    for item in list:
        formatted_list.extend(item)
    return formatted_list


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


M = greedypacker.BinManager(
    2000, 1000, pack_algo='maximal_rectangle', heuristic='bottom_left', rotation=True)

ITEM = greedypacker.Item(1320, 700)
ITEM4 = greedypacker.Item(200, 200)
ITEM2 = greedypacker.Item(1320, 700)
ITEM3 = greedypacker.Item(1320, 720)
ITEM5 = greedypacker.Item(1000, 670)

M.add_items(ITEM, ITEM2, ITEM3, ITEM4, ITEM5)

M.execute()

print(M.bins)

initial_x = 0
initial_y = 0

for p_bin in M.bins:
    draw_rectangle(doc, array('d', (initial_x, initial_y, 0)), 2000, 1000)
    for item in p_bin.items:
        print(item)
        draw_rectangle(doc, array('d', (item.x + initial_x, item.y + initial_y, 0)),
                       item.width, item.height)
    initial_y += 1000
