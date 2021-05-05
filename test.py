import greedypacker
from pyautocad import Autocad
from array import array
from copy import copy
from pprint import pprint


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
    20, 10, pack_algo='maximal_rectangle', heuristic='bottom_left', rotation=True)

ITEM = greedypacker.Item(5, 6)
ITEM4 = greedypacker.Item(3, 2)
# ITEM2 = greedypacker.Item(1320, 700)
# ITEM3 = greedypacker.Item(1320, 720)
# ITEM5 = greedypacker.Item(1000, 670)

M.add_items(ITEM, ITEM4)

M.execute()


def draw_packs():
    initial_x = 0
    initial_y = 0
    for p_bin in M.bins:
        draw_rectangle(doc, array('d', (initial_x, initial_y, 0)), 20, 10)
        for item in p_bin.items:
            draw_rectangle(doc, array('d', (item.x + initial_x, item.y + initial_y, 0)),
                           item.width, item.height)
        initial_y += 10


def generate_rectangle_matrix(width, height):
    width_list = [0] * width
    result = []
    for _ in range(height):
        result.append(copy(width_list))
    return result


def draw_in_matrix(container_matrix, elements):
    pprint(elements)
    for el in elements:
        for i in range(el.y, el.height):
            print(el.x, el.width)
            container_matrix[i][el.x: el.x + el.width] = [1] * (el.width)
    return container_matrix


mat = draw_in_matrix(generate_rectangle_matrix(20, 10), M.items)
pprint(mat)
draw_packs()
