from array import array
from acad import Acad
from copy import copy


class Table:
    def __init__(self, initial_point, number_of_rows, scale):
        self.number_of_rows = number_of_rows
        self.initial_point = initial_point
        self.scale = scale
        self.__width = 7400 / self.scale
        self.__title_row_height = 600 / self.scale
        self.__row_height = 320 / self.scale
        self.__haight = self.__title_row_height + \
            (self.__row_height * self.number_of_rows)
        self.__separator_points = (0,
                                   600 / self.scale, 2400 / self.scale, 2600 / self.scale, 400 / self.scale, 600 / self.scale, 800 / self.scale)
        self.table_coordinates = {
            "title": [],
            "row": {
                "coordinates": [],
                "first_insertion_point": 0,
            },

            "separators": [],
        }

    @staticmethod
    def __create_point(coordinates, size, coordinate_position):
        copied_coordinates = copy(coordinates)
        copied_coordinates[coordinate_position] += size
        return copied_coordinates

    @staticmethod
    def __two_demention_list_to_list(list):
        formatted_list = []
        for item in list:
            formatted_list.extend(item)
        return formatted_list

    def __create_rectangle(self, width, height):
        sizes = [width, -height, -width, height]
        points = [list(self.initial_point)]
        AMOUNT_ITERATES = 4
        for i in range(AMOUNT_ITERATES):
            if i % 2 == 0:
                new_point = Table.__create_point(points[i], sizes[i], 0)
                points.append(new_point)
            else:
                new_point = Table.__create_point(points[i], sizes[i], 1)
                points.append(new_point)
        points = array("d", Table.__two_demention_list_to_list(points))
        return points

    def __create_title_row(self):
        return self.__create_rectangle(self.__width, self.__title_row_height)

    def __create_row(self):
        return self.__create_rectangle(self.__width, self.__row_height)

    def __create_separators(self):
        last_point = list(self.initial_point)
        result = []
        for item in self.__separator_points:
            start = last_point
            start[0] += item
            end = copy(start)
            end[1] += self.__haight
            result.append(array("d", [*start, *end]))
        return result

    def create_table(self):
        self.table_coordinates["title"] = self.__create_title_row()
        for i in range(self.number_of_rows):
            self.table_coordinates["row"]["coordinates"].append(
                self.__create_row())
        self.table_coordinates["separators"] = self.__create_separators()
        print(self.table_coordinates)


a = Acad()

table = Table(a.get_point(), 4, 1)


def main():
    table.create_table()


main()
