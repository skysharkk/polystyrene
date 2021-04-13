from pyautocad import Autocad
from array import array
from comtypes.client import Constants

acad = Autocad(create_if_not_exists=True)
doc = acad.doc
ms = acad.model


c_s = [600, 2400, 2600, 400, 600, 800]
r_h = [600, 320, 320, 320, 320, 320]


class AcadTable:
    def __init__(
        self,
        initial_point,
        scale,
        column_width,
        row_height,
        acad
    ):
        self.initial_point = initial_point
        self.scale = scale
        self.column_width = column_width
        self.row_height = row_height
        self.__model_space = acad.doc.ModelSpace
        self.__constansts = Constants(acad.app)
        self.__table = None

    def apply_scale(self, sizes_list):
        for index in range(len(sizes_list)):
            sizes_list[index] /= self.scale

    def __add_table_to_ms(self):
        self.__table = self.__model_space.AddTable(
            self.initial_point,
            len(self.row_height) + 1,
            len(self.column_width),
            1,
            1
        )
        self.__table.DeleteRows(0, 1)

    def __change_table_sizes(self):
        self.apply_scale(self.row_height)
        self.apply_scale(self.column_width)

        for index in range(len(self.column_width)):
            self.__table.SetColumnWidth(index, self.column_width[index])

        for index in range(len(self.row_height)):
            self.__table.SetRowHeight(index, self.row_height[index])

    def __set_line_weight(self, position, aligne, weight):
        self.__table.SetGridLineWeight2(
            position[0], position[1], aligne, weight)

    def set_vertical_line_weight(self, position, weight):
        self.__set_line_weight(
            position,
            self.__constansts.acVertLeft + self.__constansts.acVertRight,
            weight
        )

    def set_vert_and_horz_line_weight(self, position, weight):
        self.__set_line_weight(
            position,
            self.__constansts.acVertLeft +
            self.__constansts.acVertRight + self.__constansts.acHorzBottom +
            self.__constansts.acHorzTop,
            weight
        )

    def create_table(self, line_weight):
        self.__add_table_to_ms()
        self.__change_table_sizes()
        for col_index in range(len(self.column_width)):
            for row_index in range(len(self.row_height)):
                print(row_index, col_index)
                if row_index == 0:
                    self.set_vert_and_horz_line_weight(
                        [row_index, col_index], line_weight)
                self.set_vertical_line_weight(
                    [row_index, col_index], line_weight)


table = AcadTable(array("d", [0, 0, 0]), 40, c_s, r_h, acad)

table.create_table(60)
constants = Constants(acad.app)
