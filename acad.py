from array import array
from pyautocad import Autocad
import win32com.client


class Acad:
    def __init__(self):
        self.acad = Autocad(create_if_not_exists=True)
        self.shell = win32com.client.Dispatch("WScript.Shell")
        self.doc = self.acad.doc

    def chanhge_focus(self):
        self.shell.AppActivate(self.acad.app.Caption)

    def get_point(self, message_text="Выберете точку"):
        self.chanhge_focus()
        base_point = array("d", [0, 0, 0])
        return self.doc.Utility.GetPoint(base_point, message_text)

    def draw_polyline(self, points, width):
        AMOUNT_OF_COORDINATES = 3
        polyline = self.doc.ModelSpace.AddPolyline(points)
        for index in range(int(len(points) / AMOUNT_OF_COORDINATES)):
            polyline.SetWidth(index, width, width)

    def draw_table(self, table_coordinates):
        self.draw_polyline(table_coordinates["title"], 0.6)
        for row in table_coordinates["row"]["coordinates"]:
            self.draw_polyline
