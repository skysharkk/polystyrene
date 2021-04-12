from array import array
from comtypes.automation import VARIANT
from ctypes import byref


def get_bounding_box(entity):
    min_point = VARIANT(array("d", array("d", [0, 0, 0])))
    max_point = VARIANT(array("d", array("d", [0, 0, 0])))
    ref_min_point = byref(min_point)
    ref_max_point = byref(max_point)
    entity.GetBoundingBox(ref_min_point, ref_max_point)
    return [array("d", list(*min_point)), array("d", list(*max_point))]


def get_mtext_height(model_space, width, text_height, text):
    text = model_space.AddMText(array("d", array("d", [0, 0, 0])), width, text)
    text.Height = text_height
    bounding_box_coordinates = get_bounding_box(text)
    text.Delete()
    return abs(bounding_box_coordinates[1][1] - bounding_box_coordinates[0][1])
