from pyautocad import Autocad
from p_utility.p_modules import get_selected


def create_str(returnable, irrevocable, percentage, poly_type=None):
    str_template = f"Возвратные остатки - {returnable} м3, отходы не используемые в производстве - {irrevocable} м3 ({percentage}%)"
    if poly_type:
        str_template += f"-{poly_type}"
    return str_template


def get_balance():
    balance_type_dict = {
        1: ["returnable", "возвратные"],
        0: ["irrevocable", "невозвратные"]
    }
    poly_type = {
        1: {
            "type": "для ППТ-15",
        },
        2: {
            "type": "для Эфф. утеп.",
        }
    }
    result = {
        "returnable": {
            1: [],
            2: []
        },
        "irrevocable": {
            1: [],
            2: []
        }
    }

    acad = Autocad(create_if_not_exists=True)
    doc = acad.doc
    is_continued = 1
    while is_continued == 1:
        balance_type = doc.Utility.GetInteger(
            "Выберите вид остатков и нажмите Enter\n(если возвратные - введите 1, если невозвратные - введите 2)\n")
        received_type = doc.Utility.GetInteger(
            "Выберите вид утеплителя и нажмите Enter\n(если ППТ-15-А-Р - введите 1, если Эффективный утеплитель - введите 2)\n")
        thikness = doc.Utility.GetInteger(
            "Введите толщину пакета утеплителя и нажмите Enter\n")
        chosen_type = poly_type[received_type]["type"]
        selected = get_selected(
            doc, f"Выберете возвратные остатки {chosen_type}, толщиной {thikness}")
        for index in range(selected.Count):
            item = selected.Item(index)
            result[balance_type_dict[balance_type][0]][received_type].append(
                (item.Area / 1000000) * (thikness / 1000))
        is_continued = doc.Utility.GetInteger(
            "Для того чтобы продолжить введите '1', для завершения введите '0'\n(для ввода доступны только вышеуказанные числа, иначе будет ошибка)\n")
    return result


def main():
    data = get_balance()
    for el in data:
        item = data[el]
        for i in item:
            item[i] = sum(item[i])
    print(data)
    return data


main()
