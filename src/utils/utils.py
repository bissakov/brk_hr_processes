import os
from time import sleep

import pandas as pd
from pywinauto import mouse
from pywinauto.win32structures import RECT

from src.data import Order, Process, BusinessTripOrder
from src.utils.colvir_utils import Colvir


def create_report(report_file_path: str):
    if not os.path.exists(report_file_path):
        df = pd.DataFrame(
            {
                "Дата": [],
                "Сотрудник": [],
                "Операция": [],
                "Номер приказа": [],
                "Статус": [],
            }
        )
        df.to_excel(report_file_path, index=False)


def update_report(
    order: Order,
    process: Process,
    operation: str,
    status: str,
):
    df = pd.read_excel(process.report_path)

    if not (
        (df["Дата"] == process.today)
        & (df["Сотрудник"] == order.employee_fullname)
        & (df["Операция"] == operation)
        & (df["Номер приказа"] == order.order_number)
    ).any():
        new_row = {
            "Дата": process.today,
            "Сотрудник": order.employee_fullname,
            "Операция": operation,
            "Номер приказа": order.order_number,
            "Статус": status,
        }
        df.loc[len(df)] = new_row
        df.to_excel(process.report_path, index=False)


def get_city_mappings(colvir: Colvir, order: BusinessTripOrder) -> None:
    """
    Тестовая функция для сбора маппингов
    :param colvir: Colvir
    :param order: BusinessTripOrder
    :return: None
    """
    colvir.choose_mode(mode="PRS")
    filter_win = colvir.utils.get_window(title="Фильтр")
    colvir.find_and_click_button(
        button=colvir.buttons.clear_form,
        window=filter_win,
        toolbar=filter_win["Static3"],
        target_button_name="Очистить фильтр",
    )

    filter_win["Edit8"].set_text("001")
    filter_win["Edit4"].set_text(order.employee_names[0])
    filter_win["Edit2"].set_text(order.employee_names[1])
    filter_win["OK"].click()

    personal_win = colvir.utils.get_window(title="Персонал")
    colvir.find_and_click_button(
        colvir.buttons.employee_orders,
        window=personal_win,
        toolbar=personal_win["Static4"],
        target_button_name="Приказы по сотруднику",
    )

    orders_win = colvir.utils.get_window(title="Приказы сотрудника")
    colvir.find_and_click_button(
        button=colvir.buttons.create_new_order,
        window=orders_win,
        toolbar=orders_win["Static4"],
        target_button_name="Создать новую запись (Ins)",
    )

    order_win = colvir.utils.get_window(title="Приказ")

    order_win["Edit18"].type_keys("ORD_TRP", pause=0.1)
    order_win["Edit18"].type_keys("{TAB}")
    sleep(0.5)
    if (error_win := colvir.app.window(title="Произошла ошибка")).exists():
        error_win.close()
        order_win["Edit38"].type_keys("{TAB}")
    sleep(1)

    order_win["Edit40"].type_keys(order.order_number, pause=0.1)

    sleep(1)

    order_win["Edit4"].click_input()
    order_win["Edit4"].type_keys("001", pause=0.2)
    order_win.type_keys("{TAB}", pause=1)
    order_win["Edit10"].click_input()
    order_win["Edit10"].type_keys("0975", pause=0.2)
    order_win.type_keys("{TAB}", pause=1)

    if colvir.buttons.cities_menu.x == -1 or colvir.buttons.cities_menu.y == -1:
        order_win.set_focus()
        rect: RECT = order_win["Edit28"].rectangle()
        mid_point = rect.mid_point()

        start_point = rect.right
        end_point = rect.right + 200

        x, y = rect.right, mid_point.y
        mouse.move(coords=(x, y))

        x_offset = 5

        i = 0
        while (
            not colvir.utils.persistent_win_exists(
                title_re="Страны и города.+", timeout=0.1
            )
            or x >= end_point
        ):
            x = start_point + i * 5
            mouse.click(button="left", coords=(x, y))
            i += 1

        colvir.buttons.cities_menu.x = x + x_offset
        colvir.buttons.cities_menu.y = y

        cities_win = colvir.app.window(title="Страны и города (командировки)")
        cities_win.close()

    mappings = {}

    order_win.set_focus()
    colvir.buttons.cities_menu.click()
    i = 0
    while i < 500:
        cities_win = colvir.utils.get_window(title="Страны и города (командировки)")

        if i > 0:
            cities_win.type_keys("{DOWN}")
        ok_button = cities_win["OK"]
        if ok_button.is_enabled():
            ok_button.click()
        else:
            cities_win.type_keys("{ENTER}")
            continue

        cities_win.close()

        fullname = order_win["Edit18"].window_text().strip()
        colvir_name = order_win["Edit28"].window_text().strip()
        mappings[fullname] = colvir_name

        print(
            order_win["Edit28"].window_text().strip(),
            "&&",
            order_win["Edit18"].window_text().strip(),
        )

        i += 1

        colvir.buttons.cities_menu.click()

    pass
