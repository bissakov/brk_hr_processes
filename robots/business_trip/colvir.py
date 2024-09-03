import json
import os
import pickle
from time import sleep
from typing import List, Optional

import pywinauto.timings
from pywinauto import mouse
from pywinauto.win32structures import RECT

from robots.data import ColvirInfo, BusinessTripOrder, Buttons, Button
from robots.utils.colvir_utils import (
    get_window,
    Colvir,
    choose_mode,
    persistent_win_exists,
    change_oper_day,
    save_excel,
)
from robots.utils.utils import (
    kill_all_processes,
    does_order_exist,
    create_report,
    update_report,
)
from robots.utils.wiggle import wiggle_mouse


def get_colvir_city_code(trip_place: str, work_folder: str) -> Optional[str]:
    with open(os.path.join(work_folder, "cities.json"), "r", encoding="utf-8") as f:
        cities = json.load(f)

    city_bpm = trip_place.replace("город ", "").replace("г. ", "")
    city_bpm = city_bpm.split(",")[0]
    city_colvir = cities.get(city_bpm)
    if not city_colvir:
        return None
    city_colvir = city_colvir.replace(f".{city_bpm}", "")

    return city_colvir


def get_city_mappings(
    app: pywinauto.Application, order: BusinessTripOrder, buttons: Buttons
) -> None:
    choose_mode(app=app, mode="PRS")
    filter_win = get_window(app=app, title="Фильтр")
    buttons.clear_form.find_and_click_button(
        app=app,
        window=filter_win,
        toolbar=filter_win["Static3"],
        target_button_name="Очистить фильтр",
    )

    filter_win["Edit8"].set_text("001")
    filter_win["Edit4"].set_text(order.employee_names[0])
    filter_win["Edit2"].set_text(order.employee_names[1])
    filter_win["OK"].click()

    personal_win = get_window(app=app, title="Персонал")
    buttons.employee_orders.find_and_click_button(
        app=app,
        window=personal_win,
        toolbar=personal_win["Static4"],
        target_button_name="Приказы по сотруднику",
    )

    orders_win = get_window(app=app, title="Приказы сотрудника")
    buttons.create_new_order.find_and_click_button(
        app=app,
        window=orders_win,
        toolbar=orders_win["Static4"],
        target_button_name="Создать новую запись (Ins)",
    )

    order_win = get_window(app=app, title="Приказ")

    order_win["Edit18"].type_keys("ORD_TRP", pause=0.1)
    order_win["Edit18"].type_keys("{TAB}")
    sleep(0.5)
    if (error_win := app.window(title="Произошла ошибка")).exists():
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

    if buttons.cities_menu.x == -1 or buttons.cities_menu.y == -1:
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
            not persistent_win_exists(
                app=app, title_re="Страны и города.+", timeout=0.1
            )
            or x >= end_point
        ):
            x = start_point + i * 5
            mouse.click(button="left", coords=(x, y))
            i += 1

        buttons.cities_menu.x = x + x_offset
        buttons.cities_menu.y = y

        cities_win = app.window(title="Страны и города (командировки)")
        cities_win.close()

    mappings = {}

    order_win.set_focus()
    buttons.cities_menu.click()
    i = 0
    while i < 500:
        cities_win = get_window(app=app, title="Страны и города (командировки)")

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

        buttons.cities_menu.click()

    pass


def run(
    colvir_info: ColvirInfo, today: str, report_file_path: str, orders_pickle_path: str
):
    with open(orders_pickle_path, "rb") as f:
        orders: List[BusinessTripOrder] = pickle.load(f)

    assert all(isinstance(order, BusinessTripOrder) for order in orders)

    report_folder = os.path.dirname(report_file_path)
    create_report(report_file_path)

    kill_all_processes(proc_name="COLVIR")

    buttons = Buttons()

    colvir = Colvir(colvir_info=colvir_info)
    app = colvir.app

    # get_city_mappings(app=app, order=orders[0], buttons=buttons)  #  Сбор маппингов (не запускать просто так)

    for i, order in enumerate(orders):
        choose_mode(app=app, mode="TOPERD")
        change_oper_day(app=app, start_date=order.start_date)

        choose_mode(app=app, mode="PRS")
        filter_win = get_window(app=app, title="Фильтр")
        buttons.clear_form.find_and_click_button(
            app=app,
            window=filter_win,
            toolbar=filter_win["Static3"],
            target_button_name="Очистить фильтр",
        )

        filter_win["Edit8"].set_text("001")
        filter_win["Edit4"].set_text(order.employee_names[0])
        filter_win["Edit2"].set_text(order.employee_names[1])
        filter_win["OK"].click()

        sleep(1)
        confirm_win = app.window(title="Подтверждение")
        if confirm_win.exists():
            update_report(
                person_name=order.employee_fullname,
                order_number=order.order_number,
                report_file_path=report_file_path,
                today=today,
                operation="Создание приказа",
                status="Приказ не найден",
            )
            confirm_win.close()
            filter_win.close()
            personal_win = app.window(title="Персонал")
            if personal_win.exists():
                personal_win.close()
            continue

        personal_win = get_window(app=app, title="Персонал")
        buttons.employee_orders.find_and_click_button(
            app=app,
            window=personal_win,
            toolbar=personal_win["Static4"],
            target_button_name="Приказы по сотруднику",
        )

        orders_win = get_window(app=app, title="Приказы сотрудника")
        orders_win.menu_select("#4->#4->#1")

        if does_order_exist(
            orders_file_path=save_excel(app=app, work_folder=report_folder),
            order_type="Приказ о отправке работника в командировку",
            order_number=order.order_number,
        ):
            orders_win.close()
            personal_win.close()

            update_report(
                person_name=order.employee_fullname,
                order_number=order.order_number,
                report_file_path=report_file_path,
                today=today,
                operation="Создание приказа",
                status="Приказ уже создан",
            )
            continue

        personal_win.set_focus()
        sleep(1)
        personal_win.type_keys("{ENTER}")

        employee_card = get_window(app=app, title="Карточка сотрудника")
        employee_status = employee_card["Edit30"].window_text().strip()
        print(order.employee_fullname, employee_status)
        if (
            employee_status == "Уволен"
            or employee_status == "В командировке"
            or employee_status == "В отпуске"
        ):
            employee_card.close()
            orders_win.close()
            personal_win.close()
            update_report(
                person_name=order.employee_fullname,
                order_number=order.order_number,
                report_file_path=report_file_path,
                today=today,
                operation="Создание приказа",
                status=f"Невозможно создать приказ для сотрудника "
                f'со статусом "{employee_status}"',
            )
            continue

        branch_num = employee_card["Edit60"].window_text()
        tab_num = employee_card["Edit34"].window_text()
        employee_card.close()

        orders_win.set_focus()
        sleep(1)
        buttons.create_new_order.find_and_click_button(
            app=app,
            window=orders_win,
            toolbar=orders_win["Static4"],
            target_button_name="Создать новую запись (Ins)",
        )

        order_win = get_window(app=app, title="Приказ")

        order_win["Edit18"].type_keys("ORD_TRP", pause=0.1)
        order_win["Edit18"].type_keys("{TAB}")
        sleep(0.5)
        if (error_win := app.window(title="Произошла ошибка")).exists():
            error_win.close()
            order_win["Edit38"].type_keys("{TAB}")

        sleep(1)

        order_win["Edit40"].type_keys(order.order_number, pause=0.1)

        sleep(1)

        order_win["Edit4"].click_input()
        order_win["Edit4"].type_keys(branch_num, pause=0.2)
        order_win.type_keys("{TAB}", pause=1)
        order_win["Edit10"].click_input()
        order_win["Edit10"].type_keys(tab_num, pause=0.2)
        order_win.type_keys("{TAB}", pause=1)

        if not order_win.has_focus():
            order_win.set_focus()

        order_win["Edit22"].click_input()
        order_win["Edit22"].set_text(order.start_date.short)

        order_win["Edit24"].click_input()
        order_win["Edit24"].set_text(order.end_date.short)

        city_code = get_colvir_city_code(
            trip_place=order.trip_place, work_folder=report_folder
        )

        if city_code is None:
            update_report(
                person_name=order.employee_fullname,
                order_number=order.order_number,
                report_file_path=report_file_path,
                today=today,
                operation="Создание приказа",
                status=f"Не удалось заполнить приказ. Требуется проверка специалистом. "
                f"Неизвестный город/местоположение - {order.trip_place}",
            )
            order_win.type_keys("{ESC}")

            confirm_win = get_window(app=app, title="Подтверждение")
            confirm_win["&Нет"].click()

            orders_win.close()
            personal_win.close()
            continue

        order_win["Edit28"].type_keys(city_code, pause=0.2)
        order_win["Edit28"].click_input()
        order_win.type_keys("{TAB}", pause=1)

        order_win["Edit16"].type_keys(order.trip_target, pause=0.1, with_spaces=True)
        order_win["Edit16"].click_input()
        order_win.type_keys("{TAB}", pause=1)

        buttons.order_save.find_and_click_button(
            app=app,
            window=order_win,
            toolbar=order_win["Static3"],
            target_button_name="Сохранить изменения (PgDn)",
        )

        orders_win.wait(wait_for="active enabled")

        buttons.operations_list.find_and_click_button(
            app=app,
            window=orders_win,
            toolbar=orders_win["Static4"],
            target_button_name="Выполнить операцию",
        )

        sleep(0.5)
        buttons.operation = Button(
            buttons.operations_list.x,
            buttons.operations_list.y + 30,
        )
        buttons.operation.check_and_click(app=app, target_button_name="Регистрация")

        registration_win = get_window(app=app, title="Подтверждение")
        registration_win["&Да"].click()
        sleep(2)
        confirm_win = app.window(title="Подтверждение")
        if confirm_win.exists():
            confirm_win.close()
        sleep(1)
        dossier_win = app.window(title="Досье сотрудника")
        if dossier_win.exists():
            dossier_win.close()

        wiggle_mouse(duration=3)

        buttons.operations_list.click()
        sleep(1)
        buttons.operation.click()
        confirm_win = get_window(app=app, title="Подтверждение")
        confirm_win["&Да"].click()
        wiggle_mouse(duration=3)

        buttons.operations_list.click()
        sleep(1)
        buttons.operation.click()
        confirm_win = get_window(app=app, title="Подтверждение")
        confirm_win["&Да"].click()
        wiggle_mouse(duration=3)

        command_win = app.window(title="Распоряжение на командировку")
        if command_win.exists():
            command_win.close()

        error_win = app.window(title="Произошла ошибка")
        if error_win.exists():
            error_msg = error_win.child_window(class_name="Edit").window_text()
            update_report(
                person_name=order.employee_fullname,
                order_number=order.order_number,
                report_file_path=report_file_path,
                today=today,
                operation="Создание приказа",
                status=f"Не удалось ИСПОЛНИТЬ приказ. Требуется проверка специалистом. "
                f'Текст ошибки - "{error_msg}"',
            )
            error_win.close()
            orders_win.close()
            personal_win.close()
            continue

        if order.deputy_fullname is None:
            update_report(
                person_name=order.employee_fullname,
                order_number=order.order_number,
                report_file_path=report_file_path,
                today=today,
                operation="Создание приказа",
                status="Приказ создан",
            )
            orders_win.close()
            personal_win.close()
            continue

        pass

        update_report(
            person_name=order.deputy_fullname,
            order_number=order.order_number,
            report_file_path=report_file_path,
            today=today,
            operation="Создание приказа",
            status=f"Приказ создан. Доплата за на период командировки сотрудника {order.employee_fullname}",
        )
        orders_win.close()
        personal_win.close()

    pass
