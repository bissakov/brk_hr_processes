import os
import pickle
from time import sleep
from typing import List

from robots.data import ColvirInfo, VacationOrder, Buttons, Button, Process
from robots.utils.colvir_utils import (
    get_window,
    Colvir,
    choose_mode,
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


def run(colvir_info: ColvirInfo, process: Process, buttons: Buttons) -> None:
    with open(process.pickle_path, "rb") as f:
        orders: List[VacationOrder] = pickle.load(f)

    assert all(isinstance(order, VacationOrder) for order in orders)

    report_folder = os.path.dirname(process.report_path)
    create_report(process.report_path)

    kill_all_processes(proc_name="COLVIR")

    colvir = Colvir(colvir_info=colvir_info)
    app = colvir.app

    for i, order in enumerate(orders):
        if i == 0:
            continue
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
                order=order,
                process=process,
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
            order_type="Приказ на отпуск",
            order_number=order.order_number,
        ):
            orders_win.close()
            personal_win.close()

            update_report(
                order=order,
                process=process,
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

        if employee_status == "Уволен":
            employee_card.close()
            orders_win.close()
            personal_win.close()
            update_report(
                order=order,
                process=process,
                operation="Создание приказа",
                status=f"Невозможно создать приказ для сотрудника "
                f'со статусом "{employee_status}"',
            )
            continue

        branch_num = employee_card["Edit60"].window_text()
        tab_num = employee_card["Edit34"].window_text()
        employee_card.close()

        if employee_status == "В командировке":
            orders_win.wait(wait_for="active enabled")
            personal_win.set_focus()

            buttons.operations_list.find_and_click_button(
                app=app,
                window=personal_win,
                toolbar=personal_win["Static4"],
                target_button_name="Выполнить операцию",
            )

            sleep(0.5)
            buttons.operation = Button(
                buttons.operations_list.x,
                buttons.operations_list.y + 30,
            )
            buttons.operation.check_and_click(
                app=app, target_button_name="Возврат из командировки"
            )

            confirm_win = get_window(app=app, title="Подтверждение")
            confirm_win["&Да"].click()
            wiggle_mouse(duration=2)

            return_win = get_window(app=app, title="Возврат из командировки")
            return_win["Принять"].click()

        if employee_status == "В отпуске":
            pass

        if employee_status != "Работающий":
            pass

        orders_win.set_focus()
        sleep(1)
        buttons.create_new_order.find_and_click_button(
            app=app,
            window=orders_win,
            toolbar=orders_win["Static4"],
            target_button_name="Создать новую запись (Ins)",
        )

        order_win = get_window(app=app, title="Приказ")

        order_win["Edit18"].type_keys("ORD_HOL", pause=0.1)
        order_win["Edit18"].type_keys("{TAB}")
        sleep(0.5)
        if (error_win := app.window(title="Произошла ошибка")).exists():
            error_win.close()
            order_win["Edit38"].type_keys("{TAB}")

        sleep(1)
        order_win["Edit48"].type_keys(order.order_number, pause=0.1)
        sleep(1)

        order_win["Edit4"].click_input()
        order_win["Edit4"].type_keys(branch_num, pause=0.2)
        order_win.type_keys("{TAB}", pause=1)
        order_win["Edit10"].click_input()
        order_win["Edit10"].type_keys(tab_num, pause=0.2)
        order_win.type_keys("{TAB}", pause=1)

        if not order_win.has_focus():
            order_win.set_focus()

        order_win["Edit28"].click_input()
        order_win["Edit28"].type_keys(order.order_type, pause=0.2)
        order_win.type_keys("{TAB}", pause=1)

        if not order_win.has_focus():
            order_win.set_focus()

        order_win["Edit30"].click_input()
        order_win["Edit30"].set_text(order.start_date.short)

        order_win["Edit32"].click_input()
        order_win["Edit32"].set_text(order.end_date.short)

        buttons.order_save.find_and_click_button(
            app=app,
            window=order_win,
            toolbar=order_win["Static3"],
            target_button_name="Сохранить изменения (PgDn)",
        )

        orders_win.wait(wait_for="active enabled")

        buttons.operations_list = Button()
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

        wiggle_mouse(duration=2)

        buttons.operations_list.click()
        sleep(1)
        buttons.operation.click()
        confirm_win = get_window(app=app, title="Подтверждение")
        confirm_win["&Да"].click()
        wiggle_mouse(duration=2)

        buttons.operations_list.click()
        sleep(1)
        buttons.operation.click()
        confirm_win = get_window(app=app, title="Подтверждение")
        confirm_win["&Да"].click()
        wiggle_mouse(duration=2)

        error_win = app.window(title="Произошла ошибка")
        if error_win.exists():
            error_msg = error_win.child_window(class_name="Edit").window_text()
            update_report(
                order=order,
                process=process,
                operation="Создание приказа",
                status=f"Не удалось ИСПОЛНИТЬ приказ. Требуется проверка специалистом. "
                f'Текст ошибки - "{error_msg}"',
            )
            error_win.close()
            orders_win.close()
            personal_win.close()
            continue

        pass

        if order.deputy_fullname is None:
            update_report(
                order=order,
                process=process,
                operation="Создание приказа",
                status="Приказ создан",
            )
            orders_win.close()
            personal_win.close()
            continue

        pass

        update_report(
            order=order,
            process=process,
            operation="Создание приказа",
            status=f"Приказ создан. Доплата за на период командировки сотрудника {order.employee_fullname}",
        )
        orders_win.close()
        personal_win.close()

        pass

    pass
