import pickle
from time import sleep
from typing import List

from robots.data import VacationOrder, Button, Process
from robots.notification import TelegramAPI
from robots.utils.colvir_utils import (
    Colvir,
)
from robots.utils.utils import (
    create_report,
    update_report,
)


def run(colvir: Colvir, process: Process, bot: TelegramAPI) -> None:
    with open(process.pickle_path, "rb") as f:
        orders: List[VacationOrder] = pickle.load(f)
    assert all(isinstance(order, VacationOrder) for order in orders)

    create_report(process.report_path)

    for i, order in enumerate(orders):
        personal_win, orders_win, report_status = colvir.process_employee_order_status(
            process=process, order=order
        )
        if report_status:
            update_report(
                order=order,
                process=process,
                operation="Создание приказа",
                status=report_status,
            )
            continue
        assert personal_win is not None and orders_win is not None

        report_status = colvir.process_employee_card(order)
        if report_status:
            orders_win.close()
            personal_win.close()
            update_report(
                order=order,
                process=process,
                operation="Создание приказа",
                status=report_status,
            )
            continue
        assert (
            order.branch_num is not None
            and order.tab_num is not None
            and order.employee_status is not None
        )

        personal_win.set_focus()
        if order.employee_status == "В командировке":
            colvir.return_from("Возврат из командировки", personal_win)
        elif order.employee_status == "В отпуске":
            colvir.return_from("Возврат из отпуска", personal_win)

        if order.employee_status != "Работающий":
            pass

        orders_win.set_focus()
        sleep(1)
        orders_win.wait(wait_for="active enabled")

        colvir.find_and_click_button(
            button=colvir.buttons.create_new_order,
            window=orders_win,
            toolbar=orders_win["Static4"],
            target_button_name="Создать новую запись (Ins)",
        )

        order_win = colvir.utils.get_window(title="Приказ")

        order_win["Edit18"].type_keys("ORD_HOL", pause=0.1)
        order_win["Edit18"].type_keys("{TAB}")
        sleep(0.5)
        if (error_win := colvir.app.window(title="Произошла ошибка")).exists():
            error_win.close()
            order_win["Edit38"].type_keys("{TAB}")

        sleep(1)
        order_win["Edit48"].type_keys(order.order_number, pause=0.1)
        sleep(1)

        order_win["Edit4"].click_input()
        order_win["Edit4"].type_keys(order.branch_num, pause=0.2)
        order_win.type_keys("{TAB}", pause=1)
        order_win["Edit10"].click_input()
        order_win["Edit10"].type_keys(order.tab_num, pause=0.2)
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

        colvir.find_and_click_button(
            button=colvir.buttons.order_save,
            window=order_win,
            toolbar=order_win["Static3"],
            target_button_name="Сохранить изменения (PgDn)",
        )

        orders_win.wait(wait_for="active enabled")

        colvir.find_and_click_button(
            button=colvir.buttons.operations_list_prs,
            window=orders_win,
            toolbar=orders_win["Static4"],
            target_button_name="Выполнить операцию",
        )

        sleep(0.5)
        colvir.buttons.operation = Button(
            colvir.buttons.operations_list_prs.x,
            colvir.buttons.operations_list_prs.y + 25,
        )
        colvir.check_and_click(
            button=colvir.buttons.operation, target_button_name="Регистрация"
        )

        registration_win = colvir.utils.get_window(title="Подтверждение")
        registration_win["&Да"].click()
        sleep(2)
        confirm_win = colvir.app.window(title="Подтверждение")
        if confirm_win.exists():
            confirm_win.close()
        sleep(1)
        dossier_win = colvir.app.window(title="Досье сотрудника")
        if dossier_win.exists():
            dossier_win.close()

        colvir.utils.wiggle_mouse(duration=2)

        colvir.buttons.operations_list_prs.click()
        sleep(1)
        colvir.buttons.operation.click()
        confirm_win = colvir.utils.get_window(title="Подтверждение")
        confirm_win["&Да"].click()
        colvir.utils.wiggle_mouse(duration=2)

        colvir.buttons.operations_list_prs.click()
        sleep(1)
        colvir.buttons.operation.click()
        confirm_win = colvir.utils.get_window(title="Подтверждение")
        confirm_win["&Да"].click()
        colvir.utils.wiggle_mouse(duration=2)

        error_win = colvir.app.window(title="Произошла ошибка")
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
