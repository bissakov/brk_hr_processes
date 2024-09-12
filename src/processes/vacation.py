from time import sleep
from typing import Optional

from src.data import VacationOrder, Process
from src.utils.colvir_utils import Colvir


def process_order(colvir: Colvir, process: Process, order: VacationOrder) -> str:
    personal_win, orders_win, report_status = colvir.process_employee_order_status(
        process=process, order=order
    )
    if report_status:
        return report_status

    assert personal_win is not None and orders_win is not None

    report_status = colvir.process_employee_card(order)
    if report_status:
        orders_win.close()
        personal_win.close()
        return report_status

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

    report_status = create_new_entry(colvir=colvir, order=order)
    if report_status:
        orders_win.close()
        personal_win.close()
        return report_status

    orders_win.wait(wait_for="active enabled")

    report_status = colvir.confirm_new_entry(orders_win=orders_win)
    if report_status:
        orders_win.close()
        personal_win.close()
        return report_status

    pass

    if order.deputy_fullname is None:
        orders_win.close()
        personal_win.close()
        return "Приказ создан"

    pass

    orders_win.close()
    personal_win.close()

    return f"Приказ создан. Доплата за на период командировки сотрудника {order.employee_fullname}"


def create_new_entry(
    colvir: Colvir,
    order: VacationOrder,
) -> Optional[str]:
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

    dialog_text = colvir.dialog_text()
    if dialog_text is not None:
        colvir.close_entry_without_saving(order_win=order_win)
        return (
            f"Не удалось заполнить приказ. Требуется проверка специалистом. "
            f"Неизвестное структурное подразделение - {order.branch_num}. "
            f'Текст ошибки - "{dialog_text}"'
        )

    order_win["Edit10"].click_input()
    order_win["Edit10"].type_keys(order.tab_num, pause=0.2)
    order_win.type_keys("{TAB}", pause=1)

    dialog_text = colvir.dialog_text()
    if dialog_text is not None:
        colvir.close_entry_without_saving(order_win=order_win)
        return (
            f"Не удалось заполнить приказ. Требуется проверка специалистом. "
            f"Неизвестный табельный номер - {order.tab_num}. "
            f'Текст ошибки - "{dialog_text}"'
        )

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
    return None
