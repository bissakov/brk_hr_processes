from time import sleep
from typing import Optional

from src.data import (
    Process,
    MentorshipOrder,
)
from src.utils.colvir_utils import (
    Colvir,
)


def process_order(colvir: Colvir, process: Process, order: MentorshipOrder) -> str:
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

    orders_win.close()
    personal_win.close()
    return "Приказ создан"


def create_new_entry(
    colvir: Colvir,
    order: MentorshipOrder,
) -> Optional[str]:
    return None
