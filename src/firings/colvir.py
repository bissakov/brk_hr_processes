from time import sleep
from typing import Optional

from src.data import (
    Process,
    FiringOrder,
    Button,
)
from src.utils.colvir_utils import (
    Colvir,
)


def process_order(colvir: Colvir, process: Process, order: FiringOrder) -> str:
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

    colvir.find_and_click_button(
        button=colvir.buttons.operations_list_prs,
        window=orders_win,
        toolbar=orders_win["Static4"],
        target_button_name="Выполнить операцию",
    )

    report_status = confirm_new_entry(colvir=colvir)
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
    order: FiringOrder,
) -> Optional[str]:
    return None


def confirm_new_entry(colvir: Colvir) -> Optional[str]:
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

    command_win = colvir.app.window(title="Распоряжение на командировку")
    if command_win.exists():
        command_win.close()

    error_win = colvir.app.window(title="Произошла ошибка")
    if error_win.exists():
        error_msg = error_win.child_window(class_name="Edit").window_text()
        error_win.close()
        return (
            f"Не удалось ИСПОЛНИТЬ приказ. Требуется проверка специалистом. "
            f'Текст ошибки - "{error_msg}"'
        )
    return None
