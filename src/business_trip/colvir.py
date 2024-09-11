from time import sleep
from typing import Optional

import pywinauto

from src.data import BusinessTripOrder, Process
from src.utils.colvir_utils import Colvir


def process_order(colvir: Colvir, process: Process, order: BusinessTripOrder) -> str:
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

    report_status = confirm_new_entry(colvir=colvir, orders_win=orders_win)
    if report_status:
        orders_win.close()
        personal_win.close()
        return report_status

    # command_win = colvir.app.window(title="Распоряжение на командировку")
    # command_win.close()

    komandirovka_win = colvir.app.window(title="Командировка")
    if komandirovka_win.exists():
        if not komandirovka_win.has_focus():
            komandirovka_win.set_focus()

        komandirovka_win["Edit20"].click_input()
        komandirovka_win["Edit20"].set_text(order.order_number)

        komandirovka_win["Edit22"].click_input()
        komandirovka_win["Edit22"].set_text(order.main_order_start_date.short)

        komandirovka_win["Edit18"].click_input()
        komandirovka_win["Edit18"].set_text(order.start_date.short)

        komandirovka_win["Edit16"].click_input()
        komandirovka_win["Edit16"].set_text(order.end_date.short)

        komandirovka_win["Edit14"].click_input()
        komandirovka_win["Edit14"].type_keys("К", pause=0.1)
        komandirovka_win.type_keys("{TAB}", pause=1)

        komandirovka_win["Edit10"].click_input()
        komandirovka_win["Edit10"].type_keys(order.trip_code, pause=0.1)
        komandirovka_win.type_keys("{TAB}", pause=1)

        komandirovka_win["Edit8"].click_input()
        komandirovka_win["Edit8"].type_keys(
            order.trip_reason, pause=0.1, with_spaces=True
        )
        komandirovka_win.type_keys("{TAB}", pause=1)

        if order.deputy_fullname is not None:
            pass

        komandirovka_win["Принять"].click()

        sleep(1)

        error_win = colvir.app.window(title="Произошла ошибка")
        if error_win.exists():
            error_msg = error_win.child_window(class_name="Edit").window_text()
            error_win.close()
            return (
                f"Не удалось ИСПОЛНИТЬ приказ. Требуется проверка специалистом. "
                f'Текст ошибки - "{error_msg}"'
            )

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
    order: BusinessTripOrder,
) -> Optional[str]:
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
    order_win["Edit4"].type_keys(order.branch_num, pause=0.1)
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
    order_win["Edit10"].type_keys(order.tab_num, pause=0.1)
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

    order_win["Edit22"].click_input()
    order_win["Edit22"].set_text(order.start_date.short)

    order_win["Edit24"].click_input()
    order_win["Edit24"].set_text(order.end_date.short)

    if not order.trip_code:
        colvir.close_entry_without_saving(order_win=order_win)
        return (
            f"Не удалось заполнить приказ. Требуется проверка специалистом. "
            f"Неизвестный город/местоположение - {order.trip_place}"
        )

    order_win["Edit28"].type_keys(order.trip_code, pause=0.1)
    order_win["Edit28"].click_input()
    order_win.type_keys("{TAB}", pause=1)

    dialog_text = colvir.dialog_text()
    if dialog_text is not None:
        colvir.close_entry_without_saving(order_win=order_win)
        return (
            f"Не удалось заполнить приказ. Требуется проверка специалистом. "
            f"Неизвестное место назначения - {order.trip_code}. "
            f'Текст ошибки - "{dialog_text}"'
        )

    order_win["Edit16"].type_keys(order.trip_reason, pause=0.1, with_spaces=True)
    order_win["Edit16"].click_input()
    order_win.type_keys("{TAB}", pause=1)

    colvir.find_and_click_button(
        button=colvir.buttons.order_save,
        window=order_win,
        toolbar=order_win["Static3"],
        target_button_name="Сохранить изменения (PgDn)",
    )
    return None


def confirm_new_entry(
    colvir: Colvir, orders_win: pywinauto.WindowSpecification
) -> Optional[str]:
    colvir.find_and_click_button(
        button=colvir.buttons.operations_list_prs,
        window=orders_win,
        toolbar=orders_win["Static4"],
        target_button_name="Выполнить операцию",
    )

    sleep(1)

    popup_menu = colvir.app.PopupMenu

    if not popup_menu.exists():
        raise Exception('Menu "Выполнить операцию" was not clicked')

    colvir.find_and_click_button_temp(
        window=orders_win,
        toolbar=popup_menu,
        target_button_name="Регистрация",
        horizontal=False,
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

    colvir.find_and_click_button_temp(
        window=orders_win,
        toolbar=popup_menu,
        target_button_name="Утвердить",
        horizontal=False,
    )

    confirm_win = colvir.utils.get_window(title="Подтверждение")
    confirm_win["&Да"].click()
    colvir.utils.wiggle_mouse(duration=2)

    colvir.buttons.operations_list_prs.click()
    sleep(1)

    colvir.find_and_click_button_temp(
        window=orders_win,
        toolbar=popup_menu,
        target_button_name="Исполнить",
        horizontal=False,
    )

    confirm_win = colvir.utils.get_window(title="Подтверждение")
    confirm_win["&Да"].click()
    colvir.utils.wiggle_mouse(duration=2)

    error_win = colvir.app.window(title="Произошла ошибка")
    if error_win.exists():
        error_msg = error_win.child_window(class_name="Edit").window_text()
        error_win.close()
        return (
            f"Не удалось ИСПОЛНИТЬ приказ. Требуется проверка специалистом. "
            f'Текст ошибки - "{error_msg}"'
        )

    return None
