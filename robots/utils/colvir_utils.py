import os
import random
import re
from datetime import timedelta
from time import sleep
from typing import Optional, Tuple, cast

import pandas as pd
import psutil
import pyautogui
import pywinauto
import pywinauto.timings
import win32con
import win32gui
from pywinauto import mouse, win32functions

from robots.data import (
    ColvirInfo,
    Date,
    Order,
    Buttons,
    Button,
    Process,
    BusinessTripOrder,
    VacationOrder,
    VacationWithdrawOrder,
    FiringOrder,
)
from robots.utils.excel_utils import xls_to_xlsx


pyautogui.FAILSAFE = False


def kill_all_processes(proc_name: str) -> None:
    for proc in psutil.process_iter():
        try:
            if proc_name in proc.name():
                proc.terminate()
        except (psutil.AccessDenied, psutil.NoSuchProcess):
            continue


class ColvirUtils:
    def __init__(self, app: Optional[pywinauto.Application]):
        self.app = app

    @staticmethod
    def close_window(
        win: pywinauto.WindowSpecification, raise_error: bool = False
    ) -> None:
        if win.exists():
            win.close()
            return

        if raise_error:
            raise pywinauto.findwindows.ElementNotFoundError(
                f"Window {win} does not exist"
            )

    @staticmethod
    def set_focus_win32(win: pywinauto.WindowSpecification) -> None:
        if win.wrapper_object().has_focus():
            return

        handle = win.wrapper_object().handle

        mouse.move(coords=(-10000, 500))
        if win.is_minimized():
            if win.was_maximized():
                win.maximize()
            else:
                win.restore()
        else:
            win32gui.ShowWindow(handle, win32con.SW_SHOW)
        win32gui.SetForegroundWindow(handle)

        win32functions.WaitGuiThreadIdle(handle)

    @staticmethod
    def set_focus(win: pywinauto.WindowSpecification, retries: int = 20) -> None:
        while retries > 0:
            try:
                if retries % 2 == 0:
                    ColvirUtils.set_focus_win32(win)
                else:
                    win.set_focus()
                break
            except (Exception, BaseException):
                retries -= 1
                sleep(5)
                continue

        if retries <= 0:
            raise Exception("Failed to set focus")

    @staticmethod
    def press(win: pywinauto.WindowSpecification, key: str, pause: float = 0) -> None:
        ColvirUtils.set_focus(win)
        win.type_keys(key, pause=pause, set_foreground=False)

    @staticmethod
    def type_keys(
        window: pywinauto.WindowSpecification,
        keystrokes: str,
        step_delay: float = 0.1,
        delay_after: float = 0.5,
    ) -> None:
        ColvirUtils.set_focus(window)
        for command in list(filter(None, re.split(r"({.+?})", keystrokes))):
            try:
                window.type_keys(command, set_foreground=False)
            except pywinauto.base_wrapper.ElementNotEnabled:
                sleep(1)
                window.type_keys(command, set_foreground=False)
            sleep(step_delay)

        sleep(delay_after)

    def get_window(
        self,
        title: str,
        wait_for: str = "exists",
        timeout: int = 20,
        regex: bool = False,
        found_index: int = 0,
    ) -> pywinauto.WindowSpecification:
        if regex:
            window = self.app.window(title_re=title, found_index=found_index)
        else:
            window = self.app.window(title=title, found_index=found_index)
        window.wait(wait_for=wait_for, timeout=timeout)
        sleep(0.5)
        return window

    def persistent_win_exists(self, title_re: str, timeout: float) -> bool:
        try:
            self.app.window(title_re=title_re).wait(wait_for="enabled", timeout=timeout)
        except pywinauto.timings.TimeoutError:
            return False
        return True

    @staticmethod
    def does_order_exist(
        orders_file_path: str, order_type: str, order_number: str
    ) -> bool:
        df = pd.read_excel(orders_file_path, skiprows=1)

        order_exists = (
            (df["Вид приказа"] == order_type) & (df["Номер приказа"] == order_number)
        ).any()

        return order_exists

    @staticmethod
    def wiggle_mouse(duration: int) -> None:
        def get_random_coords() -> Tuple[int, int]:
            screen = pyautogui.size()
            width = screen[0]
            height = screen[1]

            return random.randint(100, width - 200), random.randint(100, height - 200)

        max_wiggles = random.randint(4, 9)
        step_sleep = duration / max_wiggles

        for _ in range(1, max_wiggles):
            coords = get_random_coords()
            pyautogui.moveTo(x=coords[0], y=coords[1], duration=step_sleep)


class Colvir:
    def __init__(self, colvir_info: ColvirInfo, buttons: Buttons) -> None:
        self.info = colvir_info
        self.app: Optional[pywinauto.Application] = None
        self.utils = ColvirUtils(app=self.app)
        self.buttons = buttons

    def open_colvir(self) -> None:
        for _ in range(10):
            try:
                self.app = pywinauto.Application().start(cmd_line=self.info.location)
                self.login()
                self.check_interactivity()
                break
            except pywinauto.findwindows.ElementNotFoundError:
                kill_all_processes("COLVIR")
                continue
        assert self.app is not None, Exception("max_retries exceeded")
        self.utils.app = self.app

    def change_password(self) -> bool:
        attention_win = self.app.window(title="Внимание")
        if not attention_win.exists():
            return False
        attention_win["OK"].click()

        change_password_win = self.app.window(title_re="Смена пароля.+")
        change_password_win["Edit0"].set_text("ROBOTIZ2024_")
        change_password_win["Edit2"].set_text("ROBOTIZ2024_")
        change_password_win["OK"].click()

        confirm_win = self.app.window(title="Colvir Banking System", found_index=1)
        confirm_win.send_keystrokes("{ENTER}")

        mode_win = self.app.window(title="Выбор режима")
        return mode_win.exists()

    def login(self) -> None:
        login_win = self.app.window(title="Вход в систему")

        login_username = login_win["Edit2"]
        login_password = login_win["Edit"]

        login_username.set_text(text=self.info.user)
        if login_username.window_text() != self.info.user:
            login_username.set_text("")
            login_username.type_keys(self.info.user, set_foreground=False)

        login_password.set_text(text=self.info.password)
        if login_password.window_text() != self.info.password:
            login_password.set_text("")
            login_password.type_keys(self.info.password, set_foreground=False)

        login_win["OK"].click()

        sleep(1)
        if login_win.exists() and self.app.window(title="Произошла ошибка").exists():
            raise pywinauto.findwindows.ElementNotFoundError()

    def check_interactivity(self) -> None:
        self.choose_mode(mode="TREPRT")
        sleep(1)

        reports_win = self.app.window(title="Выбор отчета")
        self.utils.close_window(win=reports_win, raise_error=True)

    def choose_mode(self, mode: str) -> None:
        mode_win = self.app.window(title="Выбор режима")
        mode_win["Edit2"].set_text(text=mode)
        self.utils.press(mode_win["Edit2"], "~")

    def close_dialog(self) -> None:
        dialog_win = self.utils.get_window(title="Colvir Banking System", found_index=0)
        dialog_win.set_focus()
        sleep(0.5)
        dialog_win["OK"].click_input()

    def check_and_click(self, button: Button, target_button_name: str) -> None:
        mouse.move(coords=(button.x, button.y))
        status_bar = self.app.window(title_re="Банковская система.+")["StatusBar"]
        if status_bar.window_text().strip() == target_button_name:
            button.click()

    def find_and_click_button(
        self,
        button: Button,
        window: pywinauto.WindowSpecification,
        toolbar: pywinauto.WindowSpecification,
        target_button_name: str,
        horizontal: bool = True,
        offset: int = 5,
    ) -> Button:
        if not window.has_focus():
            window.set_focus()

        if button.x != -1 and button.y != -1:
            button.click()
            return button

        status_win = self.app.window(title_re="Банковская система.+")
        rectangle = toolbar.rectangle()
        mid_point = rectangle.mid_point()
        mouse.move(coords=(mid_point.x, mid_point.y))

        start_point = rectangle.left if horizontal else rectangle.top
        end_point = rectangle.right if horizontal else rectangle.bottom

        x, y = mid_point.x, mid_point.y
        point = 0

        x_offset = offset if horizontal else 0
        y_offset = offset if not horizontal else 0

        i = 0
        while (
            status_win["StatusBar"].window_text().strip() != target_button_name
            or point >= end_point
        ):
            point = start_point + i * 5

            if horizontal:
                x = point
            else:
                y = point

            mouse.move(coords=(x, y))
            i += 1

        window.set_focus()
        sleep(1)

        button.x = x + x_offset
        button.y = y + y_offset
        button.click()

        return button

    def save_excel(self, work_folder: str) -> str:
        file_win = self.utils.get_window(title="Выберите файл для экспорта")

        orders_file_path = os.path.join(work_folder, "orders.xls")
        orders_xlsx_file_path = os.path.join(work_folder, "orders.xlsx")

        file_win["Edit4"].set_text(orders_file_path)
        file_win["&Save"].click_input()

        sleep(1)
        confirm_win = self.app.window(title="Confirm Save As")
        if confirm_win.exists():
            confirm_win["Yes"].click()

        sort_win = self.utils.get_window(title="Сортировка")
        sort_win["OK"].click()

        while not os.path.exists(orders_file_path):
            sleep(5)
        sleep(1)

        kill_all_processes("EXCEL")

        xls_to_xlsx(orders_file_path, orders_xlsx_file_path)

        return orders_xlsx_file_path

    def change_oper_day(self, start_date: Date):
        self.choose_mode(mode="TOPERD")
        current_oper_day_win = self.utils.get_window(title="Текущий операционный день")
        current_oper_day_win["Edit2"].set_text(start_date.short)
        current_oper_day_win["OK"].click()
        attention_win = self.app.window(title="Внимание")
        if not attention_win.exists():
            start_date = Date(start_date.dt - timedelta(days=1))
            self.close_dialog()
            self.change_oper_day(start_date=start_date)
            return

        attention_win["&Да"].click()
        current_oper_day_win["OK"].click()
        sleep(0.5)
        self.close_dialog()

    def find_employee(
        self,
        employee_names: Tuple[str, str],
    ) -> bool:
        self.choose_mode(mode="PRS")
        filter_win = self.utils.get_window(title="Фильтр")
        self.find_and_click_button(
            button=self.buttons.clear_form,
            window=filter_win,
            toolbar=filter_win["Static3"],
            target_button_name="Очистить фильтр",
        )

        filter_win["Edit8"].set_text("001")
        filter_win["Edit4"].set_text(employee_names[0])
        filter_win["Edit2"].set_text(employee_names[1])
        filter_win["OK"].click()
        sleep(1)

        confirm_win = self.app.window(title="Подтверждение")
        if confirm_win.exists():
            confirm_win.close()
            filter_win.close()
            personal_win = self.app.window(title="Персонал")
            if personal_win.exists():
                personal_win.close()
            return False

        return True

    def process_employee_order_status(self, process: Process, order: Order) -> Tuple[
        Optional[pywinauto.WindowSpecification],
        Optional[pywinauto.WindowSpecification],
        Optional[str],
    ]:
        if isinstance(order, BusinessTripOrder):
            order = cast(BusinessTripOrder, order)
            start_date = order.start_date
        elif isinstance(order, VacationOrder):
            order = cast(VacationOrder, order)
            start_date = order.start_date
        elif isinstance(order, VacationWithdrawOrder):
            order = cast(VacationWithdrawOrder, order)
            start_date = order.withdraw_date
        elif isinstance(order, FiringOrder):
            order = cast(FiringOrder, order)
            start_date = order.firing_date
        else:
            raise ValueError(f"Order is of unknown type - {type(order)}")

        self.change_oper_day(start_date=start_date)
        if not self.find_employee(employee_names=order.employee_names):
            return None, None, "Приказ не найден"

        personal_win = self.utils.get_window(title="Персонал")
        self.find_and_click_button(
            button=self.buttons.employee_orders,
            window=personal_win,
            toolbar=personal_win["Static4"],
            target_button_name="Приказы по сотруднику",
        )

        orders_win = self.utils.get_window(title="Приказы сотрудника")
        orders_win.menu_select("#4->#4->#1")

        if self.utils.does_order_exist(
            orders_file_path=self.save_excel(work_folder=process.report_folder),
            order_type="Приказ о отправке работника в командировку",
            order_number=order.order_number,
        ):
            orders_win.close()
            personal_win.close()
            return None, None, "Приказ уже создан"

        personal_win.set_focus()
        sleep(1)
        personal_win.type_keys("{ENTER}")

        return personal_win, orders_win, None

    def process_employee_card(self, order: Order) -> Optional[str]:
        employee_card = self.utils.get_window(title="Карточка сотрудника")
        order.employee_status = employee_card["Edit30"].window_text().strip()
        print(order.employee_fullname, order.employee_status)

        if order.employee_status == "Уволен":
            employee_card.close()
            return (
                f"Невозможно создать приказ для сотрудника "
                f'со статусом "{order.employee_status}"'
            )

        order.branch_num = employee_card["Edit60"].window_text()
        order.tab_num = employee_card["Edit34"].window_text()

        employee_card.close()

        return None

    def return_from(
        self, target_button_name: str, personal_win: pywinauto.WindowSpecification
    ) -> None:
        self.find_and_click_button(
            button=self.buttons.operations_list_prs,
            window=personal_win,
            toolbar=personal_win["Static4"],
            target_button_name="Выполнить операцию",
        )

        sleep(0.5)
        self.buttons.operation = Button(
            self.buttons.operations_list_prs.x,
            self.buttons.operations_list_prs.y + 25,
        )
        self.check_and_click(
            button=self.buttons.operation,
            target_button_name=target_button_name,
        )

        confirm_win = self.utils.get_window(title="Подтверждение")
        confirm_win["&Да"].click()
        self.utils.wiggle_mouse(duration=2)

        return_win = self.utils.get_window(title="Возврат из командировки")
        return_win["Принять"].click()

    def __enter__(self) -> "Colvir":
        self.open_colvir()
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        if not self.app.kill():
            kill_all_processes("COLVIR")
