from datetime import datetime
from enum import Enum
from time import sleep
from typing import Optional, Tuple, Union

import pywinauto
from attr import define
from pywinauto import mouse


class ProcessType(Enum):
    BUSINESS_TRIP = 13635
    VACATION = 13630
    VACATION_WITHDRAW = 13631
    FIRING = 13632


@define
class Process:
    process_type: ProcessType
    download_url: str
    csv_path: str
    report_folder: str
    pickle_path: str
    report_path: str
    today: str


@define
class Processes:
    business_trip: Process
    vacation: Process
    vacation_withdraw: Process
    firing: Process

    def __iter__(self):
        return iter(
            (self.business_trip, self.vacation, self.vacation_withdraw, self.firing)
        )


@define
class ChromePath:
    driver_path: str
    binary_path: str


@define
class CredentialsBPM:
    user: str
    password: str


@define
class BpmInfo:
    chrome_path: ChromePath
    creds: CredentialsBPM
    download_folder: str


@define
class ColvirInfo:
    location: str
    user: str
    password: str


@define
class Date:
    dt: datetime
    long: Optional[str] = None
    short: Optional[str] = None
    colvir: Optional[str] = None

    def __attrs_post_init__(self):
        self.long = self.dt.strftime("%d.%m.%Y")
        self.short = self.dt.strftime("%d.%m.%y")
        self.colvir = self.dt.strftime("%d%m%y")

    def as_dict(self):
        return {
            "dt": self.dt.isoformat(),
            "long": self.long,
            "short": self.short,
            "colvir": self.colvir,
        }


@define
class BusinessTripOrder:
    employee_fullname: str
    employee_names: Tuple[str, str]
    order_number: str
    sign_date: Date
    start_date: Date
    end_date: Date
    trip_place: str
    trip_code: str
    trip_target: str
    main_order_number: str
    main_order_start_date: Date
    deputy_fullname: Optional[str]
    deputy_names: Optional[Tuple[str, str]]
    employee_status: Optional[str] = None
    branch_num: Optional[str] = None
    tab_num: Optional[str] = None

    def as_dict(self):
        return {
            "employee_fullname": self.employee_fullname,
            "employee_names": self.employee_names,
            "order_number": self.order_number,
            "sign_date": self.sign_date.as_dict(),
            "start_date": self.start_date.as_dict(),
            "end_date": self.end_date.as_dict(),
            "trip_place": self.trip_place,
            "trip_code": self.trip_code,
            "trip_target": self.trip_target,
            "main_order_number": self.main_order_number,
            "main_order_start_date": self.main_order_start_date.as_dict(),
            "deputy_fullname": self.deputy_fullname,
            "deputy_names": self.deputy_names,
        }


@define
class VacationOrder:
    employee_fullname: str
    employee_names: Tuple[str, str]
    order_type: str
    start_date: Date
    end_date: Date
    order_number: str
    deputy_fullname: str
    deputy_names: Optional[Tuple[str, str]]
    surcharge: str
    substitution_start: str
    substitution_end: str
    employee_status: Optional[str] = None
    branch_num: Optional[str] = None
    tab_num: Optional[str] = None

    def as_dict(self):
        return {
            "employee_fullname": self.employee_fullname,
            "employee_names": self.employee_names,
            "order_type": self.order_type,
            "start_date": self.start_date.as_dict(),
            "end_date": self.end_date.as_dict(),
            "order_number": self.order_number,
            "deputy_fullname": self.deputy_fullname,
            "deputy_names": self.deputy_names,
            "surcharge": self.surcharge,
            "substitution_start": self.substitution_start,
            "substitution_end": self.substitution_end,
        }


@define
class VacationWithdrawOrder:
    employee_fullname: str
    employee_names: Tuple[str, str]
    order_type: str
    order_number: str
    withdraw_date: Date
    employee_status: Optional[str] = None
    branch_num: Optional[str] = None
    tab_num: Optional[str] = None

    def as_dict(self):
        return {
            "employee_fullname": self.employee_fullname,
            "employee_names": self.employee_names,
            "order_type": self.order_type,
            "order_number": self.order_number,
            "withdraw_date": self.withdraw_date.as_dict(),
        }


@define
class FiringOrder:
    employee_fullname: str
    employee_names: Tuple[str, str]
    firing_reason: str
    order_number: str
    compensation: str
    firing_date: Date
    employee_status: Optional[str] = None
    branch_num: Optional[str] = None
    tab_num: Optional[str] = None

    def as_dict(self):
        return {
            "employee_fullname": self.employee_fullname,
            "employee_names": self.employee_names,
            "firing_reason": self.firing_reason,
            "order_number": self.order_number,
            "compensation": self.compensation,
            "firing_date": self.firing_date.as_dict(),
        }


Order = Union[BusinessTripOrder, VacationOrder, VacationWithdrawOrder, FiringOrder]


@define
class Button:
    x: int = -1
    y: int = -1

    def click(self) -> None:
        mouse.click(button="left", coords=(self.x, self.y))


@define
class Buttons:
    clear_form: Button = Button()
    employee_orders: Button = Button()
    create_new_order: Button = Button()
    order_save: Button = Button()
    operations_list_prs: Button = Button()
    operations_list_orders: Button = Button()
    operation: Button = Button()
    cities_menu: Button = Button()
