import dataclasses
from datetime import datetime
from enum import Enum
from typing import Optional, Union, Iterator, NamedTuple, Tuple


class ProcessType(Enum):
    BUSINESS_TRIP = 13635
    VACATION = 13630
    VACATION_WITHDRAW = 13631
    FIRING = 13632
    MENTORSHIP = 13636
    VACATION_ADD_PAY = 854


class Process(NamedTuple):
    process_type: ProcessType
    process_name: str
    order_type: str
    download_url: str
    csv_path: str
    report_folder: str
    pickle_path: str
    report_path: str
    today: str


@dataclasses.dataclass(slots=True, frozen=True)
class Processes:
    business_trip: Process
    vacation: Process
    vacation_withdraw: Process
    firing: Process
    mentorship: Process
    vacation_add_pay: Process

    def __iter__(self) -> Iterator[Process]:
        for field in dataclasses.fields(self):
            yield getattr(self, field.name)


@dataclasses.dataclass(slots=True)
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


@dataclasses.dataclass(slots=True)
class BusinessTripOrder:
    employee_fullname: str
    employee_names: Tuple[str, str]
    order_number: str
    sign_date: Date
    start_date: Date
    end_date: Date
    trip_place: str
    trip_code: str
    trip_reason: str
    main_order_number: str
    main_order_start_date: Date
    deputy_fullname: Optional[str]
    deputy_names: Optional[Tuple[str, str]]
    employee_status: Optional[str] = None
    branch_num: Optional[str] = None
    tab_num: Optional[str] = None

    def as_dict_short(self):
        return {
            "employee_fullname": self.employee_fullname,
            "order_number": self.order_number,
            "sign_date": self.sign_date.short,
            "start_date": self.start_date.short,
            "end_date": self.end_date.short,
            "trip_place": self.trip_place,
            "trip_code": self.trip_code,
            "trip_reason": self.trip_reason,
            "main_order_number": self.main_order_number,
            "main_order_start_date": self.main_order_start_date.short,
            "deputy_fullname": self.deputy_fullname,
        }

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
            "trip_reason": self.trip_reason,
            "main_order_number": self.main_order_number,
            "main_order_start_date": self.main_order_start_date.as_dict(),
            "deputy_fullname": self.deputy_fullname,
            "deputy_names": self.deputy_names,
        }


@dataclasses.dataclass(slots=True)
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

    def as_dict_short(self):
        return {
            "employee_fullname": self.employee_fullname,
            "order_type": self.order_type,
            "order_number": self.order_number,
            "start_date": self.start_date.short,
            "end_date": self.end_date.short,
            "deputy_fullname": self.deputy_fullname,
            "surcharge": self.surcharge,
            "substitution_start": self.substitution_start,
            "substitution_end": self.substitution_end,
        }

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


@dataclasses.dataclass(slots=True)
class VacationWithdrawOrder:
    employee_fullname: str
    employee_names: Tuple[str, str]
    order_type: str
    order_number: str
    withdraw_date: Date
    employee_status: Optional[str] = None
    branch_num: Optional[str] = None
    tab_num: Optional[str] = None

    def as_dict_short(self):
        return {
            "employee_fullname": self.employee_fullname,
            "order_type": self.order_type,
            "order_number": self.order_number,
            "withdraw_date": self.withdraw_date.short,
        }

    def as_dict(self):
        return {
            "employee_fullname": self.employee_fullname,
            "employee_names": self.employee_names,
            "order_type": self.order_type,
            "order_number": self.order_number,
            "withdraw_date": self.withdraw_date.as_dict(),
        }


@dataclasses.dataclass(slots=True)
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

    def as_dict_short(self):
        return {
            "employee_fullname": self.employee_fullname,
            "firing_reason": self.firing_reason,
            "order_number": self.order_number,
            "firing_date": self.firing_date.short,
        }

    def as_dict(self):
        return {
            "employee_fullname": self.employee_fullname,
            "employee_names": self.employee_names,
            "firing_reason": self.firing_reason,
            "order_number": self.order_number,
            "compensation": self.compensation,
            "firing_date": self.firing_date.as_dict(),
        }


@dataclasses.dataclass(slots=True)
class MentorshipOrder:
    employee_fullname: str
    work_start_date: Date
    contract_start_date: Date
    contract_end_date: Date
    mentor_fullname: str
    mentrorship_order_number: str
    mentorship_start_date: Date
    mentorship_end_date: Date
    creation_date: Date
    employee_status: Optional[str] = None
    branch_num: Optional[str] = None
    tab_num: Optional[str] = None

    def as_dict_short(self):
        return {
            "employee_fullname": self.employee_fullname,
            "work_start_date": self.work_start_date.short,
            "contract_start_date": self.contract_start_date.short,
            "contract_end_date": self.contract_end_date.short,
            "mentor_fullname": self.mentor_fullname,
            "mentrorship_order_number": self.mentrorship_order_number,
            "mentorship_start_date": self.mentorship_start_date.short,
            "mentorship_end_date": self.mentorship_end_date.short,
            "creation_date": self.creation_date.short,
        }

    def as_dict(self):
        return {
            "employee_fullname": self.employee_fullname,
            "work_start_date": self.work_start_date.as_dict(),
            "contract_start_date": self.contract_start_date.as_dict(),
            "contract_end_date": self.contract_end_date.as_dict(),
            "mentor_fullname": self.mentor_fullname,
            "mentrorship_order_number": self.mentrorship_order_number,
            "mentorship_start_date": self.mentorship_start_date.as_dict(),
            "mentorship_end_date": self.mentorship_end_date.as_dict(),
            "creation_date": self.creation_date.as_dict(),
        }


@dataclasses.dataclass(slots=True)
class VacationAddPayOrder:
    # FIXME
    date: Date
    employee_status: Optional[str] = None
    branch_num: Optional[str] = None
    tab_num: Optional[str] = None

    def as_dict_short(self):
        raise NotImplementedError()

    def as_dict(self):
        raise NotImplementedError()


Order = Union[
    BusinessTripOrder,
    VacationOrder,
    VacationWithdrawOrder,
    FiringOrder,
    MentorshipOrder,
    VacationAddPayOrder,
]
