import os
import pickle
import sys
import warnings
from datetime import datetime
from typing import List, Type, Callable, Tuple
from urllib.parse import urljoin

import dotenv

from processes import business_trip
from processes import firing
from processes import mentorship
from processes import vacation
from processes import vacation_add_pay
from processes import vacation_withdraw
from src import bpm
from src import mail
from src.data import (
    Processes,
    Process,
    ProcessType,
    BusinessTripOrder,
    VacationOrder,
    VacationWithdrawOrder,
    FiringOrder,
    Order,
    MentorshipOrder,
    VacationAddPayOrder,
)
from src.notification import TelegramAPI, handle_error
from src.utils.colvir_utils import Colvir, ColvirInfo
from src.utils.utils import create_report, update_report

if sys.version_info.major != 3 or sys.version_info.minor != 12:
    raise RuntimeError(f"Python {sys.version_info} is not supported")

warnings.simplefilter(action="ignore", category=UserWarning)
project_folder = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
dotenv.load_dotenv(os.path.join(project_folder, ".env.test"))


def get_from_env(key: str) -> str:
    value = os.getenv(key)
    assert isinstance(value, str), f"{key} not set in the environment variables"
    return value


def get_processes(
    bpm_base_url: str,
    download_folder: str,
    report_root_folder: str,
    today: str,
) -> Processes:
    def get_process(
        process_type: ProcessType,
    ) -> Process:
        process_type_name = process_type.name.lower()
        csv_filename = f"{process_type_name}_{today}.csv"
        pickle_filename = f"{process_type_name}_{today}.pkl"

        download_url = urljoin(
            bpm_base_url, f"?s=rep_b&id={process_type.value}&reset_page=1&gid=739"
        )

        match process_type:
            case ProcessType.BUSINESS_TRIP:
                report_filename = f"Отчет_командировки_{today}.xlsx"
                process_name = "Командировки"
                order_type = "Приказ о отправке работника в командировку"
            case ProcessType.VACATION:
                report_filename = f"Отчет_отпусков_{today}.xlsx"
                process_name = "Отпуска"
                order_type = "Приказ о отправке работника в командировку"
            case ProcessType.VACATION_WITHDRAW:
                report_filename = f"Отчет_отзывов_отпусков_{today}.xlsx"
                process_name = "Отзывы из отпусков"
                order_type = "Приказ о отправке работника в командировку"
            case ProcessType.FIRING:
                report_filename = f"Отчет_увольнений_{today}.xlsx"
                process_name = "Увольнения"
                order_type = "Приказ о отправке работника в командировку"
            case ProcessType.MENTORSHIP:
                report_filename = f"Отчет_менторств_{today}.xlsx"
                process_name = "Менторства"
                order_type = "Приказ о отправке работника в командировку"
            case ProcessType.VACATION_ADD_PAY:
                report_filename = (
                    f"Отчет_доплата_совмещение_на_период_отпуска_{today}.xlsx"
                )
                process_name = "Доплата за совмещение должностей на период отпуска"
                order_type = "Приказ о отправке работника в командировку"
                download_url = urljoin(
                    bpm_base_url, f"?s=obj_a&gid={process_type.value}&reset_page=1"
                )
            case _:
                raise ValueError(
                    f"Unknown process type: ProcessType(name={process_type.name}, value={process_type.value})"
                )

        report_folder = os.path.join(report_root_folder, process_type_name)
        os.makedirs(report_folder, exist_ok=True)

        csv_folder = os.path.join(download_folder, process_type.name.lower())
        os.makedirs(csv_folder, exist_ok=True)

        csv_path = os.path.join(csv_folder, csv_filename)
        pickle_path = os.path.join(report_folder, pickle_filename)
        report_path = os.path.join(report_folder, report_filename)

        return Process(
            process_type=process_type,
            process_name=process_name,
            order_type=order_type,
            download_url=download_url,
            csv_path=csv_path,
            report_folder=report_folder,
            pickle_path=pickle_path,
            report_path=report_path,
            today=today,
        )

    processes = Processes(
        business_trip=get_process(ProcessType.BUSINESS_TRIP),
        vacation=get_process(ProcessType.VACATION),
        vacation_withdraw=get_process(ProcessType.VACATION_WITHDRAW),
        firing=get_process(ProcessType.FIRING),
        mentorship=get_process(ProcessType.MENTORSHIP),
        vacation_add_pay=get_process(ProcessType.VACATION_ADD_PAY),
    )
    return processes


@handle_error
def run(bot: TelegramAPI) -> None:
    data_folder = os.path.join(project_folder, "data")
    os.makedirs(data_folder, exist_ok=True)

    today_dt = datetime.now()
    current_month_name = today_dt.strftime("%B")

    report_root_folder = os.path.join(
        data_folder, "reports", str(today_dt.year), current_month_name
    )
    os.makedirs(report_root_folder, exist_ok=True)

    download_folder = os.path.join(
        data_folder,
        "downloads",
        str(today_dt.year),
        current_month_name,
    )
    os.makedirs(download_folder, exist_ok=True)

    bpm_info = bpm.BpmInfo(
        creds=bpm.CredentialsBPM(
            user=get_from_env("BPM_USER"), password=get_from_env("BPM_PASSWORD")
        ),
        chrome_path=bpm.ChromePath(
            driver_path=os.path.join(project_folder, get_from_env("DRIVER_PATH")),
            binary_path=os.path.join(project_folder, get_from_env("CHROME_PATH")),
        ),
        download_folder=download_folder,
    )

    colvir_info = ColvirInfo(
        location=get_from_env("COLVIR_PATH"),
        user=get_from_env("COLVIR_USER"),
        password=get_from_env("COLVIR_PASSWORD"),
    )

    today = today_dt.strftime("%d.%m.%y")
    bot.send_message(
        f"Старт процесса за {today}\n"
        f'"Командировки, отпуска, отзывы из отпуска и увольнения"'
    )

    bpm_base_url = get_from_env("BPM_BASE_URL")
    processes = get_processes(
        bpm_base_url=bpm_base_url,
        download_folder=download_folder,
        report_root_folder=report_root_folder,
        today=today,
    )

    is_logged_in = False
    with bpm.driver_init(bpm_info=bpm_info) as driver:
        for process in processes:
            bpm.run(
                driver=driver,
                creds=bpm_info.creds,
                process=process,
                bot=bot,
                is_logged_in=is_logged_in,
            )
            is_logged_in = True

    with Colvir(colvir_info=colvir_info) as colvir:
        process_run(process=processes.business_trip, colvir=colvir, bot=bot)
        process_run(process=processes.vacation, colvir=colvir, bot=bot)
        process_run(process=processes.vacation_withdraw, colvir=colvir, bot=bot)
        process_run(process=processes.firing, colvir=colvir, bot=bot)
        process_run(process=processes.mentorship, colvir=colvir, bot=bot)
        process_run(process=processes.vacation_add_pay, colvir=colvir, bot=bot)

    bot.send_message("Успешное окончание процесса")


ProcessCallable = Callable[[Colvir, Process, Order], str]


def get_order_type_and_processor(
    process_type: ProcessType,
) -> Tuple[Type[Order], ProcessCallable]:
    match process_type:
        case ProcessType.BUSINESS_TRIP:
            return BusinessTripOrder, business_trip.process_order
        case ProcessType.VACATION:
            return VacationOrder, vacation.process_order
        case ProcessType.VACATION_WITHDRAW:
            return VacationWithdrawOrder, vacation_withdraw.process_order
        case ProcessType.FIRING:
            return FiringOrder, firing.process_order
        case ProcessType.MENTORSHIP:
            return MentorshipOrder, mentorship.process_order
        case ProcessType.VACATION_ADD_PAY:
            return VacationAddPayOrder, vacation_add_pay.process_order
        case _:
            ValueError(
                f"Unknown process type: ProcessType(name={process_type.name}, value={process_type.value})"
            )


def process_run(process: Process, colvir: Colvir, bot: TelegramAPI):
    order_t, process_order = get_order_type_and_processor(process.process_type)

    with open(process.pickle_path, "rb") as f:
        orders: List[order_t] = pickle.load(f)
    assert all(isinstance(order, order_t) for order in orders)

    create_report(process.report_path)

    mail_info = mail.Mail(
        server=get_from_env("SMTP_SERVER"),
        sender=get_from_env("SMTP_SENDER"),
        recipients=get_from_env("SMTP_RECIPIENTS"),
        subject=f'Отчет по процессу "{process.process_name}"',
        attachment_path=process.report_path,
    )

    for order in orders:
        bot.send_message(bot.to_md(order), use_md=True)
        report_status = process_order(colvir, process, order)
        if report_status:
            update_report(
                order=order,
                process=process,
                operation="Создание приказа",
                status=report_status,
            )

    mail.send_mail(mail_info)
