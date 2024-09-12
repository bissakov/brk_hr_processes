import json
import logging
import os
import pickle
import time
from datetime import datetime
from pathlib import Path
from typing import List

import numpy as np
import pandas as pd
import selenium.webdriver.chrome.service as chrome_service
from selenium.webdriver import Chrome, ChromeOptions
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as ec
from selenium.webdriver.support.ui import WebDriverWait

from src.data import (
    CredentialsBPM,
    BpmInfo,
    Date,
    ProcessType,
    VacationOrder,
    BusinessTripOrder,
    Process,
    VacationWithdrawOrder,
    Order,
    FiringOrder,
    MentorshipOrder,
)
from src.notification import TelegramAPI


def driver_init(bpm_info: BpmInfo) -> Chrome:
    service = chrome_service.Service(executable_path=bpm_info.chrome_path.driver_path)
    options = ChromeOptions()
    options.binary_location = bpm_info.chrome_path.binary_path
    prefs = {"profile.default_content_setting_values.notifications": 2}
    options.add_experimental_option("prefs", prefs)
    options.add_argument("--start-maximized")
    options.add_experimental_option("excludeSwitches", ["enable-logging"])
    options.add_argument("--log-level=3")
    options.add_experimental_option(
        "prefs",
        {
            "download.default_directory": bpm_info.download_folder,
            "download.directory_upgrade": True,
            "download.prompt_for_download": False,
        },
    )
    driver = Chrome(service=service, options=options)
    return driver


def login(driver: Chrome, wait: WebDriverWait, creds: CredentialsBPM) -> None:
    driver.get("https://bpmtest.kdb.kz/")

    user_input = wait.until(ec.presence_of_element_located((By.NAME, "u_login")))
    user_input.send_keys(creds.user)

    psw_input = wait.until(ec.presence_of_element_located((By.NAME, "pwd")))
    psw_input.send_keys(creds.password)

    submit_button = wait.until(ec.presence_of_element_located((By.NAME, "submit")))
    submit_button.click()


def download_report(
    driver: Chrome,
    wait: WebDriverWait,
    process: Process,
    timeout: int = 30,
) -> bool:
    wait.until(
        ec.visibility_of_element_located(
            (By.CSS_SELECTOR, ".cp_menu_section_div_v.cp_menu_simple")
        )
    )
    driver.get(process.download_url)

    if process.process_type == ProcessType.VACATION_ADD_PAY:
        order_date_input = wait.until(
            ec.visibility_of_element_located(
                (By.CSS_SELECTOR, 'input[data-col-id="656198"]')
            )
        )
        today = datetime.strptime(process.today, "%d.%m.%y").strftime("%d.%m.%Y")
        order_date_input.send_keys(today)
        time.sleep(2)

        download_csv_button_selector = 'a[title="Экспорт в CSV"]'
    else:
        download_csv_button_selector = (
            "input.input_button.but_img.bps_submit.sub_h_export_csv"
        )

    do_reports_exist = (
        len(driver.find_elements(By.CSS_SELECTOR, ".empty_notice_header")) == 0
    )

    if not do_reports_exist:
        logging.info("No reports")
        return False

    download_csv_button = wait.until(
        ec.visibility_of_element_located(
            (
                By.CSS_SELECTOR,
                download_csv_button_selector,
            )
        )
    )

    download_folder = os.path.dirname(os.path.dirname(process.csv_path))
    before_download = set(Path(download_folder).iterdir())

    download_csv_button.click()

    end_time = time.time() + timeout
    while time.time() < end_time:
        time.sleep(1)
        after_download = set(Path(download_folder).iterdir())
        new_files = after_download - before_download
        if new_files:
            old_file_path = new_files.pop().as_posix()
            if os.path.exists(process.csv_path):
                os.remove(process.csv_path)
            os.rename(
                old_file_path,
                process.csv_path,
            )
            return True

    print("Download did not complete within the timeout period.")
    return False


def convert_to_dataclass(process: Process, is_empty: bool) -> int:
    orders: List[Order] = []

    if is_empty:
        with open(process.pickle_path, "wb") as f:
            pickle.dump(orders, f)

        orders_json_path = process.pickle_path.replace(".pkl", ".json")
        with open(orders_json_path, "w", encoding="utf-8") as f:
            json.dump(
                [order.as_dict() for order in orders], f, ensure_ascii=False, indent=2
            )

        return len(orders)

    df = pd.read_csv(process.csv_path, delimiter=";", dtype=str)

    match process.process_type:
        case ProcessType.BUSINESS_TRIP:
            df = df.rename(
                columns={
                    "Имя сотрудника": "employee_fullname",
                    "Номер приказа": "order_number",
                    "Дата подписания": "sign_date",
                    "Дата начала": "start_date",
                    "Дата окончания": "end_date",
                    "Место командирования": "trip_place",
                    "Цель командировки": "trip_reason",
                    "Номер основного приказа": "main_order_number",
                    "Дата начала основного приказа": "main_order_start_date",
                    "Имя замещающего сотрудника": "deputy_fullname",
                }
            )

            df = df.dropna(subset=["employee_fullname"])
            df = df.dropna(subset=["sign_date"])

            assert len(df.columns) == 10, "Странное кол-во колонок"

            df.loc[:, "employee_names"] = df["employee_fullname"].str.split()
            df.loc[:, "deputy_names"] = df["deputy_fullname"].str.split()

            df["sign_date"] = pd.to_datetime(df["sign_date"], format="%d.%m.%Y")
            df["start_date"] = pd.to_datetime(df["start_date"], format="%d.%m.%Y")
            df["end_date"] = pd.to_datetime(df["end_date"], format="%d.%m.%Y")
            df["main_order_start_date"] = pd.to_datetime(
                df["main_order_start_date"], format="%d.%m.%Y"
            )
            df = df.replace({np.nan: None})

            report_folder = os.path.dirname(process.report_path)
            cities_mapping_path = os.path.join(
                os.path.dirname(os.path.dirname(os.path.dirname(report_folder))),
                "cities.json",
            )
            with open(cities_mapping_path, "r", encoding="utf-8") as f:
                cities = json.load(f)

            for _, order_dict in df.iterrows():
                trip_place = order_dict["trip_place"]
                trip_bpm = trip_place.replace("город ", "").replace("г. ", "")
                trip_bpm = trip_bpm.split(",")[0]
                trip_code = cities.get(trip_bpm)
                if not trip_code:
                    trip_code = ""
                else:
                    trip_code = trip_code.replace(f".{trip_bpm}", "")

                order = BusinessTripOrder(
                    employee_fullname=order_dict["employee_fullname"],
                    employee_names=order_dict["employee_names"],
                    order_number=order_dict["order_number"],
                    sign_date=Date(order_dict["sign_date"]),
                    start_date=Date(order_dict["start_date"]),
                    end_date=Date(order_dict["end_date"]),
                    trip_place=trip_place,
                    trip_code=trip_code,
                    trip_reason=order_dict["trip_reason"],
                    main_order_number=order_dict["main_order_number"],
                    main_order_start_date=Date(order_dict["main_order_start_date"]),
                    deputy_fullname=order_dict["deputy_fullname"],
                    deputy_names=order_dict["deputy_names"],
                )
                orders.append(order)

        case ProcessType.VACATION:
            df = df.rename(
                columns={
                    "Имя сотрудника": "employee_fullname",
                    "Тип приказа": "order_type",
                    "Дата начала": "start_date",
                    "Дата окончания": "end_date",
                    "Номер приказа": "order_number",
                    "Имя замещающего": "deputy_fullname",
                    "Доплата": "surcharge",
                    "Начало замещения": "substitution_start",
                    "Конец замещения": "substitution_end",
                }
            )

            df = df.dropna(subset=["employee_fullname"])

            df.loc[:, "employee_names"] = df["employee_fullname"].str.split()
            df.loc[:, "deputy_names"] = df["deputy_fullname"].str.split()

            df["start_date"] = pd.to_datetime(df["start_date"], format="%d.%m.%Y")
            df["end_date"] = pd.to_datetime(df["end_date"], format="%d.%m.%Y")
            df = df.replace({np.nan: None})

            for _, order_dict in df.iterrows():
                order_type_long = order_dict["order_type"].lower()
                if "ежегодный" in order_type_long:
                    order_type = "О"
                elif "учебный" in order_type_long:
                    order_type = "УО"
                else:
                    order_type = "Б/С"

                order_number = order_dict["order_number"]
                order_number = "1" if not order_number else order_number

                order = VacationOrder(
                    employee_fullname=order_dict["employee_fullname"],
                    employee_names=order_dict["employee_names"],
                    order_type=order_type,
                    start_date=Date(order_dict["start_date"]),
                    end_date=Date(order_dict["end_date"]),
                    order_number=order_number,
                    deputy_fullname=order_dict["deputy_fullname"],
                    deputy_names=order_dict["deputy_names"],
                    surcharge=order_dict["surcharge"],
                    substitution_start=order_dict["substitution_start"],
                    substitution_end=order_dict["substitution_end"],
                )
                orders.append(order)

        case ProcessType.VACATION_WITHDRAW:
            df = df.rename(
                columns={
                    "Имя сотрудника": "employee_fullname",
                    "Дата отзыва": "withdraw_date",
                    "Тип приказа": "order_type",
                    "Номер приказа": "order_number",
                }
            )
            df = df.dropna(subset=["employee_fullname"])
            df = df.dropna(subset=["withdraw_date"])

            df.loc[:, "employee_names"] = df["employee_fullname"].str.split()

            df["withdraw_date"] = pd.to_datetime(df["withdraw_date"], format="%d.%m.%Y")
            df = df.replace({np.nan: None})

            for _, order_dict in df.iterrows():
                order = VacationWithdrawOrder(
                    employee_fullname=order_dict["employee_fullname"],
                    employee_names=order_dict["employee_names"],
                    order_type=order_dict["order_type"],
                    order_number=order_dict["order_number"],
                    withdraw_date=Date(order_dict["withdraw_date"]),
                )
                orders.append(order)

        case ProcessType.FIRING:
            df = df.rename(
                columns={
                    "Имя сотрудника": "employee_fullname",
                    "Дата увольнения": "firing_date",
                    "Причина увольнения": "firing_reason",
                    "Номер приказа": "order_number",
                    "Компенсация": "compensation",
                }
            )
            df = df.dropna(subset=["employee_fullname"])
            df = df.dropna(subset=["firing_date"])

            df.loc[:, "employee_names"] = df["employee_fullname"].str.split()

            df["firing_date"] = pd.to_datetime(df["firing_date"], format="%d.%m.%Y")
            df = df.replace({np.nan: None})

            for _, order_dict in df.iterrows():
                order = FiringOrder(
                    employee_fullname=order_dict["employee_fullname"],
                    employee_names=order_dict["employee_names"],
                    firing_reason=order_dict["firing_reason"],
                    order_number=order_dict["order_number"],
                    compensation=order_dict["compensation"],
                    firing_date=Date(order_dict["firing_date"]),
                )
                orders.append(order)

        case ProcessType.MENTORSHIP:
            df = df[
                [
                    "Имя сотрудника",
                    "Первый рабочий день",
                    "Начало договора",
                    "Окончание договора",
                    "ФИО ментора",
                    "Номер приказа о менторстве",
                    "Начало менторства",
                    "Окончание менторства",
                    "Дата создания",
                ]
            ]

            df = df.rename(
                columns={
                    "Имя сотрудника": "employee_fullname",
                    # "ИИН": "iin",
                    # "Дата рождения": "birth_date",
                    # "Пол": "gender",
                    # "Национальность": "nationality",
                    # "Гражданство": "citizenship",
                    # "Семейное положение": "marital_status",
                    # "Количество детей": "children_count",
                    # "Номер приказа": "order_number",
                    # "Должность": "position",
                    # "Подразделение": "branch",
                    "Первый рабочий день": "work_start_date",
                    # "График работы": "work_schedule",
                    # "Тип работы": "work_type",
                    # "Характер работы": "work_nature",
                    "Начало договора": "contract_start_date",
                    "Окончание договора": "contract_end_date",
                    "ФИО ментора": "mentor_fullname",
                    "Номер приказа о менторстве": "mentrorship_order_number",
                    "Начало менторства": "mentorship_start_date",
                    "Окончание менторства": "mentorship_end_date",
                    "Дата создания": "creation_date",
                }
            )
            df = df.dropna(subset=["employee_fullname"])
            df = df.dropna(subset=["work_start_date"])
            df = df.dropna(subset=["contract_start_date"])
            df = df.dropna(subset=["mentor_fullname"])

            df.loc[:, "employee_names"] = df["employee_fullname"].str.split()
            df.loc[:, "mentor_names"] = df["mentor_fullname"].str.split()

            for col in [
                "work_start_date",
                "contract_start_date",
                "contract_start_date",
                "mentorship_start_date",
                "mentorship_end_date",
            ]:
                df[col] = pd.to_datetime(df[col], format="%d.%m.%Y")
            df = df.replace({np.nan: None})

            for _, order_dict in df.iterrows():
                order = MentorshipOrder(
                    employee_fullname=order_dict["employee_fullname"],
                    work_start_date=order_dict["work_start_date"],
                    contract_start_date=order_dict["contract_start_date"],
                    contract_end_date=order_dict["contract_end_date"],
                    mentor_fullname=order_dict["mentor_fullname"],
                    mentrorship_order_number=order_dict["mentrorship_order_number"],
                    mentorship_start_date=order_dict["mentorship_start_date"],
                    mentorship_end_date=order_dict["mentorship_end_date"],
                    creation_date=order_dict["creation_date"],
                )
                orders.append(order)

        case _:
            ValueError(
                f"Unknown process type: ProcessType(name={process.process_type.name}, "
                f"value={process.process_type.value})"
            )

    with open(process.pickle_path, "wb") as f:
        pickle.dump(orders, f)

    orders_json_path = process.pickle_path.replace(".pkl", ".json")
    with open(orders_json_path, "w", encoding="utf-8") as f:
        json.dump(
            [order.as_dict() for order in orders], f, ensure_ascii=False, indent=2
        )

    return len(orders)


def run(
    driver: Chrome,
    creds: CredentialsBPM,
    process: Process,
    bot: TelegramAPI,
    is_logged_in: bool,
) -> None:
    wait = WebDriverWait(driver, 10)
    if not is_logged_in:
        login(driver, wait, creds=creds)

    is_empty = not download_report(
        driver=driver,
        wait=wait,
        process=process,
    )
    order_count = convert_to_dataclass(process=process, is_empty=is_empty)

    bot.send_message(
        f"{process.process_type.name} - {order_count} - кол-во приказов из BPM"
    )
