import json
import os
import pickle
import time
from pathlib import Path
from typing import List

import numpy as np
import pandas as pd
import selenium.webdriver.chrome.service as chrome_service
from selenium.webdriver import Chrome, ChromeOptions
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as ec
from selenium.webdriver.support.ui import WebDriverWait

from robots.data import (
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
)
from robots.notification import TelegramAPI


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
    download_url: str,
    csv_path: str,
    timeout: int = 30,
) -> bool:
    wait.until(
        ec.visibility_of_element_located(
            (By.CSS_SELECTOR, ".cp_menu_section_div_v.cp_menu_simple")
        )
    )
    driver.get(download_url)

    do_reports_exist = (
        len(driver.find_elements(By.CSS_SELECTOR, ".empty_notice_header")) == 0
    )

    if not do_reports_exist:
        print("No report")
        return False

    download_csv_button = wait.until(
        ec.visibility_of_element_located(
            (
                By.CSS_SELECTOR,
                "input.input_button.but_img.bps_submit.sub_h_export_csv",
            )
        )
    )

    download_folder = os.path.dirname(csv_path)
    before_download = set(Path(download_folder).iterdir())

    download_csv_button.click()

    end_time = time.time() + timeout
    while time.time() < end_time:
        time.sleep(1)
        after_download = set(Path(download_folder).iterdir())
        new_files = after_download - before_download
        if new_files:
            old_file_path = new_files.pop().as_posix()
            if os.path.exists(csv_path):
                os.remove(csv_path)
            os.rename(
                old_file_path,
                csv_path,
            )
            return True

    print("Download did not complete within the timeout period.")
    return False


def convert_to_dataclass(process: Process) -> int:
    orders: List[Order] = []
    df = pd.read_csv(process.csv_path, delimiter=";", dtype=str)

    if process.process_type == ProcessType.BUSINESS_TRIP:
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
        with open(
            os.path.join(report_folder, "cities.json"), "r", encoding="utf-8"
        ) as f:
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
    elif process.process_type == ProcessType.VACATION:
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
    elif process.process_type == ProcessType.VACATION_WITHDRAW:
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
    elif process.process_type == ProcessType.FIRING:
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
    is_logged_in: bool,
    bpm_info: BpmInfo,
    process: Process,
    bot: TelegramAPI,
) -> None:
    wait = WebDriverWait(driver, 10)
    if not is_logged_in:
        login(driver, wait, creds=bpm_info.creds)

    if not download_report(
        driver=driver,
        wait=wait,
        download_url=process.download_url,
        csv_path=process.csv_path,
    ):
        raise Exception("Failed to download the report")

    order_count = convert_to_dataclass(process=process)

    if order_count == 1:
        bot.send_message(f"{process.process_type.name} - 1 приказ из BPM")
    elif 2 <= order_count <= 4:
        bot.send_message(f"{process.process_type.name} - {order_count} приказа из BPM")
    else:
        bot.send_message(f"{process.process_type.name} - {order_count} приказов из BPM")
