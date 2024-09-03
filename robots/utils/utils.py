import os

import pandas as pd
import psutil


def kill_all_processes(proc_name: str) -> None:
    for proc in psutil.process_iter():
        try:
            if proc_name in proc.name():
                proc.terminate()
        except (psutil.AccessDenied, psutil.NoSuchProcess):
            continue


def create_report(report_file_path: str):
    if not os.path.exists(report_file_path):
        df = pd.DataFrame(
            {
                "Дата": [],
                "Сотрудник": [],
                "Операция": [],
                "Номер приказа": [],
                "Статус": [],
            }
        )
        df.to_excel(report_file_path, index=False)


def update_report(
    person_name: str,
    order_number: str,
    report_file_path: str,
    today: str,
    operation: str,
    status: str,
):
    df = pd.read_excel(report_file_path)

    if not (
        (df["Дата"] == today)
        & (df["Сотрудник"] == person_name)
        & (df["Операция"] == operation)
        & (df["Номер приказа"] == order_number)
    ).any():
        new_row = {
            "Дата": today,
            "Сотрудник": person_name,
            "Операция": operation,
            "Номер приказа": order_number,
            "Статус": status,
        }
        df.loc[len(df)] = new_row
        df.to_excel(report_file_path, index=False)


def does_order_exist(orders_file_path: str, order_type: str, order_number: str) -> bool:
    df = pd.read_excel(orders_file_path, skiprows=1)

    order_exists = (
        (df["Вид приказа"] == order_type) & (df["Номер приказа"] == order_number)
    ).any()

    return order_exists
