import os
import shutil
from datetime import datetime
from typing import Optional

import pandas as pd

PATH = r"C:\Users\robotX4\Desktop\business-trip-py\data\reports"


def get_header_row(current_day_report_path: str):
    df = pd.read_excel(current_day_report_path, header=None)

    header_row_index: Optional[int] = None
    for index, row in df.iterrows():
        if list(row)[0] == "Имя сотрудника":
            header_row_index = int(str(index))
            break
    return header_row_index


def main():
    downloads_folder = r"C:\Users\robotX4\Downloads"
    work_folder = PATH

    today = datetime.now().strftime("%d.%m.%y")
    current_day_report_path = os.path.join(work_folder, f"orders_{today}.xlsx")

    for download_file in os.listdir(downloads_folder):
        if download_file.startswith("rep") and download_file.endswith(".xlsx"):
            shutil.copyfile(
                os.path.join(downloads_folder, download_file), current_day_report_path
            )

    skiprows = get_header_row(current_day_report_path)
    if not skiprows:
        raise ValueError(f"Header not found in {current_day_report_path}")

    df = pd.read_excel(current_day_report_path, skiprows=skiprows)
    df = df.rename(
        columns={
            "Имя сотрудника": "employee_fullname",
            "Номер приказа": "order_number",
            "Дата подписания": "sign_date",
            "Дата начала": "start_date",
            "Дата окончания": "end_date",
            "Место командирования": "trip_place",
            "Цель командировки": "trip_target",
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
    df.loc[:, "sign_date"] = df["sign_date"].str.replace(".", "")
    df.loc[:, "start_date"] = df["start_date"].str.replace(".", "")
    df.loc[:, "end_date"] = df["end_date"].str.replace(".", "")

    df.to_json(
        os.path.join(work_folder, f"orders_{today}.json"),
        orient="records",
        force_ascii=False,
        indent=2,
    )


def foo():
    today = datetime.now().strftime("%d-%m-%Y")
    work_folder = PATH

    report_name = f"Отчет_командировки_{today}.xlsx"
    excel_file = os.path.join(work_folder, report_name)

    data = {
        "Дата": today,
        "Сотрудник": "Нет",
        "Операция": "Сбор информации с заявок",
        "Номер приказа": "Нет",
        "Статус": "Заявок нет",
    }

    df = pd.DataFrame([data])
    df.to_excel(excel_file, index=False)


if __name__ == "__main__":
    main()
