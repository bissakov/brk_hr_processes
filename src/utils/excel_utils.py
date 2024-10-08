import os
from contextlib import contextmanager

import psutil
import win32com.client as win32


def kill_all_processes(proc_name: str) -> None:
    for proc in psutil.process_iter():
        try:
            if proc_name in proc.name():
                proc.terminate()
        except (psutil.AccessDenied, psutil.NoSuchProcess):
            continue


@contextmanager
def dispatch(application: str) -> None:
    app = win32.Dispatch(application)
    app.DisplayAlerts = False
    try:
        yield app
    finally:
        kill_all_processes(proc_name="EXCEL")


@contextmanager
def workbook_open(excel: win32.Dispatch, file_path: str) -> None:
    wb = excel.Workbooks.Open(file_path)
    try:
        yield wb
    finally:
        wb.Close()


def xls_to_xlsx(source: str, dest: str):
    if os.path.exists(dest):
        os.remove(dest)
    with dispatch(application="Excel.Application") as excel:
        with workbook_open(excel=excel, file_path=source) as wb:
            wb.SaveAs(dest, FileFormat=51)
    os.remove(source)
