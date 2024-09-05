import os
import sys
import warnings
from datetime import datetime
from urllib.parse import urljoin

import dotenv

from robots.notification import TelegramAPI

try:
    sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
    import business_trip.colvir
    import vacations.colvir
    import vacation_withdraws.colvir
    import firings.colvir
    from robots import bpm
    from robots.bpm import driver_init
    from robots.data import (
        BpmInfo,
        CredentialsBPM,
        ChromePath,
        ColvirInfo,
        Processes,
        Process,
        ProcessType,
        Buttons,
    )
except ModuleNotFoundError as error:
    raise error


if sys.version_info.major != 3 or sys.version_info.minor != 12:
    raise RuntimeError(f"Python {sys.version_info} is not supported")

warnings.simplefilter(action="ignore", category=UserWarning)
project_folder = os.path.dirname(os.path.dirname(__file__))
dotenv.load_dotenv(os.path.join(project_folder, ".env.test"))


def get_from_env(key: str) -> str:
    value = os.getenv(key)
    assert isinstance(value, str), f"{key} not set in the environment variables"
    return value


class BusinessTripsError(Exception):
    pass


class FiringsError(Exception):
    pass


class VacationsError(Exception):
    pass


class VacationWithdrawsError(Exception):
    pass


def get_process(
    process_type: ProcessType,
    bpm_base_url: str,
    download_folder: str,
    report_parent_folder: str,
    today: str,
) -> Process:
    csv_filename = f"{process_type.name.lower()}_{today}.csv"
    pickle_filename = f"{process_type.name.lower()}_{today}.pkl"
    if process_type == ProcessType.BUSINESS_TRIP:
        report_filename = f"Отчет_командировки_{today}.xlsx"
    elif process_type == ProcessType.BUSINESS_TRIP:
        report_filename = f"Отчет_отпусков_{today}.xlsx"
    elif process_type == ProcessType.BUSINESS_TRIP:
        report_filename = f"Отчет_отзывов_отпусков_{today}.xlsx"
    else:
        report_filename = f"Отчет_увольнений_{today}.xlsx"

    return Process(
        process_type=process_type,
        download_url=urljoin(
            bpm_base_url, f"?s=rep_b&id={process_type.value}&reset_page=1&gid=739"
        ),
        csv_path=os.path.join(download_folder, csv_filename),
        pickle_path=os.path.join(
            report_parent_folder, process_type.name.lower(), pickle_filename
        ),
        report_path=os.path.join(
            report_parent_folder, process_type.name.lower(), report_filename
        ),
        today=today,
    )


def main(bot: TelegramAPI) -> None:
    data_folder = os.path.join(project_folder, "data")
    os.makedirs(data_folder, exist_ok=True)

    report_parent_folder = os.path.join(data_folder, "reports")
    os.makedirs(report_parent_folder, exist_ok=True)

    download_folder = os.path.join(data_folder, "downloads")
    os.makedirs(download_folder, exist_ok=True)

    bpm_info = BpmInfo(
        creds=CredentialsBPM(
            user=get_from_env("BPM_USER"), password=get_from_env("BPM_PASSWORD")
        ),
        chrome_path=ChromePath(
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

    today = datetime.now().strftime("%d.%m.%y")

    bot.send_message(
        f"Старт процесса за {today}\n"
        f'"Командировки, отпуска, отзывы из отпуска и увольнения"'
    )

    bpm_base_url = get_from_env("BPM_BASE_URL")
    processes = Processes(
        business_trip=get_process(
            ProcessType.BUSINESS_TRIP,
            bpm_base_url,
            download_folder,
            report_parent_folder,
            today,
        ),
        vacation=get_process(
            ProcessType.VACATION,
            bpm_base_url,
            download_folder,
            report_parent_folder,
            today,
        ),
        vacation_withdraw=get_process(
            ProcessType.VACATION_WITHDRAW,
            bpm_base_url,
            download_folder,
            report_parent_folder,
            today,
        ),
        firing=get_process(
            ProcessType.FIRING,
            bpm_base_url,
            download_folder,
            report_parent_folder,
            today,
        ),
    )

    with driver_init(bpm_info=bpm_info) as driver:
        is_logged_in = False
        for process in processes:
            bpm.run(
                driver=driver,
                is_logged_in=is_logged_in,
                bpm_info=bpm_info,
                process=process,
                bot=bot,
            )
            is_logged_in = True

    buttons = Buttons()

    # business_trip.colvir.run(colvir_info, processes.business_trip, buttons)
    # vacations.colvir.run(colvir_info, processes.vacation, buttons)
    # vacation_withdraws.colvir.run(colvir_info, processes.vacation_withdraw, buttons)
    # firings.colvir.run(colvir_info, processes.firing, buttons)

    bot.send_message("Успешное окончание процесса")


if __name__ == "__main__":
    telegram_bot = TelegramAPI()
    main(bot=telegram_bot)
