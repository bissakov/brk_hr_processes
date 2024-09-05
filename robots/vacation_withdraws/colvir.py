import os
import pickle
from typing import List

from robots.data import (
    ColvirInfo,
    Buttons,
    Process,
    VacationWithdrawOrder,
)
from robots.utils.colvir_utils import (
    Colvir,
)
from robots.utils.utils import (
    kill_all_processes,
    create_report,
)


def run(colvir_info: ColvirInfo, process: Process, buttons: Buttons) -> None:
    with open(process.pickle_path, "rb") as f:
        orders: List[VacationWithdrawOrder] = pickle.load(f)

    assert all(isinstance(order, VacationWithdrawOrder) for order in orders)

    report_folder = os.path.dirname(process.report_path)
    create_report(process.report_path)

    kill_all_processes(proc_name="COLVIR")

    colvir = Colvir(colvir_info=colvir_info)
    app = colvir.app

    for i, order in enumerate(orders):
        pass
