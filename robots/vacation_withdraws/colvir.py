import pickle
from typing import List

from robots.data import (
    Process,
    VacationWithdrawOrder,
)
from robots.notification import TelegramAPI
from robots.utils.colvir_utils import (
    Colvir,
)
from robots.utils.utils import (
    create_report,
)


def run(colvir: Colvir, process: Process, bot: TelegramAPI) -> None:
    with open(process.pickle_path, "rb") as f:
        orders: List[VacationWithdrawOrder] = pickle.load(f)

    assert all(isinstance(order, VacationWithdrawOrder) for order in orders)

    create_report(process.report_path)

    for i, order in enumerate(orders):
        pass
