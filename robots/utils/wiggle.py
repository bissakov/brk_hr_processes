import random
from typing import Tuple

import pyautogui

pyautogui.FAILSAFE = False


def wiggle_mouse(duration: int) -> None:
    max_wiggles = random.randint(4, 9)
    step_sleep = duration / max_wiggles

    for _ in range(1, max_wiggles):
        coords = get_random_coords()
        pyautogui.moveTo(x=coords[0], y=coords[1], duration=step_sleep)


def get_random_coords() -> Tuple[int, int]:
    screen = pyautogui.size()
    width = screen[0]
    height = screen[1]

    return random.randint(100, width - 200), random.randint(100, height - 200)
