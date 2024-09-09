import os
import sys

project_folder = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.append(project_folder)

from src.notification import TelegramAPI
from src import process_manager

if __name__ == "__main__":
    telegram_bot = TelegramAPI()
    process_manager.run(bot=telegram_bot)
