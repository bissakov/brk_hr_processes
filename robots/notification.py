import json
import logging
import os
import traceback
import urllib.parse
from functools import wraps
from typing import Dict, Optional, Tuple, Callable, Any, Union, List

import requests
import requests.adapters

from robots.data import Order


def get_secrets() -> Tuple[str, str]:
    token = os.getenv("TOKEN")
    if token is None:
        raise EnvironmentError("Environment variable 'TOKEN' is not set.")

    chat_id = os.getenv("CHAT_ID")
    if chat_id is None:
        raise EnvironmentError("Environment variable 'CHAT_ID' is not set.")

    return token, chat_id


class TelegramAPI:
    def __init__(self) -> None:
        self.session = requests.Session()
        self.session.mount("http://", requests.adapters.HTTPAdapter(max_retries=5))
        self.token, self.chat_id = get_secrets()
        self.api_url = f"https://api.telegram.org/bot{self.token}/"

    def reload_session(self) -> None:
        self.session = requests.Session()
        self.session.mount("http://", requests.adapters.HTTPAdapter(max_retries=5))

    def send_message(
        self, message: str, use_session: bool = True, use_md: bool = False
    ) -> bool:
        send_data: Dict[str, Optional[str]] = {
            "chat_id": self.chat_id,
        }

        if use_md:
            send_data["parse_mode"] = "MarkdownV2"

        files = None

        url = urllib.parse.urljoin(self.api_url, "sendMessage")
        send_data["text"] = message

        if use_session:
            response = self.session.post(url, data=send_data, files=files)
        else:
            response = requests.post(url, data=send_data, files=files)

        method = url.split("/")[-1]
        data = "" if not hasattr(response, "json") else response.json()
        logging.info(
            f"Response for '{method}': {response}\n"
            f"Is 200: {response.status_code == 200}\n"
            f"Data: {data}"
        )
        response.raise_for_status()
        return response.status_code == 200

    def send_with_retry(
        self,
        message: str,
    ) -> bool:
        retry = 0
        while retry < 5:
            try:
                use_session = retry < 5
                success = self.send_message(message, use_session)
                return success
            except (
                requests.exceptions.ConnectionError,
                requests.exceptions.SSLError,
                requests.exceptions.HTTPError,
            ) as e:
                self.reload_session()
                logging.exception(e)
                logging.warning(f"{e} intercepted. Retry {retry + 1}/10")
                retry += 1

        return False

    @staticmethod
    def to_md(obj: Union[Dict[str, Any], List[Any], Order]) -> str:
        try:
            if isinstance(obj, dict) or isinstance(obj, list):
                obj_json = json.dumps(obj, ensure_ascii=False, indent=2)
            elif isinstance(obj, Order):
                obj_json = json.dumps(
                    obj.as_dict_short(),
                    indent=2,
                    ensure_ascii=False,
                )
            else:
                raise ValueError(f"obj is of the wrong type - {type(obj)}")

            return f"```json\n{obj_json}\n```"

        except (Exception, BaseException) as error:
            logging.exception(error)
            return ""


def handle_error(func: Callable[..., Any]) -> Callable[..., Any]:
    @wraps(func)
    def wrapper(*args, **kwargs) -> Any:
        bot: Optional[TelegramAPI] = kwargs.get("bot")

        try:
            return func(*args, **kwargs)
        except KeyboardInterrupt as error:
            raise error
        except (Exception, BaseException) as error:
            logging.exception(error)
            error_msg = traceback.format_exc()

            if bot:
                bot.send_message(error_msg)
            raise error

    return wrapper
