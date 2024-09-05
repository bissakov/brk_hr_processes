import logging
import os
import urllib.parse
from typing import Dict, Optional, Tuple

import requests
import requests.adapters


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
        self,
        message: str,
        use_session: bool = True,
    ) -> bool:
        send_data: Dict[str, Optional[str]] = {"chat_id": self.chat_id}
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
