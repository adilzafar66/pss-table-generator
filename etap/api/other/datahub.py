import requests
from .. import settings

requests.packages.urllib3.disable_warnings()


def get(url_absolute: str, token: str = None) -> str:
    """
    Makes an HTTP GET request to ETAP DataHub using the given absolute URL.

    :param url_absolute: Absolute URL to request
    :param token: Optional authorization token
    :return: Response text from the GET request
    """
    headers = {"Authorization": token} if token else {}
    response = requests.get(url_absolute, verify=not settings.https_enable, headers=headers)
    return response.text


def post(url_absolute: str, data: dict, token: str = None, headers: dict = None) -> str:
    """
    Makes an HTTP POST request to ETAP DataHub using the given absolute URL.

    :param url_absolute: Absolute URL to request
    :param data: Dictionary to send in the request body
    :param token: Optional authorization token
    :param headers: Optional additional headers
    :return: Response text from the POST request
    """
    final_headers = headers or {}
    if token:
        final_headers["Authorization"] = token

    response = requests.post(
        url_absolute,
        json=data,
        verify=not settings.https_enable,
        headers=final_headers
    )
    return response.text
