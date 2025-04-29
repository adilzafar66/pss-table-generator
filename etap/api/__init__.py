from typing import Optional
from etap.api.other.etap_client import EtapClient
import requests


def connect(base_address: str, project_name: Optional[str] = None) -> EtapClient:
    """
    Establish a connection with ETAP. This should be called before any other ETAP API call.

    :param str base_address: DataHub base address (e.g., 'http://localhost:50000')
    :param Optional[str] project_name: Name of the project to connect to (optional)
    :return EtapClient: An instance of EtapClient for communicating with ETAP
    """
    try:
        client = EtapClient()
        client.connect(base_address, project_name)
        return client
    except (requests.RequestException, ConnectionError) as error:
        raise RuntimeError(f"Failed to connect to ETAP at {base_address}: {error}") from error
