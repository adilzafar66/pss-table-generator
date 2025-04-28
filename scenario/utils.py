import json
from consts.common import DEFAULT_SW_CONFIGS
from etap.api.other.etap_client import EtapClient


def filter_switching_configs(switching_config):
    return True if switching_config in DEFAULT_SW_CONFIGS else False

def get_study_cases(etap: EtapClient):
    study_cases = json.loads(etap.projectdata.getstudycasenames())
    return study_cases['NonDefault']