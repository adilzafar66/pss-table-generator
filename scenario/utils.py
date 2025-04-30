from consts.common import DEFAULT_SW_CONFIGS


def filter_switching_configs(switching_config):
    return True if switching_config in DEFAULT_SW_CONFIGS else False
