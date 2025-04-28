PROGRAM_TITLE = 'ETAP Report Generator'
PORT_FILE = 'port_number.pickle'

HTTP = 'http'
HTTPS = 'https'
HEADER_ROW = 1
SUBHEAD_ROW = 2
ROUND_DIGITS = 2

DEFAULT_SW_CONFIGS = [
    'Normal',
    'Present',
    'PRES',
    'Ultimate',
    'ULT',
    'Generator',
    'GEN'
]

TYPE_MAP = {
    'Fuse': 'FUSE',
    'SPDT Switch': 'DOUBLESWITCH',
    'SPST Switch': 'SINGLESWITCH',
    'Molded Case': 'LVCB'
}

CONFIG_MAP = {
    'PRES': 'Present',
    'STBY': 'Standby',
    'ULT': 'Ultimate',
    'GEN': 'Generator',
    'MAX': 'Maximum',
    'MIN': 'Minimum',
}

INV_CONFIG_MAP = {
    'Present': 'PRES',
    'Normal': 'PRES',
    'Standby': 'STBY',
    'Ultimate': 'ULT',
    'Generator': 'GEN'
}
