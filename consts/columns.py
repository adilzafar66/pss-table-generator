SC_CONST_HEADERS = ['Device', None]
SC_FAULT_CONST_COLS = ['Bus', 'kV']
SC_FAULT_VAR_COLS = ['3PH', 'LG', 'LL', 'LLG']
SC_IMP_CONST_COLS = ['Bus', 'kV']
SC_IMP_VAR_COLS = ['Pos. Seq. R', 'Pos Seq. X', 'Zero Seq. R', 'Zero Seq. X']

DD_CONST_HEADERS = ['Device', 'Device Capability']
DD_MOM_CONST_COLS = ['ID', 'Nominal kV', 'Type']
DD_MOM_VAR_COLS = ['Symm. kA rms', 'Asymm. kA rms']
DD_INT_CONST_COLS = ['ID', 'Voltage', 'Bus', 'Device']
DD_INT_VAR_COLS = ['Int. Adj. Symm. kA']
DD_INT_IEC_VAR_COLS = ['Ib Sym. kA', 'Ib Asym. kA']

AF_CONST_COLS = [
    'ID',
    'kV',
    'Type',
    'Electrode Configuration',
    'Working Distance LL (in)',
    'LAB to Fixed Part (ft-in)',
    'RAB (ft-in)',
    'Output Rpt.',
    'Configuration',
    'Total Energy (cal/cm²)',
    'AFB (ft-in)',
    'Final FCT (sec)',
    'Source PD ID',
    '% Ia Variation',
    "Total Ia'' (kA)",
    'Source PD Ia at FCT (kA)',
    'Total Ibf at FCT (kA)',
    'Source PD Ibf at FCT (kA)'
]

AF_SI_CONST_COLS = [
    'ID',
    'kV',
    'Type',
    'Electrode Configuration',
    'Working Distance LL (m)',
    'LAB to Fixed Part (m)',
    'RAB (m)',
    'Output Rpt.',
    'Configuration',
    'Total Energy (cal/cm²)',
    'AFB (m)',
    'Final FCT (sec)',
    'Source PD ID',
    '% Ia Variation',
    "Total Ia'' (kA)",
    'Source PD Ia at FCT (kA)',
    'Total Ibf at FCT (kA)',
    'Source PD Ibf at FCT (kA)'
]

AF_COL_INDICES = {
    'ecf': 2,
    'lab': 4,
    'rab': 5,
    'rep': 6,
    'con': 7,
    'ie': 8,
    'afb': 9,
    'fct': 10,
    'la_var': 12
}