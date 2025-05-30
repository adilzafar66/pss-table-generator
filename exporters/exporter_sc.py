from openpyxl.cell import Cell
from consts.common import CONFIG_MAP
from consts.columns import SC_CONST_HEADERS
from consts.sheets import WS_SHORT_CIRCUIT, WS_SEQ_IMP
from exporters.exporter import Exporter, SUBHEAD_ROW


class ShortCircuitExporter(Exporter):
    """
    A class to export Short Circuit data to an Excel workbook using the openpyxl library.

    Attributes:
        ansi_data (dict): Data specific to ANSI standards.
        wb (Workbook): The Excel workbook.
    """
    def __init__(self):
        """
        Initializes a new instance of the ShortCircuitExporter class,
        creating a new workbook and setting up the worksheets.
        """
        sc_ws_names = [WS_SHORT_CIRCUIT, WS_SEQ_IMP]
        super().__init__(sc_ws_names, SC_CONST_HEADERS)

    def insert_data(self, ws_index: int, data_type: str, spec_keys: list[str], round_to: int = 2):
        """
        Inserts data into a specified worksheet.

        :param int ws_index: Index of the worksheet in the workbook.
        :param str data_type: Key for the data section to insert.
        :param list[str] spec_keys: Dictionary of specification keys used to extract data.
        :param int round_to: Number of decimal digits to round up to.
        """
        start_row = SUBHEAD_ROW + 1
        ws = self.wb.worksheets[ws_index]
        data_entries = self.ansi_data[data_type]
        ids = list(data_entries.keys())
        values = list(data_entries.values())
        for i, _id in enumerate(ids):
            current_row = ws[start_row + i]
            self._insert_row_data(ws_index, current_row, _id, values[i], spec_keys, round_to)

    def _insert_row_data(self, ws_index: int, ws_row: tuple, entry_id: str, entry_data: dict,
                         spec_keys: list[str], round_to: int):
        """
        Inserts a single row of data into the specified sheet.

        :param int ws_index: Index of the worksheet where the data is being inserted.
        :param int ws_row: The row object where the data is being inserted.
        :param str entry_id: ID of the data being inserted.
        :param dict entry_data: Dictionary of data to be inserted in the row.
        :param list[str] spec_keys: Dictionary of specification keys used to extract data.
        :param int round_to: Number of decimal digits to round up to.
        """
        ws_row[0].value = entry_id.strip()
        ws_row[1].value = round(entry_data['Voltage'], 3)
        data_vals = [entry_data[k] for k in spec_keys]
        for j in range(len(data_vals)):
            self.insert_fault_data(ws_index, ws_row, data_vals[j], round_to, offset=j)

    def insert_fault_data(self, sheet_index: int, row: tuple[Cell, ...], values: dict,
                          round_to: int, offset: int = 0):
        """
        Inserts fault data into specified cells in a row.

        :param int sheet_index: Index of the worksheet in the workbook.
        :param tuple[Cell, ...] row: The row to insert data into.
        :param dict values: A dictionary of fault values by configuration.
        :param int round_to: Number of decimal digits to round up to.
        :param int offset: Offset for column insertion, default is 0.
        """
        for config, fault_val in values.items():
            heading_key = CONFIG_MAP.get(config, config)
            col_index = self._get_col_index(sheet_index, heading_key)
            if col_index:
                cell = row[col_index + offset]
                if isinstance(fault_val, list):
                    mag = round(fault_val[0], round_to)
                    phase = round(fault_val[1])
                    cell.value = f"{mag} ∠ {phase}°"
                else:
                    cell.value = round(fault_val, round_to)
