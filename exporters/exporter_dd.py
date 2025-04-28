from openpyxl import Workbook
from openpyxl.cell import Cell
from consts.columns import DD_CONST_HEADERS
from consts.common import CONFIG_MAP
from consts.styles import FILL_ROW_RED, FILL_ROW_ORANGE
from consts.sheets import WS_ANSI_MOM, WS_ANSI_INT, WS_IEC_INT
from exporters.exporter import Exporter, SUBHEAD_ROW


class DeviceDutyExporter(Exporter):
    """
    A class to export Device Duty data to an Excel workbook using the openpyxl library.

    Attributes:
        ansi_data (dict): Data specific to ANSI standards.
        iec_data (dict): Data specific to IEC standards.
        wb (Workbook): The Excel workbook.
    """

    def __init__(self):
        """
        Initializes a new instance of the DeviceDutyExporter class,
        creating a new workbook and setting up the worksheets.
        """
        dd_ws_names = [WS_ANSI_MOM, WS_ANSI_INT, WS_IEC_INT]
        super().__init__(dd_ws_names, DD_CONST_HEADERS)

    @staticmethod
    def highlight_high_duty(cell: Cell, fault_val: float, cap_val: float):
        """
        Highlights cells where the fault value exceeds the capability value.

        :param Cell cell: The cell to apply formatting to.
        :param float fault_val: The fault value to compare.
        :param float cap_val: The capability value to compare against.
        """
        if fault_val > cap_val:
            cell.fill = FILL_ROW_RED

    @staticmethod
    def highlight_series_rated(cell, is_series_rated: bool | None):
        """
        Highlights cells that are series rated.

        :param Cell cell: The cell to apply formatting to.
        :param bool | None is_series_rated: Indicates whether the cell should be series rated.
        """
        if is_series_rated:
            cell.fill = FILL_ROW_ORANGE

    def insert_data(self, ws_index: int, data_type: str, spec_keys: dict, dataset: str = 'ansi'):
        """
        Inserts data into a specified worksheet.

        :param int ws_index: Index of the worksheet in the workbook.
        :param str data_type: Key for the data section to insert.
        :param dict spec_keys: Dictionary of specification keys used to extract data.
        :param str dataset: The dataset to use ('ansi' or 'iec'). Default is 'ansi'.
        """
        data = self.ansi_data if dataset == 'ansi' else self.iec_data
        ws = self.wb.worksheets[ws_index]
        start_row = SUBHEAD_ROW + 1
        data_entries = data[data_type]
        ids = list(data_entries.keys())
        values = list(data_entries.values())

        for i, _id in enumerate(ids):
            current_row = ws[start_row + i]
            self._insert_row_data(ws_index, current_row, _id, values[i], spec_keys)

    def _insert_row_data(self, ws_index: int, ws_row: tuple, entry_id: str, entry_data: dict, spec_keys: dict):
        """
        Inserts a single row of data into the specified sheet.

        :param int ws_index: Index of the worksheet where the data is being inserted.
        :param int ws_row: The row object where the data is being inserted.
        :param str entry_id: ID of the data being inserted.
        :param dict entry_data: Dictionary of data to be inserted in the row.
        :param dict spec_keys: Specification keys for fault and capability data.
        """
        clean_id = entry_id.strip()
        assumed = entry_data.get('Assumed')
        ws_row[0].value = clean_id + '*' if assumed else clean_id
        ws_row[1].value = round(entry_data['Voltage'], 3)
        ws_row[2].value = entry_data.get(spec_keys['Type'], '').strip()
        ws_row[3].value = entry_data.get('Device', '').strip()
        self._insert_fault_and_cap_data(ws_index, ws_row, entry_data, spec_keys)

    def _insert_fault_and_cap_data(self, ws_index: int, ws_row: tuple, entry_data: dict, spec_keys: dict):
        """
        Inserts fault and capability data for a row.

        :param int ws_index: Index of the worksheet in the workbook.
        :param tuple ws_row: Row where data will be inserted.
        :param dict entry_data: Data containing fault and capability values.
        :param dict spec_keys: Specification keys.
        """
        all_fault_vals = [entry_data[k] for k in spec_keys['Fault']]
        cap_vals = [entry_data[k] for k in spec_keys['Capability']]

        for j, fault_vals in enumerate(all_fault_vals):
            self.insert_fault_data(ws_index, ws_row, fault_vals, cap_vals[j], j)
        last_col_index = self._get_col_index(ws_index, DD_CONST_HEADERS[1])

        for j, cap_val in enumerate(cap_vals):
            ws_row[last_col_index + j].value = round(cap_val, 2) or '--'
            self.highlight_series_rated(ws_row[last_col_index + j], entry_data.get('SeriesRated'))


    def insert_fault_data(self, ws_index: int, ws_row: tuple, values: dict, cap_val: float, offset: int = 0):
        """
        Inserts fault data into specified cells in a row.

        :param int ws_index: Index of the worksheet in the workbook.
        :param tuple[Cell, ...] ws_row: The row to insert data into.
        :param dict values: A dictionary of fault values by configuration.
        :param float cap_val: The capability value to compare against.
        :param int offset: Offset for column insertion, default is 0.
        """
        for config, fault_val in values.items():
            heading_key = CONFIG_MAP.get(config, config)
            col_index = self._get_col_index(ws_index, heading_key)
            if col_index:
                cell = ws_row[col_index + offset]
                cell.value = round(fault_val, 2) or '--'
                self.highlight_high_duty(cell, fault_val, cap_val)