from pathlib import Path
from openpyxl.cell import Cell
from openpyxl.workbook import Workbook
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet
from consts.common import CONFIG_MAP
from consts import styles

HEADER_ROW = 1
SUBHEAD_ROW = 2


class Exporter:

    def __init__(self, sheet_names: list[str], header_names: list[str]):
        """
        Initializes a new instance of the Exporter class,
        creating a new workbook and setting up the worksheets.
        """
        self.ansi_data = {}
        self.iec_data = {}
        self.configs = []
        self.wb = Workbook()
        self._initialize_sheets(sheet_names)
        self.const_cols_header, self.end_cols_header = header_names

    def _initialize_sheets(self, ws_names):
        """
        Initializes and names the worksheets.
        """

        self.wb.active.title = ws_names[0]
        for ws_name in ws_names[1:]:
            self.wb.create_sheet(ws_name)

    def _get_sorted_configs(self) -> list:
        """
        Retrieves and sorts configurations based on ANSI data.

        :return: A list of sorted configurations.
        :rtype: list
        """
        configs = set()
        random_key = next(iter(self.ansi_data))
        for _id, entry in self.ansi_data[random_key].items():
            random_dict = next((v for v in entry.values() if isinstance(v, dict)), {})
            configs.update(random_dict.keys())
        return self.rearrange_list(list(configs), list(CONFIG_MAP.keys()))

    def set_ansi_data(self, ansi_data: dict):
        """
        Sets ANSI data and updates the configurations.

        :param dict ansi_data: A dictionary containing ANSI data.
        """
        self.ansi_data = ansi_data
        self.configs = self._get_sorted_configs()

    def set_iec_data(self, iec_data: dict):
        """
        Sets IEC data and updates the configurations.

        :param dict iec_data: A dictionary containing IEC data.
        """
        self.iec_data = iec_data
        self.configs = self._get_sorted_configs()

    def _get_col_index(self, sheet_index: int, heading: str):
        """
        Gets the column index for a given column heading in a specified sheet.

        :param int sheet_index: Index of the worksheet in the workbook.
        :param str heading: The heading of the column to find.
        :return: The index of the column with the specified heading.
        :rtype: int
        """
        sheet = self.wb.worksheets[sheet_index]
        for col in sheet.iter_cols(1, sheet.max_column):
            if col[0].value and heading in col[0].value:
                return col[0].column - 1

    @staticmethod
    def rearrange_list(target_list: list, reference_list: list) -> list:
        """
        Rearranges a target list based on the order of a reference list.

        :param list target_list: The list to be rearranged.
        :param list reference_list: The list that provides the desired order.
        :return: A new list with elements ordered according to the reference list.
        :rtype: list
        """
        reference_dict = {element: index for index, element in enumerate(reference_list)}
        in_reference = [element for element in target_list if element in reference_dict]
        not_in_reference = [element for element in target_list if element not in reference_dict]
        in_reference_sorted = sorted(in_reference, key=lambda x: reference_dict[x])
        return in_reference_sorted + not_in_reference

    @staticmethod
    def _set_cols_subheads(sheet: Worksheet, columns: list, start_col: int):
        """
        Helper method to set subheads for a worksheet.

        :param Worksheet sheet: The worksheet object.
        :param list columns: List of subhead column names.
        :param int start_col: The starting column index for subhead.
        """
        for i, col_name in enumerate(columns):
            sheet.cell(SUBHEAD_ROW, start_col + i).value = col_name


    def create_headers(self, sheet_index: int, const_cols: list, var_cols: list, col_prefix: str):
        """
        Creates and merges headers for a specified sheet.

        :param int sheet_index: Index of the worksheet in the workbook.
        :param list const_cols: List of constant columns.
        :param list var_cols: List of variable columns.
        :param str col_prefix: Prefix for dynamic column headers.
        """
        var_cols_buff = len(var_cols) + 1
        const_cols_buff = len(const_cols) + 2
        sheet = self.wb.worksheets[sheet_index]
        self._set_const_cols_header(sheet, const_cols)
        self._set_dynamic_cols_headers(sheet, var_cols, col_prefix, const_cols_buff, var_cols_buff)
        if self.end_cols_header:
            self._set_end_cols_headers(sheet, var_cols, const_cols_buff, len(self.configs) * var_cols_buff)


    def _set_const_cols_header(self, sheet: Worksheet, const_cols: list):
        """
        Helper method to set the main header for a sheet.

        :param Worksheet sheet: The worksheet object.
        :param list const_cols: List of constant columns.
        """
        sheet.cell(HEADER_ROW, 1).value = self.const_cols_header
        sheet.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(const_cols))
        self._set_cols_subheads(sheet, const_cols, start_col=1)

    def _set_dynamic_cols_headers(self, sheet: Worksheet, var_cols: list, col_prefix: str,
                                  const_cols_buff: int, var_cols_buff: int):
        """
        Helper method to set dynamic headers for configurations.

        :param Worksheet sheet: The worksheet object.
        :param list var_cols: List of variable columns.
        :param str col_prefix: Prefix for dynamic column headers.
        :param int const_cols_buff: Offset for constant columns.
        :param int var_cols_buff: Offset for variable columns.
        """
        for i, config in enumerate(self.configs):
            col_name = f'{col_prefix}: {CONFIG_MAP.get(config, config)}'
            col_index = const_cols_buff + var_cols_buff * i
            sheet.cell(HEADER_ROW, col_index).value = col_name
            sheet.merge_cells(start_row=1, start_column=col_index, end_row=1,
                              end_column=col_index + len(var_cols) - 1)
            self._set_cols_subheads(sheet, var_cols, start_col=col_index)

    def _set_end_cols_headers(self, sheet: Worksheet, var_cols: list, const_cols_buff: int, offset: int):
        """
        Helper method to set end headers for a sheet.

        :param Worksheet sheet: The worksheet object.
        :param list var_cols: List of variable columns.
        :param int const_cols_buff: Offset for constant columns.
        :param int offset: Offset for the end headers.
        """
        last_col_index = const_cols_buff + offset
        sheet.cell(HEADER_ROW, last_col_index).value = self.end_cols_header
        sheet.merge_cells(start_row=1, start_column=last_col_index, end_row=1,
                          end_column=last_col_index + len(var_cols) - 1)
        self._set_cols_subheads(sheet, var_cols, start_col=last_col_index)

    def format_headers(self, sheet_index: int):
        """
        Applies formatting to the header rows of a specified sheet.

        :param int sheet_index: Index of the worksheet in the workbook.
        """
        sheet = self.wb.worksheets[sheet_index]
        for row in sheet.iter_rows(HEADER_ROW, SUBHEAD_ROW):
            for cell in row:
                self._apply_cell_format(cell, styles.FONT_HEADER, styles.ALIGNMENT,
                                        styles.BORDER_ALL, styles.FILL_HEADER)

    @staticmethod
    def _apply_cell_format(cell: Cell, font, align, border, fill=None):
        """
        Applies formatting to a single cell.

        :param Cell cell: The cell to apply formatting to.
        :param font: Font style for the cell.
        :param align: Alignment for the cell.
        :param border: Border style for the cell.
        :param fill: Fill pattern for the cell.
        """
        cell.font = font
        cell.alignment = align
        cell.border = border
        if fill is not None:
            cell.fill = fill

    def format_sheet(self, sheet_index: int, const_cols_len: int, var_cols_len: int, col_width: float):
        """
        Applies formatting to the entire sheet.

        :param int sheet_index: Index of the worksheet in the workbook.
        :param int const_cols_len: Number of constant columns.
        :param int var_cols_len: Number of variable columns.
        :param float col_width: Width of the columns.
        """
        sheet = self.wb.worksheets[sheet_index]
        self._apply_row_format(sheet, SUBHEAD_ROW + 1)
        self._set_column_widths(sheet, const_cols_len, var_cols_len, col_width)
        self._format_numbers(sheet)

    def _apply_row_format(self, sheet: Worksheet, start_row):
        """
        Helper method to apply formatting to data rows.

        :param Worksheet sheet: The worksheet object.
        """
        for i in range(start_row, sheet.max_row + 1):
            for cell in sheet[i]:
                fill = self._get_row_cell_fill(cell)
                self._apply_cell_format(cell, styles.FONT_ENTRIES, styles.ALIGNMENT, styles.BORDER_ALL, fill)

    @staticmethod
    def _get_row_cell_fill(cell: Cell):
        """
        Applies fill to a single row cell.

        :param Cell cell: The cell to apply fill to.
        """
        if cell.row % 2 == 0 and not cell.fill.patternType:
            return styles.FILL_ROW_BLUE

    @staticmethod
    def _set_column_widths(sheet: Worksheet, const_cols_len: int, var_cols_len: int, col_width: float):
        """
        Helper method to set column widths.

        :param Worksheet sheet: The worksheet object.
        :param int const_cols_len: Number of constant columns.
        :param int var_cols_len: Number of variable columns.
        :param float col_width: Width of the columns.
        """
        for i in range(1, sheet.max_column + 1):
            if i % (var_cols_len + 1) == (const_cols_len + 1) % (var_cols_len + 1) and i > const_cols_len:
                sheet.column_dimensions[get_column_letter(i)].width = styles.WIDTH_BUFFER
                for column in sheet.iter_cols(i, i):
                    for cell in column:
                        cell.fill = styles.FILL_ROW_BLANK
                        cell.border = styles.BORDER_VERTICAL
            else:
                sheet.column_dimensions[get_column_letter(i)].width = col_width

    @staticmethod
    def _format_numbers(sheet: Worksheet):
        """
        Formats cells with numerical values.

        :param Worksheet sheet: The worksheet object.
        """
        for i in range(1, sheet.max_column + 1):
            for column in sheet.iter_cols(i, i):
                for cell in column:
                    cell.number_format = styles.NUMBER_FORMAT

    def save_workbook(self, wb_path: Path):
        """
        Saves the workbook to a specified filename.

        :param Path wb_path: The filename to save the workbook as.
        """
        self.wb.save(wb_path)
        self.wb.close()