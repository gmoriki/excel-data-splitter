import openpyxl
from openpyxl.styles import Border, Side
from openpyxl.utils.dataframe import dataframe_to_rows

import pandas as pd


def apply_border_to_adjacent_cells(ws):
    """
    Applies border to cells that are adjacent to non-empty cells.

    Args:
        ws (Worksheet): The openpyxl Worksheet object to process.
    """

    max_row = ws.max_row
    max_col = ws.max_column

    def is_adjacent_cell_empty(cell):
        row, col = cell.row, cell.column
        left = ws.cell(row=row, column=col - 1) if col > 1 else None
        right = ws.cell(row=row, column=col + 1) if col < max_col else None
        top = ws.cell(row=row - 1, column=col) if row > 1 else None
        bottom = ws.cell(row=row + 1, column=col) if row < max_row else None

        return not (left and left.value or right and right.value or top and top.value or bottom and bottom.value)

    thin_border = Border(
        left=Side(style="thin"), right=Side(style="thin"), top=Side(style="thin"), bottom=Side(style="thin")
    )

    for row in ws.iter_rows(min_row=1, max_col=max_col, max_row=max_row):
        for cell in row:
            if cell.value and not is_adjacent_cell_empty(cell):
                cell.border = thin_border


def apply_border_to_range(ws, min_row, max_row, min_col, max_col):
    """
    Applies border to a range of cells within the specified row and column boundaries.

    Args:
        ws (Worksheet): The openpyxl Worksheet object to process.
        min_row (int): The minimum row index.
        max_row (int): The maximum row index.
        min_col (int): The minimum column index.
        max_col (int): The maximum column index.
    """

    thin_border = Border(
        left=Side(style="thin"), right=Side(style="thin"), top=Side(style="thin"), bottom=Side(style="thin")
    )

    for row in range(min_row, max_row + 1):
        for col in range(min_col, max_col + 1):
            ws.cell(row=row, column=col).border = thin_border


def unmerge_and_fill_cells(ws):
    """
    Unmerges merged cells and fills unmerged cells with the same value as the original cell.

    Args:
        ws (Worksheet): The openpyxl Worksheet object to process.
    """

    merged_ranges = ws.merged_cells.ranges.copy()
    for merged_cells in merged_ranges:
        top_left_cell = ws.cell(row=merged_cells.min_row, column=merged_cells.min_col)
        top_left_value = top_left_cell.value
        ws.unmerge_cells(str(merged_cells))
        apply_border_to_range(
            ws, merged_cells.min_row, merged_cells.max_row, merged_cells.min_col, merged_cells.max_col
        )
        for row in range(merged_cells.min_row, merged_cells.max_row + 1):
            for col in range(merged_cells.min_col, merged_cells.max_col + 1):
                cell = ws.cell(row=row, column=col)
                if cell.value is None:
                    cell.value = top_left_value

    # プレーンテキストに変換
    for i, row in enumerate(ws.iter_rows(), start=1):
        for j, cell in enumerate(row, start=1):
            ws.cell(row=i, column=j).value = cell.value


def find_bordered_tables(ws):
    """
    Finds tables that are enclosed by borders in the worksheet.

    Args:
        ws (Worksheet): The openpyxl Worksheet object to process.

    Returns:
        list: A list of tuples containing the start and end cell coordinates of the bordered tables.
    """

    tables = []
    max_row = ws.max_row
    max_col = ws.max_column
    visited_cells = set()

    for r in range(1, max_row + 1):
        for c in range(1, max_col + 1):
            if (r, c) not in visited_cells:
                cell = ws.cell(row=r, column=c)
                if cell.border.left.style and cell.border.top.style:
                    start = cell.coordinate
                    r_end, c_end = r, c

                    while ws.cell(row=r_end + 1, column=c).border.top.style:
                        r_end += 1
                    while ws.cell(row=r, column=c_end + 1).border.left.style:
                        c_end += 1

                    end = ws.cell(row=r_end, column=c_end).coordinate
                    tables.append((start, end))

                    for row in range(r, r_end + 1):
                        for col in range(c, c_end + 1):
                            visited_cells.add((row, col))
    return tables


def read_table_as_dataframe(ws, start, end):
    """
    Reads a table within the specified cell range as a pandas DataFrame.

    Args:
        ws (Worksheet): The openpyxl Worksheet object to process.
        start (str): The start cell coordinate of the table.
        end (str): The end cell coordinate of the table.

    Returns:
        DataFrame: A pandas DataFrame object containing the table data.
    """

    data = []
    for row in ws[start:end]:
        data.append([cell.value for cell in row])
    return pd.DataFrame(data[1:], columns=data[0])


def save_dataframe_to_excel_sheet(df, workbook, sheet_name):
    """
    Saves a pandas DataFrame to a new sheet in the specified openpyxl Workbook.

    Args:
        df (DataFrame): The pandas DataFrame to save.
        workbook (Workbook): The openpyxl Workbook to save the DataFrame to.
        sheet_name (str): The name of the new sheet to create.
    """

    ws = workbook.create_sheet(sheet_name)

    for r_idx, row in enumerate(dataframe_to_rows(df, index=False, header=True), start=1):
        for c_idx, value in enumerate(row, start=1):
            ws.cell(row=r_idx, column=c_idx, value=value)


def unmerge_and_delete_unbordered_cells(ws):
    """
    Unmerges cells outside the borders and deletes cells that are not bordered.

    Args:
        ws (Worksheet): The openpyxl Worksheet object to process.
    """

    # 枠線で囲まれたセルの範囲を特定する
    bordered_cells = []

    for row in ws:
        for cell in row:
            if any([cell.border.left.style, cell.border.right.style, cell.border.top.style, cell.border.bottom.style]):
                bordered_cells.append(cell.coordinate)

    # 枠線で囲まれていないセルの結合を解除する
    for merge_range in ws.merged_cells.ranges:
        is_outside_borders = False
        for cell in merge_range:
            if cell.coordinate not in bordered_cells:
                is_outside_borders = True
                break

        if is_outside_borders:
            ws.unmerge_cells(str(merge_range))

    # 枠線で囲まれていないセルを削除する
    for row in ws.iter_rows():
        for cell in row:
            if cell.coordinate not in bordered_cells:
                ws[cell.coordinate].value = None


def main(input_excel_file, output_excel_file):
    """
    Processes an input Excel file and saves the processed data to a new output Excel file.

    Args:
        input_excel_file (str): The path of the input Excel file.
        output_excel_file (str): The path of the output Excel file.
    """
    # input_excel_file = "input.xlsx"
    # Excelファイルを開く
    wb = openpyxl.load_workbook(input_excel_file, data_only=True)

    # output_excel_file = "output.xlsx"
    output_wb = openpyxl.Workbook()
    output_wb.remove(output_wb.active)  # デフォルトの空シートを削除

    for ws in wb.worksheets:
        # 枠線で囲まれていないセルを全て削除
        unmerge_and_delete_unbordered_cells(ws)

        # 隣接するセルのみテーブルと定めて、枠線を引く処理
        apply_border_to_adjacent_cells(ws)

        # 結合セルを解除し、元のセルと同じ値を代入
        unmerge_and_fill_cells(ws)

        # シート内の枠線で括られたテーブルを検出
        tables = find_bordered_tables(ws)

        # テーブルを読み込み
        dataframes = [read_table_as_dataframe(ws, start, end) for start, end in tables]

        # データフレームを新しいExcelファイルの複数シートに出力
        for idx, df in enumerate(dataframes):
            sheet_name = f"{ws.title}_Table_{idx+1}"
            save_dataframe_to_excel_sheet(df, output_wb, sheet_name)

    output_wb.save(output_excel_file)


if __name__ == "__main__":
    main("./input.xlsx", "./output.xlsx")
