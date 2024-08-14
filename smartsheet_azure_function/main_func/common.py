import pandas as pd

###
#  シート情報をdataframeに変換
#  @param sheetオブジェクト
#  @return data_frame変換情報
###
def simple_sheet_to_dataframe(sheet):
    col_names = [col.title for col in sheet.columns]
    rows = []
    for row in sheet.rows:
        cells = []
        for cell in row.cells:
            cells.append(cell.value)
        rows.append(cells)
    data_frame = pd.DataFrame(rows, columns=col_names)
    return data_frame


