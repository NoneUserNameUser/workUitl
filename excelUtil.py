import openpyxl


def extract_data(file_path, sheet_name) -> dict:
    """

    :rtype: object
    """
    data_map = {}  # 使用字典存储数据

    wb = openpyxl.load_workbook(file_path)
    sheet = wb[sheet_name]

    # 从第 3 行开始遍历工作表
    for row in sheet.iter_rows(min_row=3, values_only=True):
        # 获取 D 列和 G 列的值
        d_value = row[3]
        g_value = row[6]

        # 将 D 列的值作为键，G 列的值作为值，添加到字典中
        data_map[d_value] = g_value

    return data_map




