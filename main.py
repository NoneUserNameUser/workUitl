import os
from excelUtil import extract_data
from fileUtil import convert_file_format, find_targe_file
from wordUtil import process_document

while True:
    folder_name = input("请输入文件夹名称(按 Enter 键退出): ")
    if not folder_name:
        break
    parent_directory = os.path.dirname(os.getcwd())
    file_path = os.path.join(parent_directory, folder_name)
    print("\n文件夹实际路径为:" f'{file_path}')

    #  Excel 和 Word 文件路径
    fpa_file = find_targe_file(file_path, 'FPA')
    mr_file = find_targe_file(file_path, 'MR')
    if fpa_file is None or mr_file is None:
        continue
    sheet_name = 'FPA功能点估算'

    # 将 .doc 文件转换为 .docx 文件
    convert_file_format(mr_file)
    word_path = mr_file.replace('.doc', '.docx')

    print("\n读取FPA")
    # 读取Excel
    excel_data = extract_data(fpa_file, sheet_name)

    print("\n写入MR")
    # 写入Word
    process_document(word_path, excel_data)

    # 将 .docx 文件转换回 .doc 文件
    convert_file_format(word_path)
    print("\n写入完成\n")
