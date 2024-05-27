import fnmatch
import os
import win32com.client as win32


def convert_doc_to_docx(doc_path):
    """Convert .doc to .docx and delete the original .doc file"""
    word = win32.Dispatch("Word.Application")
    doc = word.Documents.Open(doc_path)
    docx_path = os.path.splitext(doc_path)[0] + ".docx"
    doc.SaveAs(docx_path, FileFormat=16)  # 16 is the format code for docx
    doc.Close()
    word.Quit()
    os.remove(doc_path)
    print(f"\ndoc转为docx {doc_path} \nto {docx_path} ")


def convert_docx_to_doc(docx_path):
    """Convert .docx to .doc and delete the original .docx file"""
    word = win32.Dispatch("Word.Application")
    doc = word.Documents.Open(docx_path)
    doc_path = os.path.splitext(docx_path)[0] + ".doc"
    doc.SaveAs(doc_path, FileFormat=0)  # 0 is the format code for doc
    doc.Close()
    word.Quit()
    os.remove(docx_path)
    print(f"\ndocx转为doc {docx_path} \nto {doc_path} ")


def convert_file_format(file_path):
    """Convert file format between .doc and .docx and delete the original file"""
    if file_path.lower().endswith(".doc"):
        convert_doc_to_docx(file_path)
    elif file_path.lower().endswith(".docx"):
        convert_docx_to_doc(file_path)
    else:
        print("\n该文件既不是 .doc 格式也不是 .docx 格式")


def find_targe_file(directory, target):
    for root, dirs, files in os.walk(directory):
        for file in files:
            if fnmatch.fnmatch(file, f'*{target}*'):
                fpa_file = os.path.join(root, file)
                print("\n找到" + f'{target}' + "文件: "f'{fpa_file}')
                return fpa_file
    print(f"\n未找到包含 {target} 的文件。")
    return None
