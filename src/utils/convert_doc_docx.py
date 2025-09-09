# 将.doc文件转换为.docx文件，并替换源文件
import os
import logging
import comtypes.client

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# --- DOC 转 DOCX 并替换源文件 ---
def convert_doc_to_docx_and_replace(doc_path):
    """
    将 .doc 文件转换为 .docx 格式，并替换原始 .doc 文件。
    此功能仅在 Windows 操作系统上有效，且需要安装 Microsoft Word。

    Args:
        doc_path (str): .doc 文件的完整路径。

    Returns:
        str: 转换后的 .docx 文件路径，如果转换失败则返回 None。
    """
    doc_path = os.path.abspath(doc_path)
    if not os.path.exists(doc_path):
        logging.error(f"错误：文件不存在 - {doc_path}")
        return None

    if not doc_path.lower().endswith('.doc'):
        logging.error(f"错误：文件不是.doc 格式 - {doc_path}")
        return None

    # 构建新的 .docx 文件路径
    docx_path = doc_path + 'x' # 简单地在后缀名后面加 'x'

    try:
        # 启动 Word 应用程序
        word = comtypes.client.CreateObject("Word.Application") # 使用 comtypes.client 接口
        word.Visible = False # 不显示 Word 界面

        # 打开 .doc 文档
        doc = word.Documents.Open(doc_path)

        # 保存为 .docx 格式
        # wdFormatDocumentDefault = 16，是 Word 2007 及以上版本的默认文档格式 (.docx)
        doc.SaveAs(docx_path, FileFormat=16)

        # 关闭文档
        doc.Close()
        # 退出 Word 应用程序
        word.Quit()

        # 删除原始 .doc 文件
        os.remove(doc_path)
        logging.info(f"成功将 '{doc_path}' 转换为 '{docx_path}' 并删除原始文件")
        return docx_path

    except Exception as e:
        logging.error(f"转换 '{doc_path}' 为 .docx 时发生错误：{e}")
        # 尝试关闭 Word 应用程序以避免残留进程
        try:
            if 'word' in locals() and word:
                word.Quit()
        except:
            pass
        return None
    
if __name__ == "__main__":
    test_doc_path = r"input_files\2023到访工厂打印资料\6月\上海塑柯新材料有限公司-2023.06.06.doc"
    fil=convert_doc_to_docx_and_replace(test_doc_path)
    print(fil)