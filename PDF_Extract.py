import PyPDF2
import os

def extract_pages(input_pdf, output_pdf, pages):
    """
    從 PDF 檔案中擷取特定頁面。

    Args:
        input_pdf (str): 輸入 PDF 檔案的路徑。
        output_pdf (str): 輸出 PDF 檔案的路徑。
        pages (list): 要擷取的頁面編號列表（從 0 開始）。
    """
    try:
        with open(input_pdf, 'rb') as file:
            pdf_reader = PyPDF2.PdfReader(file)
            pdf_writer = PyPDF2.PdfWriter()

            for page_num in pages:
                if 0 <= page_num < len(pdf_reader.pages):
                    page = pdf_reader.pages[page_num]
                    pdf_writer.add_page(page)

            with open(output_pdf, 'wb') as output_file:
                pdf_writer.write(output_file)
        print(f"成功擷取頁面並儲存至 {output_pdf}")

    except FileNotFoundError:
        print(f"錯誤：找不到檔案 {input_pdf}")
    except Exception as e:
        print(f"發生錯誤：{e}")

# 取得使用者輸入
input_filename = input("請輸入 PDF 檔案名稱（不包含副檔名）：")
input_pdf = os.path.join("D:\\PDF_Extract", f"{input_filename}.pdf")

# 取得原檔案名稱，用於輸出檔案命名
output_pdf = os.path.join("D:\\PDF_Extract", f"{input_filename}_Extract.pdf")

# 取得使用者輸入頁數
pages_input = input("請輸入要擷取的頁面編號（用逗號分隔，從 1 開始）：")
pages_to_extract = [int(page) - 1 for page in pages_input.split(',')] # 將頁碼轉換為從 0 開始的索引

# 呼叫函式
extract_pages(input_pdf, output_pdf, pages_to_extract)