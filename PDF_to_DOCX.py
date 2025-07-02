import os
from pdf2docx import Converter

def convert_pdf_to_docx_interactive():
    """
    引導使用者輸入 PDF 檔名，將其轉換為 Word (.docx) 檔案。
    輸入與輸出檔案都位於 D:\PDF_to_DOC 資料夾。
    """
    # 將字串改為原始字串 (r"...") 以避免反斜杠警告
    base_dir = r"D:\PDF_to_DOC" 

    # 確保目標資料夾存在
    if not os.path.exists(base_dir):
        print(f"警告：資料夾 '{base_dir}' 不存在，將自動建立。")
        os.makedirs(base_dir)

    # 讓使用者輸入 PDF 檔名 (不含副檔名)
    pdf_filename_base = input("請輸入 PDF 檔案名稱（不包含副檔名）：")

    input_pdf_path = os.path.join(base_dir, f"{pdf_filename_base}.pdf")
    output_docx_path = os.path.join(base_dir, f"{pdf_filename_base}.docx")

    # 檢查輸入 PDF 檔案是否存在
    if not os.path.exists(input_pdf_path):
        print(f"錯誤：找不到檔案 '{input_pdf_path}'。請確認檔案是否存在於 '{base_dir}' 資料夾中。")
        return

    print(f"\n正在將 '{input_pdf_path}' 轉換為 Word 檔案...")
    
    try:
        cv = Converter(input_pdf_path)
        cv.convert(output_docx_path, start=0, end=None)
        cv.close()
        print(f"成功轉換！Word 檔案已儲存至：'{output_docx_path}'")
        
        # 顯示轉換前後檔案大小
        original_size_mb = os.path.getsize(input_pdf_path) / (1024 * 1024)
        converted_size_mb = os.path.getsize(output_docx_path) / (1024 * 1024)
        print(f"原始 PDF 檔案大小: {original_size_mb:.2f} MB")
        print(f"轉換後的 Word 檔案大小: {converted_size_mb:.2f} MB")

    except Exception as e:
        print(f"轉換檔案時發生錯誤：{e}")

# 執行轉換功能
if __name__ == "__main__":
    convert_pdf_to_docx_interactive()
