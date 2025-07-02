from pypdf import PdfReader, PdfWriter
import os

def rotate_pdf_pages_interactive():
    """
    引導使用者輸入 PDF 檔名、旋轉角度和要旋轉的頁面，
    並將旋轉後的 PDF 儲存回 D:\PDF_Rotate 資料夾。
    """
    base_dir = r"D:\PDF_Rotate"

    # 確保目標資料夾存在
    if not os.path.exists(base_dir):
        print(f"警告：資料夾 '{base_dir}' 不存在，將自動建立。")
        os.makedirs(base_dir)

    # 讓使用者輸入 PDF 檔名 (不含副檔名)
    pdf_filename_base = input("請輸入 PDF 檔案名稱（不包含副檔名）：")
    input_pdf_path = os.path.join(base_dir, f"{pdf_filename_base}.pdf")
    
    # 輸出檔案名稱：在原檔名後加上 _Rotated
    output_pdf_path = os.path.join(base_dir, f"{pdf_filename_base}_Rotated.pdf")

    # 檢查輸入 PDF 檔案是否存在
    if not os.path.exists(input_pdf_path):
        print(f"錯誤：找不到檔案 '{input_pdf_path}'。請確認檔案是否存在於 '{base_dir}' 資料夾中。")
        return

    # 讓使用者輸入旋轉角度
    rotation_angle = None
    while rotation_angle not in [90, 180, 270]:
        try:
            angle_input = int(input("請輸入旋轉角度 (90, 180, 270)："))
            if angle_input in [90, 180, 270]:
                rotation_angle = angle_input
            else:
                print("無效的旋轉角度，請輸入 90, 180 或 270。")
        except ValueError:
            print("無效的輸入，請輸入數字。")

    # 讓使用者輸入要旋轉的頁面 (可選)
    pages_to_rotate = None
    pages_input = input("請輸入要旋轉的頁面編號（用逗號分隔，從 1 開始，留空則旋轉所有頁面）：")
    if pages_input.strip(): # 檢查輸入是否為空字串或只包含空格
        try:
            # 將使用者輸入的頁碼 (從 1 開始) 轉換為 Python 列表索引 (從 0 開始)
            pages_to_rotate = [int(page) - 1 for page in pages_input.split(',')]
            # 濾除無效的頁碼（小於0的）
            pages_to_rotate = [p for p in pages_to_rotate if p >= 0]
        except ValueError:
            print("頁面編號輸入無效，將旋轉所有頁面。")
            pages_to_rotate = None
    
    print(f"\n正在旋轉 '{input_pdf_path}' 的頁面...")

    try:
        reader = PdfReader(input_pdf_path)
        writer = PdfWriter()

        num_pages = len(reader.pages)

        # 遍歷所有頁面
        for i in range(num_pages):
            page = reader.pages[i]
            
            # 判斷是否需要旋轉當前頁面
            # 如果 pages_to_rotate 為 None (旋轉所有頁面)
            # 或者當前頁面的索引 i 在 pages_to_rotate 列表中 (旋轉特定頁面)
            if pages_to_rotate is None or i in pages_to_rotate:
                page.rotate(rotation_angle)
                print(f"旋轉頁面 {i + 1} {rotation_angle} 度。")
            
            writer.add_page(page) # 將處理後的頁面加入到輸出中

        with open(output_pdf_path, "wb") as f:
            writer.write(f)
        print(f"\n旋轉完成！新檔案已儲存至：'{output_pdf_path}'")

    except Exception as e:
        print(f"處理檔案時發生錯誤：{e}")

# 執行旋轉功能
if __name__ == "__main__":
    rotate_pdf_pages_interactive()