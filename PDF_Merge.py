# 把要合併的檔案放在D:\PDF_Merge，會依檔名排序合併，輸出檔案在同一個資料夾
import os  # 匯入 os 模組以處理檔案和路徑
import sys  # 匯入 sys 模組以處理系統相關功能
import subprocess  # 匯入 subprocess 模組以執行外部命令

# 檢查並安裝 pypdf
try:
    from pypdf import PdfWriter  # 嘗試匯入 pypdf 的 PdfWriter
except ImportError:  # 如果匯入失敗（即 pypdf 未安裝）
    print("pypdf 未安裝，正在嘗試自動安裝...")
    try:
        subprocess.check_call([sys.executable, "-m", "pip", "install", "pypdf"])  # 使用 subprocess 執行 pip install pypdf
        from pypdf import PdfWriter  # 安裝成功後再次匯入
        print("pypdf 已成功安裝！")
    except Exception as e:  # 如果安裝失敗
        print(f"安裝 pypdf 失敗: {str(e)}")
        print("請手動執行 'pip install pypdf' 或檢查網路與權限")
        sys.exit(1)  # 退出程式

def merge_pdfs_in_folder(folder_path, output_filename="merged_output.pdf"):
    # 定義函數，接受資料夾路徑和可選的輸出檔名，預設為 "merged_output.pdf"
    try:
        merger = PdfWriter()  # 建立 PdfWriter 物件用於合併 PDF
        pdf_files = [f for f in os.listdir(folder_path) if f.lower().endswith('.pdf')]  # 獲取資料夾中所有 .pdf 檔案
        
        if not pdf_files:  # 檢查是否有 PDF 檔案
            print("錯誤：資料夾中沒有找到 PDF 檔案")  # 若無 PDF，輸出錯誤訊息
            return  # 結束函數執行
        
        pdf_files.sort()  # 按檔名排序，確保合併順序一致（可選）
        
        for pdf in pdf_files:  # 遍歷每個 PDF 檔案
            pdf_path = os.path.join(folder_path, pdf)  # 組合完整檔案路徑
            try:
                merger.append(pdf_path)  # 將 PDF 加入合併器
                print(f"已加入 {pdf}")  # 輸出成功訊息
            except Exception as e:  # 捕獲單個檔案處理時的錯誤
                print(f"處理 {pdf} 時發生錯誤: {str(e)}")  # 輸出錯誤訊息
        
        output_path = os.path.join(folder_path, output_filename)  # 組合輸出檔案的完整路徑
        merger.write(output_path)  # 將合併結果寫入輸出檔案
        merger.close()  # 關閉 merger 物件，釋放資源
        print(f"PDF 已成功合併至 {output_path}")  # 輸出完成訊息
        
    except Exception as e:  # 捕獲整個合併過程中的其他錯誤
        print(f"合併過程中發生錯誤: {str(e)}")  # 輸出錯誤訊息

if __name__ == "__main__":
    # 主程式區塊，只有直接執行此腳本時才運行
    folder_path = r"D:\PDF_Merge"  # 指定 PDF 檔案所在的資料夾
    merge_pdfs_in_folder(folder_path)  # 呼叫函數執行合併
