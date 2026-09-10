import win32com.client
import os

def export_pptx_to_png(pptx_path, output_dir, width=1920, height=1080):
    # 將相對路徑轉換為絕對路徑
    pptx_path = os.path.abspath(pptx_path)
    
    # 檢查 PPTX 檔案是否存在
    if not os.path.exists(pptx_path):
        print(f"錯誤：找不到簡報檔案 -> {pptx_path}")
        print("請確認檔名是否正確，或檔案是否放在與 Python 腳本相同的資料夾中。")
        return

    output_dir = os.path.abspath(output_dir)

    # 確保輸出資料夾存在
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    print("啟動 PowerPoint 中...")
    powerpoint = win32com.client.Dispatch("PowerPoint.Application")
    
    try:
        # 開啟簡報檔 (唯讀, 無視窗)
        presentation = powerpoint.Presentations.Open(pptx_path, WithWindow=False)
        
        # 走訪每一頁投影片並匯出
        for i, slide in enumerate(presentation.Slides):
            output_path = os.path.join(output_dir, f"slide_{i+1:03d}.png")
            slide.Export(output_path, "PNG", width, height)
            print(f"已匯出: {output_path}")
            
        presentation.Close()
        print("轉檔完成！")
        
    except Exception as e:
        print(f"PowerPoint 處理時發生錯誤: {e}")
    finally:
        # 確保關閉 PowerPoint 處理程序
        powerpoint.Quit()

# ==========================================
# 請將下方的 "您的實際檔名.pptx" 換成真正的檔名
# ==========================================
target_file = "Home_Care_Navigator.pptx"  
export_pptx_to_png(target_file, "output_images")

# import win32com.client
# import os
# # pip install pywin32
# def export_pptx_to_png(pptx_path, output_dir, width=1920, height=1080):
#     # COM 介面需要絕對路徑
#     pptx_path = os.path.abspath(pptx_path)
#     output_dir = os.path.abspath(output_dir)

#     # 確保輸出資料夾存在
#     if not os.path.exists(output_dir):
#         os.makedirs(output_dir)

#     # 啟動 PowerPoint 應用程式 (背景執行)
#     powerpoint = win32com.client.Dispatch("PowerPoint.Application")
    
#     try:
#         # 開啟簡報檔 (唯讀, 無視窗)
#         presentation = powerpoint.Presentations.Open(pptx_path, WithWindow=False)
        
#         # 走訪每一頁投影片並匯出
#         for i, slide in enumerate(presentation.Slides):
#             output_path = os.path.join(output_dir, f"slide_{i+1:03d}.png")
#             # 匯出語法: Export(FileName, FilterName, ScaleWidth, ScaleHeight)
#             slide.Export(output_path, "PNG", width, height)
#             print(f"已匯出: {output_path}")
            
#         presentation.Close()
#     finally:
#         # 確保關閉 PowerPoint 處理程序
#         powerpoint.Quit()

# # 執行範例
# export_pptx_to_png("sample.pptx", "output_images")