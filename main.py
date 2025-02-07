import tkinter as tk
from tkinter import ttk, scrolledtext
from PIL import Image, ImageTk, ImageGrab
import pytesseract
import ctypes

# 設定 DPI Awareness（僅限 Windows），以避免因系統縮放比例造成的座標偏移問題
try:
    ctypes.windll.user32.SetProcessDPIAware()
except Exception as e:
    print("設定 DPI Awareness 失敗:", e)


# 若 tesseract 未加入系統環境變數，請設定正確路徑，例如：
pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

class ScreenCaptureOCR:
    def __init__(self, master):
        self.master = master
        self.master.title("螢幕截圖 OCR 工具")
        self.master.geometry("800x600")

        # 按鈕：開始截圖
        self.capture_button = ttk.Button(master, text="開始截圖", command=self.start_capture)
        self.capture_button.pack(pady=10)

        # 用來顯示截圖圖片的 Label
        self.image_label = tk.Label(master)
        self.image_label.pack(pady=10)

        # 顯示 OCR 結果的捲軸文字框
        self.result_text = scrolledtext.ScrolledText(master, wrap=tk.WORD, width=90, height=20)
        self.result_text.pack(padx=10, pady=10)

    def start_capture(self):
        # 隱藏主視窗，避免 UI 進入截圖
        self.master.withdraw()

        # 建立全螢幕、無邊框、半透明的 Toplevel 視窗作為選取介面
        self.select_win = tk.Toplevel(self.master)
        self.select_win.attributes("-fullscreen", True)
        self.select_win.attributes("-alpha", 0.3)  # 半透明效果
        self.select_win.overrideredirect(True)       # 移除邊框

        # 建立填滿 Toplevel 的 Canvas，並綁定滑鼠事件（用來畫選取框）
        self.canvas = tk.Canvas(self.select_win, cursor="cross", bg="gray")
        self.canvas.pack(fill=tk.BOTH, expand=True)
        self.canvas.bind("<ButtonPress-1>", self.on_button_press)
        self.canvas.bind("<B1-Motion>", self.on_mouse_drag)
        self.canvas.bind("<ButtonRelease-1>", self.on_button_release)

    def on_button_press(self, event):
        # 記錄滑鼠按下時的全域（螢幕）座標
        self.start_x_global = event.x_root
        self.start_y_global = event.y_root

        # 同時轉換成 Canvas 座標用於畫圖（offset 為 Toplevel 的螢幕偏移）
        offset_x = self.select_win.winfo_rootx()
        offset_y = self.select_win.winfo_rooty()
        self.start_x_canvas = event.x_root - offset_x
        self.start_y_canvas = event.y_root - offset_y

        # 在 Canvas 上建立選取矩形
        self.rect = self.canvas.create_rectangle(
            self.start_x_canvas, self.start_y_canvas,
            self.start_x_canvas, self.start_y_canvas,
            outline="red", width=2
        )

    def on_mouse_drag(self, event):
        offset_x = self.select_win.winfo_rootx()
        offset_y = self.select_win.winfo_rooty()
        cur_x_canvas = event.x_root - offset_x
        cur_y_canvas = event.y_root - offset_y
        # 更新矩形位置
        self.canvas.coords(self.rect, self.start_x_canvas, self.start_y_canvas, cur_x_canvas, cur_y_canvas)

    def on_button_release(self, event):
        # 取得滑鼠放開時的全域座標
        end_x_global = event.x_root
        end_y_global = event.y_root

        # 計算截圖區域的左上角與右下角
        x1 = min(self.start_x_global, end_x_global)
        y1 = min(self.start_y_global, end_y_global)
        x2 = max(self.start_x_global, end_x_global)
        y2 = max(self.start_y_global, end_y_global)

        # 關閉選取視窗
        self.select_win.destroy()

        # 延遲一小段時間再截圖，確保工具 UI 已完全隱藏
        self.master.after(200, lambda: self.process_capture(x1, y1, x2, y2))

    def process_capture(self, x1, y1, x2, y2):
        try:
            # 使用 Pillow 擷取指定區域（bbox 格式：(左, 上, 右, 下)）
            img = ImageGrab.grab(bbox=(x1, y1, x2, y2))
            
            # 將截圖轉成 Tkinter 可用的 PhotoImage 並顯示到畫面上
            self.photo = ImageTk.PhotoImage(img)
            self.image_label.configure(image=self.photo)
            
            # 利用 pytesseract 進行 OCR 辨識（預設英文；若需中文請改為 lang="chi_sim" 並安裝中文語言包）
            text = pytesseract.image_to_string(img, lang="eng")
            self.result_text.delete(1.0, tk.END)
            self.result_text.insert(tk.END, text)
        except Exception as e:
            self.result_text.delete(1.0, tk.END)
            self.result_text.insert(tk.END, "發生錯誤：\n" + str(e))
        finally:
            # 截圖完成後再顯示主視窗
            self.master.deiconify()

if __name__ == "__main__":
    root = tk.Tk()
    app = ScreenCaptureOCR(root)
    root.mainloop()
