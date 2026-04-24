import os
import tkinter as tk
from tkinter import filedialog, colorchooser, messagebox
from datetime import datetime
from PIL import Image


# ==========================================
# 核心轉換模組
# ==========================================
def convert_images_to_pdf(image_paths, output_pdf_path, bg_color=(255, 255, 255)):
    if not image_paths:
        return False, "沒有提供圖片路徑。"

    image_list = []
    first_image = None

    try:
        for path in image_paths:
            if not os.path.exists(path):
                continue

            img = Image.open(path)

            # 處理透明通道
            if img.mode in ('RGBA', 'LA') or (img.mode == 'P' and 'transparency' in img.info):
                img_rgba = img.convert('RGBA')
                background = Image.new('RGB', img_rgba.size, bg_color)
                background.paste(img_rgba, mask=img_rgba.split()[3])
                img = background
            else:
                img = img.convert('RGB')

            if first_image is None:
                first_image = img
            else:
                image_list.append(img)

        if first_image is None:
            return False, "找不到有效的圖片可以轉換。"

        first_image.save(
            output_pdf_path,
            "PDF",
            resolution=100.0,
            save_all=True,
            append_images=image_list
        )
        return True, f"成功儲存至：\n{output_pdf_path}"

    except Exception as e:
        return False, f"轉換發生錯誤：{e}"


# ==========================================
# 圖形化介面 (GUI) 模組
# ==========================================
class ImageToPdfApp:
    def __init__(self, root):
        self.root = root
        self.root.title("圖片轉 PDF 工具")
        # 稍微加高視窗以容納新的輸入框
        self.root.geometry("520x500")
        self.root.resizable(False, False)

        # 變數初始化
        self.image_paths = []
        self.output_dir = os.getcwd()
        self.bg_color = (255, 255, 255)
        self.merge_mode = tk.BooleanVar(value=True)

        # 預設檔名使用當下時間 (格式：YYYYMMDD_HHMMSS)
        default_filename = datetime.now().strftime("%Y%m%d_%H%M%S")
        self.custom_filename = tk.StringVar(value=default_filename)

        self.create_widgets()

    def create_widgets(self):
        # --- 圖片選擇區 ---
        frame_top = tk.Frame(self.root, pady=10)
        frame_top.pack(fill=tk.X, padx=20)

        self.btn_select_imgs = tk.Button(frame_top, text="1. 選擇圖片 (可多選)", command=self.select_images, width=20)
        self.btn_select_imgs.pack(anchor=tk.W)

        self.lbl_img_count = tk.Label(frame_top, text="目前已選擇: 0 張圖片", fg="blue")
        self.lbl_img_count.pack(anchor=tk.W, pady=5)

        # --- 輸出選項區 ---
        frame_mid = tk.LabelFrame(self.root, text="輸出設定", pady=10, padx=10)
        frame_mid.pack(fill=tk.X, padx=20, pady=10)

        # 1. 合併或分開 (加入 command 連動輸入框狀態)
        tk.Radiobutton(frame_mid, text="全部合併成單一 PDF", variable=self.merge_mode, value=True,
                       command=self.toggle_filename_entry).grid(row=0, column=0, sticky=tk.W)
        tk.Radiobutton(frame_mid, text="每張獨立轉成一個 PDF", variable=self.merge_mode, value=False,
                       command=self.toggle_filename_entry).grid(row=0, column=1, sticky=tk.W)

        # 2. 自訂檔名 (只在合併模式下啟用)
        frame_filename = tk.Frame(frame_mid)
        frame_filename.grid(row=1, column=0, columnspan=2, sticky=tk.W, pady=5)

        self.lbl_filename = tk.Label(frame_filename, text="合併輸出檔名:")
        self.lbl_filename.pack(side=tk.LEFT)

        self.entry_filename = tk.Entry(frame_filename, textvariable=self.custom_filename, width=25)
        self.entry_filename.pack(side=tk.LEFT, padx=5)

        tk.Label(frame_filename, text=".pdf").pack(side=tk.LEFT)

        # 3. 輸出位置
        self.btn_output_dir = tk.Button(frame_mid, text="選擇輸出資料夾", command=self.select_output_dir)
        self.btn_output_dir.grid(row=2, column=0, pady=10, sticky=tk.W)

        self.lbl_output_dir = tk.Label(frame_mid, text=f"路徑: {self.output_dir}", wraplength=300, justify=tk.LEFT,
                                       fg="gray")
        self.lbl_output_dir.grid(row=2, column=1, sticky=tk.W, padx=10)

        # 4. 背景顏色選擇
        self.btn_color = tk.Button(frame_mid, text="選擇 PNG 透明背景替換色", command=self.choose_color)
        self.btn_color.grid(row=3, column=0, pady=5, sticky=tk.W)

        self.color_display = tk.Label(frame_mid, text="      ", bg="#FFFFFF", relief="solid", borderwidth=1)
        self.color_display.grid(row=3, column=1, sticky=tk.W, padx=10)
        self.lbl_color_rgb = tk.Label(frame_mid, text="(255, 255, 255)", fg="gray")
        self.lbl_color_rgb.grid(row=3, column=1, sticky=tk.W, padx=45)

        # --- 執行區 ---
        frame_bottom = tk.Frame(self.root, pady=10)
        frame_bottom.pack(fill=tk.X, padx=20)

        self.btn_convert = tk.Button(frame_bottom, text="開始轉換", command=self.start_conversion, bg="#4CAF50",
                                     fg="white", font=("Arial", 12, "bold"), height=2)
        self.btn_convert.pack(fill=tk.X)

    # --- 互動功能綁定 ---
    def toggle_filename_entry(self):
        """根據單選按鈕的狀態，決定是否讓使用者輸入檔名"""
        if self.merge_mode.get():
            self.entry_filename.config(state=tk.NORMAL)
            self.lbl_filename.config(fg="black")
        else:
            self.entry_filename.config(state=tk.DISABLED)
            self.lbl_filename.config(fg="gray")

    def select_images(self):
        paths = filedialog.askopenfilenames(
            title="請選擇圖片",
            filetypes=[("Image Files", "*.png *.jpg *.jpeg *.bmp *.gif")]
        )
        if paths:
            self.image_paths = list(paths)
            self.lbl_img_count.config(text=f"目前已選擇: {len(self.image_paths)} 張圖片")

    def select_output_dir(self):
        dir_path = filedialog.askdirectory(title="選擇 PDF 輸出位置")
        if dir_path:
            self.output_dir = dir_path
            self.lbl_output_dir.config(text=f"路徑: {self.output_dir}")

    def choose_color(self):
        color = colorchooser.askcolor(title="請選擇背景顏色", initialcolor=self.bg_color)
        if color[0]:
            self.bg_color = tuple(int(c) for c in color[0])
            hex_color = color[1]
            self.color_display.config(bg=hex_color)
            self.lbl_color_rgb.config(text=str(self.bg_color))

    def start_conversion(self):
        if not self.image_paths:
            messagebox.showwarning("警告", "請先選擇至少一張圖片！")
            return

        is_merge = self.merge_mode.get()
        success_count = 0

        self.btn_convert.config(state=tk.DISABLED, text="轉換中...")
        self.root.update()

        if is_merge:
            # --- 模式 1：全部合併 (使用自訂檔名) ---
            filename = self.custom_filename.get().strip()

            # 防呆：如果使用者把輸入框清空，給它一個備用的時間檔名
            if not filename:
                filename = datetime.now().strftime("%Y%m%d_%H%M%S")

            # 防呆：避免使用者重複輸入 .pdf
            if not filename.lower().endswith(".pdf"):
                filename += ".pdf"

            output_file = os.path.join(self.output_dir, filename)

            success, msg = convert_images_to_pdf(self.image_paths, output_file, self.bg_color)
            if success:
                messagebox.showinfo("成功", f"合併轉換完成！\n\n{msg}")
            else:
                messagebox.showerror("錯誤", msg)
        else:
            # --- 模式 2：獨立檔案 ---
            for path in self.image_paths:
                base_name = os.path.splitext(os.path.basename(path))[0]
                output_file = os.path.join(self.output_dir, f"{base_name}.pdf")

                success, msg = convert_images_to_pdf([path], output_file, self.bg_color)
                if success:
                    success_count += 1

            if success_count == len(self.image_paths):
                messagebox.showinfo("成功", f"成功將 {success_count} 張圖片分別轉換為 PDF！\n儲存於：{self.output_dir}")
            else:
                messagebox.showwarning("部分成功", f"完成了 {success_count}/{len(self.image_paths)} 張圖片的轉換。")

        self.btn_convert.config(state=tk.NORMAL, text="開始轉換")


if __name__ == "__main__":
    root = tk.Tk()
    app = ImageToPdfApp(root)
    root.mainloop()