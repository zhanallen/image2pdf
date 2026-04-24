import os
import tkinter as tk
from tkinter import filedialog, colorchooser, messagebox
from datetime import datetime
from PIL import Image


# ==========================================
# 核心轉換模組 (維持不變)
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
        self.root.geometry("520x520")
        self.root.resizable(False, False)

        # 變數初始化
        self.image_paths = []
        self.output_dir = os.getcwd()
        self.bg_color = (255, 255, 255)
        self.merge_mode = tk.BooleanVar(value=True)

        # 紀錄程式「上一次自動產生的檔名」
        self.last_auto_filename = datetime.now().strftime("%Y%m%d_%H%M%S")
        self.custom_filename = tk.StringVar(value=self.last_auto_filename)

        self.create_widgets()

    def create_widgets(self):
        # --- 圖片選擇區 ---
        frame_top = tk.Frame(self.root, pady=10)
        frame_top.pack(fill=tk.X, padx=20)

        tk.Label(frame_top, text="1. 準備圖片", font=("Arial", 10, "bold")).pack(anchor=tk.W, pady=(0, 5))

        frame_btns = tk.Frame(frame_top)
        frame_btns.pack(anchor=tk.W)

        self.btn_add_imgs = tk.Button(frame_btns, text="➕ 新增圖片", command=self.add_images, width=15)
        self.btn_add_imgs.pack(side=tk.LEFT, padx=(0, 10))

        self.btn_clear_imgs = tk.Button(frame_btns, text="🗑️ 清空列表", command=self.clear_images, width=15)
        self.btn_clear_imgs.pack(side=tk.LEFT)

        self.lbl_img_count = tk.Label(frame_top, text="目前已加入: 0 張圖片", fg="blue")
        self.lbl_img_count.pack(anchor=tk.W, pady=5)

        # --- 輸出選項區 ---
        frame_mid = tk.LabelFrame(self.root, text="2. 輸出設定", pady=10, padx=10)
        frame_mid.pack(fill=tk.X, padx=20, pady=10)

        tk.Radiobutton(frame_mid, text="全部合併成單一 PDF", variable=self.merge_mode, value=True,
                       command=self.toggle_filename_entry).grid(row=0, column=0, sticky=tk.W)
        tk.Radiobutton(frame_mid, text="每張獨立轉成一個 PDF", variable=self.merge_mode, value=False,
                       command=self.toggle_filename_entry).grid(row=0, column=1, sticky=tk.W)

        frame_filename = tk.Frame(frame_mid)
        frame_filename.grid(row=1, column=0, columnspan=2, sticky=tk.W, pady=5)

        self.lbl_filename = tk.Label(frame_filename, text="合併輸出檔名:")
        self.lbl_filename.pack(side=tk.LEFT)

        self.entry_filename = tk.Entry(frame_filename, textvariable=self.custom_filename, width=25)
        self.entry_filename.pack(side=tk.LEFT, padx=5)

        tk.Label(frame_filename, text=".pdf").pack(side=tk.LEFT)

        self.btn_output_dir = tk.Button(frame_mid, text="選擇輸出資料夾", command=self.select_output_dir)
        self.btn_output_dir.grid(row=2, column=0, pady=10, sticky=tk.W)

        self.lbl_output_dir = tk.Label(frame_mid, text=f"路徑: {self.output_dir}", wraplength=300, justify=tk.LEFT,
                                       fg="gray")
        self.lbl_output_dir.grid(row=2, column=1, sticky=tk.W, padx=10)

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

    def refresh_auto_filename(self):
        """檢查並更新預設的檔名"""
        current_val = self.custom_filename.get().strip()

        # 如果現在輸入框的值等於「上一次自動產生的值」，或者使用者把它刪到全空
        # 就代表使用者沒有自己決定檔名，我們就幫他更新成當下最新時間
        if current_val == self.last_auto_filename or not current_val:
            new_time_str = datetime.now().strftime("%Y%m%d_%H%M%S")
            self.custom_filename.set(new_time_str)
            self.last_auto_filename = new_time_str  # 更新紀錄

    def add_images(self):
        paths = filedialog.askopenfilenames(
            title="請選擇要加入的圖片",
            filetypes=[("Image Files", "*.png *.jpg *.jpeg *.bmp *.gif")]
        )
        if paths:
            new_count = 0
            for path in paths:
                if path not in self.image_paths:
                    self.image_paths.append(path)
                    new_count += 1

            self.lbl_img_count.config(text=f"目前已加入: {len(self.image_paths)} 張圖片")

            # 每次成功加入圖片後，嘗試更新時間檔名
            self.refresh_auto_filename()

            if new_count > 0:
                print(f"成功加入 {new_count} 張圖片。")
            else:
                print("所選圖片皆已在列表中。")

    def clear_images(self):
        if self.image_paths:
            self.image_paths.clear()
            self.lbl_img_count.config(text="目前已加入: 0 張圖片")

            # 重新開始，也嘗試更新時間檔名
            self.refresh_auto_filename()
            print("圖片列表已清空。")

    def toggle_filename_entry(self):
        if self.merge_mode.get():
            self.entry_filename.config(state=tk.NORMAL)
            self.lbl_filename.config(fg="black")
        else:
            self.entry_filename.config(state=tk.DISABLED)
            self.lbl_filename.config(fg="gray")

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
            messagebox.showwarning("警告", "請先新增至少一張圖片！")
            return

        is_merge = self.merge_mode.get()
        success_count = 0

        self.btn_convert.config(state=tk.DISABLED, text="轉換中...")
        self.root.update()

        if is_merge:
            filename = self.custom_filename.get().strip()
            if not filename:
                filename = datetime.now().strftime("%Y%m%d_%H%M%S")
            if not filename.lower().endswith(".pdf"):
                filename += ".pdf"

            output_file = os.path.join(self.output_dir, filename)
            success, msg = convert_images_to_pdf(self.image_paths, output_file, self.bg_color)
            if success:
                messagebox.showinfo("成功", f"合併轉換完成！\n\n{msg}")
            else:
                messagebox.showerror("錯誤", msg)
        else:
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