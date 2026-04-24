import os
import tkinter as tk
from tkinter import filedialog, messagebox
from datetime import datetime
from PIL import Image
import customtkinter as ctk
from CTkColorPicker import AskColor
from tkinterdnd2 import TkinterDnD, DND_FILES  # 引入拖放套件

# 設定 CustomTkinter 的全域外觀
ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")


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
# 黑科技：將 CustomTkinter 與 TkinterDnD2 融合
# ==========================================
class CTk_DnD(ctk.CTk, TkinterDnD.DnDWrapper):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.TkdndVersion = TkinterDnD._require(self)


# ==========================================
# 現代化圖形化介面 (GUI) 模組
# ==========================================
class ImageToPdfApp:
    def __init__(self, root):
        self.root = root
        self.root.title("圖片轉 PDF 專業版")
        self.root.geometry("600x680")  # 稍微加高以容納拖放區
        self.root.resizable(False, False)

        # 變數初始化
        self.image_paths = []
        self.output_dir = os.getcwd()
        self.bg_color = (255, 255, 255)
        self.hex_color = "#FFFFFF"

        self.merge_mode = ctk.BooleanVar(value=True)
        self.last_auto_filename = datetime.now().strftime("%Y%m%d_%H%M%S")
        self.custom_filename = ctk.StringVar(value=self.last_auto_filename)

        self.create_widgets()

    def create_widgets(self):
        title_lbl = ctk.CTkLabel(self.root, text="圖片轉 PDF 工具", font=ctk.CTkFont(size=24, weight="bold"))
        title_lbl.pack(pady=(20, 10))

        # --- 卡片 1：圖片選擇區 (新增拖放熱區) ---
        frame_top = ctk.CTkFrame(self.root, corner_radius=10)
        frame_top.pack(fill="x", padx=30, pady=10)

        ctk.CTkLabel(frame_top, text="1. 圖片來源", font=ctk.CTkFont(size=16, weight="bold")).pack(anchor="w", padx=20,
                                                                                                   pady=(15, 5))

        # 拖放熱區視覺設計 (Drop Zone)
        self.drop_zone = ctk.CTkFrame(frame_top, height=100, corner_radius=8, fg_color=("gray85", "gray25"),
                                      border_width=2, border_color=("gray60", "gray40"))
        self.drop_zone.pack(fill="x", padx=20, pady=10)
        self.drop_zone.pack_propagate(False)  # 讓 Frame 保持設定的高度

        lbl_drop = ctk.CTkLabel(self.drop_zone, text="📥 將圖片拖放至此區域\n(支援多選)", font=ctk.CTkFont(size=14),
                                text_color=("gray40", "gray60"))
        lbl_drop.pack(expand=True)

        # 【註冊拖放事件】讓整個 drop_zone 都能接收檔案
        self.drop_zone.drop_target_register(DND_FILES)
        self.drop_zone.dnd_bind('<<Drop>>', self.handle_drop)
        # 讓文字標籤也支援拖放，避免滑鼠指在文字上時失效
        lbl_drop.drop_target_register(DND_FILES)
        lbl_drop.dnd_bind('<<Drop>>', self.handle_drop)

        # 按鈕區
        frame_btns = ctk.CTkFrame(frame_top, fg_color="transparent")
        frame_btns.pack(anchor="w", padx=20, pady=(0, 5))

        self.btn_add_imgs = ctk.CTkButton(frame_btns, text="➕ 點擊選擇圖片", command=self.add_images, width=120)
        self.btn_add_imgs.pack(side="left", padx=(0, 15))

        self.btn_clear_imgs = ctk.CTkButton(frame_btns, text="🗑️ 清空列表", command=self.clear_images,
                                            fg_color="#F44336", hover_color="#D32F2F", width=120)
        self.btn_clear_imgs.pack(side="left")

        self.lbl_img_count = ctk.CTkLabel(frame_top, text="目前已加入: 0 張圖片", text_color=("gray10", "gray80"))
        self.lbl_img_count.pack(anchor="w", padx=20, pady=(0, 15))

        # --- 卡片 2：輸出設定區 ---
        frame_mid = ctk.CTkFrame(self.root, corner_radius=10)
        frame_mid.pack(fill="x", padx=30, pady=10)

        ctk.CTkLabel(frame_mid, text="2. 輸出設定", font=ctk.CTkFont(size=16, weight="bold")).pack(anchor="w", padx=20,
                                                                                                   pady=(15, 5))

        # 選項 1: 合併或分開
        frame_radio = ctk.CTkFrame(frame_mid, fg_color="transparent")
        frame_radio.pack(fill="x", padx=20, pady=5)

        ctk.CTkRadioButton(frame_radio, text="全部合併成單一 PDF", variable=self.merge_mode, value=True,
                           command=self.toggle_filename_entry).pack(side="left", padx=(0, 20))
        ctk.CTkRadioButton(frame_radio, text="每張獨立轉成一個 PDF", variable=self.merge_mode, value=False,
                           command=self.toggle_filename_entry).pack(side="left")

        # 選項 2: 自訂檔名
        frame_filename = ctk.CTkFrame(frame_mid, fg_color="transparent")
        frame_filename.pack(fill="x", padx=20, pady=10)

        self.lbl_filename = ctk.CTkLabel(frame_filename, text="合併檔名:")
        self.lbl_filename.pack(side="left", padx=(0, 10))

        self.entry_filename = ctk.CTkEntry(frame_filename, textvariable=self.custom_filename, width=250)
        self.entry_filename.pack(side="left")

        ctk.CTkLabel(frame_filename, text=".pdf").pack(side="left", padx=(5, 0))

        # 選項 3: 輸出位置
        frame_dir = ctk.CTkFrame(frame_mid, fg_color="transparent")
        frame_dir.pack(fill="x", padx=20, pady=5)

        self.btn_output_dir = ctk.CTkButton(frame_dir, text="📂 選擇輸出資料夾", command=self.select_output_dir,
                                            fg_color="transparent", border_width=2, text_color=("gray10", "#DCE4EE"))
        self.btn_output_dir.pack(side="left", padx=(0, 10))

        self.lbl_output_dir = ctk.CTkLabel(frame_dir, text=f"路徑: {self.output_dir}", wraplength=250, justify="left",
                                           text_color=("gray50", "gray70"))
        self.lbl_output_dir.pack(side="left", fill="x", expand=True)

        # 選項 4: 背景顏色選擇
        frame_color = ctk.CTkFrame(frame_mid, fg_color="transparent")
        frame_color.pack(fill="x", padx=20, pady=(10, 20))

        ctk.CTkLabel(frame_color, text="🎨 PNG 替換背景色:").pack(side="left", padx=(0, 10))

        self.color_display = ctk.CTkLabel(frame_color, text="", width=40, height=30, fg_color=self.hex_color,
                                          corner_radius=6, cursor="hand2")
        self.color_display.pack(side="left", padx=(0, 10))
        self.color_display.bind("<Button-1>", self.choose_color)

        self.lbl_color_rgb = ctk.CTkLabel(frame_color, text="(255, 255, 255)", text_color=("gray50", "gray70"))
        self.lbl_color_rgb.pack(side="left")
        ctk.CTkLabel(frame_color, text="(點擊色塊可更改)", text_color="gray", font=ctk.CTkFont(size=12)).pack(
            side="left", padx=(10, 0))

        # --- 執行區 ---
        frame_bottom = ctk.CTkFrame(self.root, fg_color="transparent")
        frame_bottom.pack(fill="x", padx=30, pady=10)

        self.btn_convert = ctk.CTkButton(frame_bottom, text="🚀 開始轉換", command=self.start_conversion, height=50,
                                         font=ctk.CTkFont(size=18, weight="bold"), fg_color="#4CAF50",
                                         hover_color="#45a049")
        self.btn_convert.pack(fill="x", pady=10)

    # --- 互動功能綁定 ---

    # 【全新加入的拖放處理邏輯】
    def handle_drop(self, event):
        """解析拖放進來的檔案路徑並加入列表"""
        # tkinterdnd2 會將多個檔案路徑轉為字串，如果路徑有空白會用大括號 {} 包起來
        # 使用 root.tk.splitlist 是最安全、最標準的解析方式
        dropped_files = self.root.tk.splitlist(event.data)

        valid_extensions = ('.png', '.jpg', '.jpeg', '.bmp', '.gif')
        new_count = 0

        for path in dropped_files:
            # 轉換為標準路徑格式，並過濾掉不是圖片的檔案
            normalized_path = os.path.normpath(path)
            if normalized_path.lower().endswith(valid_extensions):
                if normalized_path not in self.image_paths:
                    self.image_paths.append(normalized_path)
                    new_count += 1

        self.lbl_img_count.configure(text=f"目前已加入: {len(self.image_paths)} 張圖片")
        self.refresh_auto_filename()

        if new_count > 0:
            # 讓拖放區閃爍一下綠色作為成功的回饋
            self.drop_zone.configure(border_color="#4CAF50", border_width=2)
            self.root.after(500, lambda: self.drop_zone.configure(border_color=("gray60", "gray40")))
        elif len(dropped_files) > 0:
            messagebox.showinfo("提示", "加入失敗，請確保拖放的是支援的圖片格式！")

    def refresh_auto_filename(self):
        current_val = self.custom_filename.get().strip()
        if current_val == self.last_auto_filename or not current_val:
            new_time_str = datetime.now().strftime("%Y%m%d_%H%M%S")
            self.custom_filename.set(new_time_str)
            self.last_auto_filename = new_time_str

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

            self.lbl_img_count.configure(text=f"目前已加入: {len(self.image_paths)} 張圖片")
            self.refresh_auto_filename()

    def clear_images(self):
        if self.image_paths:
            self.image_paths.clear()
            self.lbl_img_count.configure(text="目前已加入: 0 張圖片")
            self.refresh_auto_filename()

    def toggle_filename_entry(self):
        if self.merge_mode.get():
            self.entry_filename.configure(state="normal")
            self.lbl_filename.configure(text_color=("black", "white"))
        else:
            self.entry_filename.configure(state="disabled")
            self.lbl_filename.configure(text_color="gray")

    def select_output_dir(self):
        dir_path = filedialog.askdirectory(title="選擇 PDF 輸出位置")
        if dir_path:
            self.output_dir = dir_path
            self.lbl_output_dir.configure(text=f"路徑: {self.output_dir}")

    def choose_color(self, event=None):
        pick_color = AskColor(title="選擇透明背景替換色", initial_color=self.hex_color)
        selected_hex = pick_color.get()

        if selected_hex:
            self.hex_color = selected_hex
            hex_str = self.hex_color.lstrip('#')
            self.bg_color = tuple(int(hex_str[i:i + 2], 16) for i in (0, 2, 4))

            self.color_display.configure(fg_color=self.hex_color)
            self.lbl_color_rgb.configure(text=str(self.bg_color))
            self.root.update_idletasks()

    def start_conversion(self):
        if not self.image_paths:
            messagebox.showwarning("警告", "請先新增至少一張圖片！")
            return

        is_merge = self.merge_mode.get()
        success_count = 0

        self.btn_convert.configure(state="disabled", text="轉換中...")
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

        self.btn_convert.configure(state="normal", text="🚀 開始轉換")


if __name__ == "__main__":
    # 【關鍵修改】：不再使用標準的 ctk.CTk()，而是改用我們混合了 DnD 功能的自訂類別
    root = CTk_DnD()
    app = ImageToPdfApp(root)
    root.mainloop()