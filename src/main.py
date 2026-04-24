# main.py
import os
import tkinter as tk
from tkinter import filedialog, messagebox
from datetime import datetime
import customtkinter as ctk
from CTkColorPicker import AskColor
from tkinterdnd2 import TkinterDnD, DND_FILES

# 呼叫我們的核心轉檔模組
from pdf_converter import convert_images_to_pdf

# 設定 CustomTkinter 的全域外觀
ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")


# ==========================================
# 融合 CustomTkinter 與 TkinterDnD2
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
        self.root.geometry("620x800")
        self.root.resizable(False, False)

        # 變數初始化
        self.image_paths = []
        self.selected_index = None  # 記錄目前在自訂列表中選取的項目

        self.output_dir = os.getcwd()
        self.bg_color = (255, 255, 255)
        self.hex_color = "#FFFFFF"

        self.merge_mode = ctk.BooleanVar(value=True)
        self.last_auto_filename = datetime.now().strftime("%Y%m%d_%H%M%S")
        self.custom_filename = ctk.StringVar(value=self.last_auto_filename)

        self.create_widgets()
        self.setup_dnd()
        self.update_listbox()  # 初始化畫面狀態

    def create_widgets(self):
        title_lbl = ctk.CTkLabel(self.root, text="圖片轉 PDF 工具", font=ctk.CTkFont(size=24, weight="bold"))
        title_lbl.pack(pady=(15, 10))

        # --- 卡片 1：整合型圖片來源與列表區 ---
        frame_top = ctk.CTkFrame(self.root, corner_radius=10)
        frame_top.pack(fill="x", padx=30, pady=5)

        ctk.CTkLabel(frame_top, text="1. 圖片來源與排序", font=ctk.CTkFont(size=16, weight="bold")).pack(anchor="w",
                                                                                                         padx=20,
                                                                                                         pady=(15, 5))

        # 【全新升級：整合拖放區與現代化清單】
        frame_list_area = ctk.CTkFrame(frame_top, fg_color="transparent")
        frame_list_area.pack(fill="x", padx=20, pady=5)

        # 使用 ScrollableFrame 作為自訂清單的容器
        self.list_frame = ctk.CTkScrollableFrame(
            frame_list_area,
            height=160,
            corner_radius=8,
            fg_color=("gray85", "gray20"),
            border_width=2,
            border_color=("gray60", "gray40")
        )
        self.list_frame.pack(side="left", fill="both", expand=True)

        # 右側控制按鈕
        frame_list_btns = ctk.CTkFrame(frame_list_area, fg_color="transparent")
        frame_list_btns.pack(side="left", padx=(10, 0))

        ctk.CTkButton(frame_list_btns, text="🔼 上移", width=80, command=self.move_up).pack(pady=(0, 5))
        ctk.CTkButton(frame_list_btns, text="🔽 下移", width=80, command=self.move_down).pack(pady=(0, 5))
        ctk.CTkButton(frame_list_btns, text="❌ 移除選定", width=80, fg_color="#FF9800", hover_color="#F57C00",
                      command=self.remove_selected).pack(pady=(0, 5))

        # 底部操作按鈕區
        frame_btns = ctk.CTkFrame(frame_top, fg_color="transparent")
        frame_btns.pack(anchor="w", padx=20, pady=(10, 15))

        self.btn_add_imgs = ctk.CTkButton(frame_btns, text="➕ 點擊選擇圖片", command=self.add_images, width=120)
        self.btn_add_imgs.pack(side="left", padx=(0, 15))
        self.btn_clear_imgs = ctk.CTkButton(frame_btns, text="🗑️ 清空全部", command=self.clear_images,
                                            fg_color="#F44336", hover_color="#D32F2F", width=120)
        self.btn_clear_imgs.pack(side="left")

        self.lbl_img_count = ctk.CTkLabel(frame_top, text="目前已加入: 0 張圖片", text_color=("gray10", "gray80"))
        self.lbl_img_count.pack(anchor="w", padx=20, pady=(0, 15))

        # --- 卡片 2：輸出設定區 ---
        frame_mid = ctk.CTkFrame(self.root, corner_radius=10)
        frame_mid.pack(fill="x", padx=30, pady=10)

        ctk.CTkLabel(frame_mid, text="2. 輸出設定", font=ctk.CTkFont(size=16, weight="bold")).pack(anchor="w", padx=20,
                                                                                                   pady=(15, 5))

        frame_radio = ctk.CTkFrame(frame_mid, fg_color="transparent")
        frame_radio.pack(fill="x", padx=20, pady=5)
        ctk.CTkRadioButton(frame_radio, text="全部合併成單一 PDF", variable=self.merge_mode, value=True,
                           command=self.toggle_filename_entry).pack(side="left", padx=(0, 20))
        ctk.CTkRadioButton(frame_radio, text="每張獨立轉成一個 PDF", variable=self.merge_mode, value=False,
                           command=self.toggle_filename_entry).pack(side="left")

        frame_filename = ctk.CTkFrame(frame_mid, fg_color="transparent")
        frame_filename.pack(fill="x", padx=20, pady=10)
        self.lbl_filename = ctk.CTkLabel(frame_filename, text="合併檔名:")
        self.lbl_filename.pack(side="left", padx=(0, 10))
        self.entry_filename = ctk.CTkEntry(frame_filename, textvariable=self.custom_filename, width=250)
        self.entry_filename.pack(side="left")
        ctk.CTkLabel(frame_filename, text=".pdf").pack(side="left", padx=(5, 0))

        frame_dir = ctk.CTkFrame(frame_mid, fg_color="transparent")
        frame_dir.pack(fill="x", padx=20, pady=5)
        self.btn_output_dir = ctk.CTkButton(frame_dir, text="📂 選擇輸出資料夾", command=self.select_output_dir,
                                            fg_color="transparent", border_width=2, text_color=("gray10", "#DCE4EE"))
        self.btn_output_dir.pack(side="left", padx=(0, 10))
        self.lbl_output_dir = ctk.CTkLabel(frame_dir, text=f"路徑: {self.output_dir}", wraplength=250, justify="left",
                                           text_color=("gray50", "gray70"))
        self.lbl_output_dir.pack(side="left", fill="x", expand=True)

        frame_color = ctk.CTkFrame(frame_mid, fg_color="transparent")
        frame_color.pack(fill="x", padx=20, pady=(10, 20))
        ctk.CTkLabel(frame_color, text="🎨 PNG 替換背景色:").pack(side="left", padx=(0, 10))
        self.color_display = ctk.CTkLabel(frame_color, text="", width=40, height=30, fg_color=self.hex_color,
                                          corner_radius=6, cursor="hand2")
        self.color_display.pack(side="left", padx=(0, 10))
        self.color_display.bind("<Button-1>", self.choose_color)
        self.lbl_color_rgb = ctk.CTkLabel(frame_color, text="(255, 255, 255)", text_color=("gray50", "gray70"))
        self.lbl_color_rgb.pack(side="left")

        # --- 執行區 ---
        frame_bottom = ctk.CTkFrame(self.root, fg_color="transparent")
        frame_bottom.pack(fill="x", padx=30, pady=10)
        self.btn_convert = ctk.CTkButton(frame_bottom, text="🚀 開始轉換", command=self.start_conversion, height=50,
                                         font=ctk.CTkFont(size=18, weight="bold"), fg_color="#4CAF50",
                                         hover_color="#45a049")
        self.btn_convert.pack(fill="x", pady=5)

    # ==========================================
    # 自訂現代化清單渲染邏輯
    # ==========================================

    def update_listbox(self):
        """根據 image_paths 重新渲染自訂清單畫面"""

        # 1. 清空捲動區塊內的所有舊元件
        for widget in self.list_frame.winfo_children():
            widget.destroy()

        # 2. 如果沒有圖片，顯示拖放提示
        if not self.image_paths:
            lbl_drop = ctk.CTkLabel(
                self.list_frame,
                text="📥 將圖片拖放至此區域\n(或程式內任意位置)",
                font=ctk.CTkFont(size=14),
                text_color=("gray40", "gray60")
            )
            lbl_drop.pack(expand=True, fill="both", pady=40)
            self.selected_index = None  # 清空選取狀態
            return

        # 3. 如果有圖片，逐一渲染成按鈕外觀的「列表項目」
        for i, path in enumerate(self.image_paths):
            is_selected = (i == self.selected_index)

            # 根據是否選中來決定顏色
            bg_color = "#1F6AA5" if is_selected else "transparent"
            hover_color = "#144870" if is_selected else ("gray75", "gray25")
            text_col = "white" if is_selected else ("black", "white")

            btn_item = ctk.CTkButton(
                self.list_frame,
                text=f"{i + 1}. {os.path.basename(path)}",
                anchor="w",  # 文字靠左對齊
                fg_color=bg_color,
                hover_color=hover_color,
                text_color=text_col,
                corner_radius=6,
                height=32,
                command=lambda idx=i: self.select_item(idx)  # 綁定點擊事件
            )
            btn_item.pack(fill="x", padx=5, pady=2)

    def select_item(self, index):
        """處理清單項目的點擊選取"""
        self.selected_index = index
        self.update_listbox()  # 重新渲染以更新選中顏色

    # ==========================================
    # 互動功能與列表管理邏輯
    # ==========================================

    def move_up(self):
        if self.selected_index is not None and self.selected_index > 0:
            idx = self.selected_index
            self.image_paths[idx - 1], self.image_paths[idx] = self.image_paths[idx], self.image_paths[idx - 1]
            self.selected_index -= 1  # 讓選取狀態跟著往上走
            self.update_listbox()

    def move_down(self):
        if self.selected_index is not None and self.selected_index < len(self.image_paths) - 1:
            idx = self.selected_index
            self.image_paths[idx + 1], self.image_paths[idx] = self.image_paths[idx], self.image_paths[idx + 1]
            self.selected_index += 1  # 讓選取狀態跟著往下走
            self.update_listbox()

    def remove_selected(self):
        if self.selected_index is not None:
            del self.image_paths[self.selected_index]
            self.selected_index = None  # 刪除後取消選取狀態
            self.update_listbox()

            self.lbl_img_count.configure(text=f"目前已加入: {len(self.image_paths)} 張圖片")
            self.refresh_auto_filename()

    def setup_dnd(self):
        # 依然將整個視窗註冊為拖放目標，方便使用者隨意拖放
        self.root.drop_target_register(DND_FILES)
        self.root.dnd_bind('<<Drop>>', self.handle_drop)

    def process_added_images(self, paths):
        valid_extensions = ('.png', '.jpg', '.jpeg', '.bmp', '.gif')
        new_to_add = []
        duplicates_found = []

        for path in paths:
            normalized_path = os.path.normpath(path)
            if normalized_path.lower().endswith(valid_extensions):
                if normalized_path in self.image_paths:
                    duplicates_found.append(normalized_path)
                else:
                    new_to_add.append(normalized_path)

        added_count = len(new_to_add)
        self.image_paths.extend(new_to_add)

        if duplicates_found:
            keep_duplicates = messagebox.askyesno(
                "發現重複圖片",
                f"您加入的檔案中有 {len(duplicates_found)} 張圖片已經在列表中了。\n\n請問是否要「重複加入」這些圖片？"
            )
            if keep_duplicates:
                self.image_paths.extend(duplicates_found)
                added_count += len(duplicates_found)

        if added_count > 0:
            self.update_listbox()
            self.lbl_img_count.configure(text=f"目前已加入: {len(self.image_paths)} 張圖片")
            self.refresh_auto_filename()
            return True
        elif not new_to_add and not duplicates_found and paths:
            messagebox.showinfo("提示", "加入失敗，請確保選擇的是支援的圖片格式！")
            return False

        return False

    def handle_drop(self, event):
        dropped_files = self.root.tk.splitlist(event.data)
        success = self.process_added_images(dropped_files)

        if success:
            # 成功加入後，讓自訂清單的邊框閃爍綠色作為回饋
            self.list_frame.configure(border_color="#4CAF50", border_width=2)
            self.root.after(500, lambda: self.list_frame.configure(border_color=("gray60", "gray40")))

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
            self.process_added_images(paths)

    def clear_images(self):
        if self.image_paths:
            self.image_paths.clear()
            self.selected_index = None
            self.update_listbox()
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
    root = CTk_DnD()
    app = ImageToPdfApp(root)
    root.mainloop()