# main.py
import os
import sys
import platform
import tkinter as tk
from tkinter import filedialog, messagebox
from datetime import datetime
import threading
import customtkinter as ctk
import CTkColorPicker
from tkinterdnd2 import TkinterDnD, DND_FILES

try:
    import pywinstyles

    HAS_WINSTYLES = True
except ImportError:
    HAS_WINSTYLES = False

from pdf_converter import convert_images_to_pdf

ctk_cp_path = os.path.dirname(CTkColorPicker.__file__)
# ==========================================
# Material 3 風格與全域設定
# ==========================================
ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

M3_CARD_RADIUS = 16
M3_BTN_RADIUS = 20
M3_CARD_COLOR = ("#F2F2F7", "#1E1E1E")
M3_PRIMARY_CONTAINER = ("#D3E3FD", "#004A77")
M3_ON_PRIMARY_CONTAINER = ("#041E49", "#C2E7FF")
M3_PRIMARY = ("#006493", "#8ECAE6")
M3_ON_PRIMARY = ("#FFFFFF", "#00344F")


class CTk_DnD(ctk.CTk, TkinterDnD.DnDWrapper):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.TkdndVersion = TkinterDnD._require(self)


class EditableAskColor(ctk.CTkToplevel):
    """客製化的調色盤彈窗，新增「獨立顏色預覽區塊」與手動輸入"""

    def __init__(self, title="選擇透明背景替換色", initial_color="#FFFFFF", **kwargs):
        super().__init__(**kwargs)
        self.title(title)
        self.geometry("320x380")  # 稍微再拉高一點點，讓排版更舒適
        self.resizable(False, False)
        self.attributes("-topmost", True)

        self.color = initial_color
        self.last_color = initial_color

        # 1. 裝載調色盤的容器
        self.picker_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.picker_frame.pack(pady=(10, 5))

        self._build_picker(initial_color)

        # 2. 建立底部輸入區容器 (水平排列預覽框與輸入框)
        self.input_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.input_frame.pack(pady=10)

        # 【新增】：獨立的顏色預覽大方塊，解決拉桿預覽太小/被擋住的問題
        self.color_preview = ctk.CTkFrame(
            self.input_frame,
            width=40, height=40,
            corner_radius=8,
            border_width=1, border_color=("gray70", "gray30"),
            fg_color=initial_color
        )
        self.color_preview.pack(side="left", padx=(0, 10))

        # 輸入框
        self.hex_entry = ctk.CTkEntry(
            self.input_frame,
            width=120, height=40, corner_radius=8, justify="center",
            font=ctk.CTkFont(weight="bold", size=14)
        )
        self.hex_entry.pack(side="left")
        self.hex_entry.insert(0, initial_color)

        self.hex_entry.bind("<Return>", self._on_hex_enter)
        self.hex_entry.bind("<FocusOut>", self._on_hex_enter)

        self._poll_color()

        # 確認按鈕
        self.button = ctk.CTkButton(self, text="確定", height=36, command=self._ok_event)
        self.button.pack(pady=(5, 20))

        self.grab_set()
        self.wait_window()

    def _build_picker(self, color):
        """建構或重建調色盤元件"""
        for widget in self.picker_frame.winfo_children():
            widget.destroy()

        self.picker = CTkColorPicker.CTkColorPicker(self.picker_frame, initial_color=color)
        self.picker.pack()

        for child in self.picker.winfo_children():
            if isinstance(child, ctk.CTkLabel) and child.cget("text").startswith("#"):
                child.pack_forget()
                break

    def _on_hex_enter(self, event=None):
        """處理使用者手動輸入色碼"""
        hex_val = self.hex_entry.get().strip().upper()
        if len(hex_val) > 0 and not hex_val.startswith("#"):
            hex_val = "#" + hex_val

        if len(hex_val) == 7:
            try:
                int(hex_val[1:], 16)
                if hex_val != self.last_color:
                    self._build_picker(hex_val)
                    self.last_color = hex_val
                    # 【同步】：手動輸入成功時，更新大預覽框
                    self.color_preview.configure(fg_color=hex_val)
            except ValueError:
                pass

        self.hex_entry.delete(0, 'end')
        self.hex_entry.insert(0, self.picker.get())

    def _poll_color(self):
        """定期檢查調色盤顏色是否被滑鼠拖拉改變"""
        if not self.winfo_exists(): return
        current_color = self.picker.get()
        if current_color != self.last_color:
            # 【同步】：滑鼠拖曳調色盤時，即時更新大預覽框
            self.color_preview.configure(fg_color=current_color)

            if self.focus_get() != self.hex_entry:
                self.hex_entry.delete(0, 'end')
                self.hex_entry.insert(0, current_color)
            self.last_color = current_color
        self.after(50, self._poll_color)

    def _ok_event(self):
        self._on_hex_enter()
        self.color = self.picker.get()
        self.destroy()

    def get(self):
        return self.color

class ImageToPdfApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Image to PDF")
        self.root.geometry("620x730")
        self.root.resizable(False, False)

        self.apply_glassmorphism()

        # 變數初始化
        self.image_paths = []
        self.selected_index = None
        self.list_btn_refs = []

        self.output_dir = os.getcwd()
        self.bg_color = (255, 255, 255)
        self.hex_color = "#FFFFFF"
        self.merge_mode = ctk.BooleanVar(value=True)
        self.last_auto_filename = datetime.now().strftime("%Y%m%d_%H%M%S")
        self.custom_filename = ctk.StringVar(value=self.last_auto_filename)

        self.create_widgets()
        self.setup_dnd()
        self.update_listbox()

    def apply_glassmorphism(self):
        if not HAS_WINSTYLES:
            return
        if platform.system() == "Windows":
            try:
                build_number = int(platform.version().split('.')[2])
                if build_number >= 22000:
                    # Windows 11 完美支援 Mica
                    pywinstyles.apply_style(self.root, "mica")
                else:
                    # Windows 10 的 Acrylic 在淺色模式容易導致背景變純黑
                    # 【修正】：若是淺色模式，給予原生 M3 淺色背景；若是深色模式再套用特效
                    if ctk.get_appearance_mode() == "Light":
                        self.root.configure(fg_color="#F9F9FA")
                    else:
                        pywinstyles.apply_style(self.root, "acrylic")

                pywinstyles.change_header_color(self.root, color="transparent")
            except Exception as e:
                print(f"特效載入失敗: {e}")

    def create_widgets(self):
        # 略... (維持你原本極優秀的 UI 設計，完全不用動)
        title_lbl = ctk.CTkLabel(
            self.root, text="圖片轉換工具",
            font=ctk.CTkFont(family="Segoe UI Variable Display", size=26, weight="bold"),
            fg_color="transparent",  # 強制背景透明
            text_color=("black", "white")  # 淺色模式黑字，深色模式白字
        )
        title_lbl.pack(pady=(20, 10))

        frame_top = ctk.CTkFrame(self.root, corner_radius=M3_CARD_RADIUS, fg_color=M3_CARD_COLOR)
        frame_top.pack(fill="x", padx=25, pady=8)

        ctk.CTkLabel(frame_top, text="圖片來源", font=ctk.CTkFont(size=16, weight="bold")).pack(anchor="w", padx=20,
                                                                                                pady=(15, 5))

        frame_list_area = ctk.CTkFrame(frame_top, fg_color="transparent")
        frame_list_area.pack(fill="x", padx=20, pady=5)

        self.list_frame = ctk.CTkScrollableFrame(
            frame_list_area, height=120, corner_radius=12,
            fg_color=("gray90", "gray14"), border_width=0
        )
        self.list_frame.pack(side="left", fill="both", expand=True)

        frame_list_btns = ctk.CTkFrame(frame_list_area, fg_color="transparent")
        frame_list_btns.pack(side="left", padx=(10, 0), fill="y")
        inner_btns_frame = ctk.CTkFrame(frame_list_btns, fg_color="transparent")
        inner_btns_frame.pack(expand=True)

        icon_font = ctk.CTkFont(family="Arial", size=16, weight="bold")
        btn_kwargs = {
            "width": 36, "height": 36, "font": icon_font,
            "corner_radius": 6, "anchor": "center",
            "fg_color": ("gray85", "gray25"), "text_color": ("black", "white"),
            "hover_color": ("gray75", "gray35")
        }

        ctk.CTkButton(inner_btns_frame, text="▲", command=self.move_up, **btn_kwargs).pack(pady=(0, 8))
        ctk.CTkButton(inner_btns_frame, text="▼", command=self.move_down, **btn_kwargs).pack(pady=(0, 8))

        btn_kwargs["fg_color"] = ("#FFCDD2", "#4A0000")
        btn_kwargs["text_color"] = ("#B71C1C", "#FFCDD2")
        btn_kwargs["hover_color"] = ("#EF9A9A", "#7A0000")
        ctk.CTkButton(inner_btns_frame, text="✖", command=self.remove_selected, **btn_kwargs).pack()

        frame_btns = ctk.CTkFrame(frame_top, fg_color="transparent")
        frame_btns.pack(fill="x", padx=20, pady=(5, 15))

        self.btn_add_imgs = ctk.CTkButton(
            frame_btns, text="➕ 選擇圖片", command=self.add_images,
            width=110, corner_radius=M3_BTN_RADIUS,
            fg_color=M3_PRIMARY_CONTAINER, text_color=M3_ON_PRIMARY_CONTAINER
        )
        self.btn_add_imgs.pack(side="left", padx=(0, 10))

        self.btn_clear_imgs = ctk.CTkButton(
            frame_btns, text="清空列表", command=self.clear_images,
            width=90, corner_radius=M3_BTN_RADIUS,
            fg_color="transparent", text_color=("gray30", "gray70"), hover_color=("gray85", "gray25")
        )
        self.btn_clear_imgs.pack(side="left")

        # 【核心修正】：將文字標籤的父容器改成 frame_btns，並設定靠右對齊
        self.lbl_img_count = ctk.CTkLabel(
            frame_btns, text="目前加入: 0 張",
            text_color=("gray40", "gray50"), font=ctk.CTkFont(size=12)
        )
        self.lbl_img_count.pack(side="right")

        frame_mid = ctk.CTkFrame(self.root, corner_radius=M3_CARD_RADIUS, fg_color=M3_CARD_COLOR)
        frame_mid.pack(fill="x", padx=25, pady=8)

        ctk.CTkLabel(frame_mid, text="輸出設定", font=ctk.CTkFont(size=16, weight="bold")).pack(anchor="w", padx=20,
                                                                                                pady=(15, 5))

        frame_radio = ctk.CTkFrame(frame_mid, fg_color="transparent")
        frame_radio.pack(fill="x", padx=20, pady=5)
        ctk.CTkRadioButton(frame_radio, text="合併為單一檔案", variable=self.merge_mode, value=True,
                           command=self.toggle_filename_entry).pack(side="left", padx=(0, 20))
        ctk.CTkRadioButton(frame_radio, text="獨立檔案輸出", variable=self.merge_mode, value=False,
                           command=self.toggle_filename_entry).pack(side="left")

        frame_filename = ctk.CTkFrame(frame_mid, fg_color="transparent")
        frame_filename.pack(fill="x", padx=20, pady=10)
        self.lbl_filename = ctk.CTkLabel(frame_filename, text="檔名設定:")
        self.lbl_filename.pack(side="left", padx=(0, 10))
        self.entry_filename = ctk.CTkEntry(frame_filename, textvariable=self.custom_filename, width=220,
                                           corner_radius=8, border_width=0, fg_color=("gray90", "gray14"))
        self.entry_filename.pack(side="left")
        ctk.CTkLabel(frame_filename, text=".pdf").pack(side="left", padx=(5, 0))

        frame_dir = ctk.CTkFrame(frame_mid, fg_color="transparent")
        frame_dir.pack(fill="x", padx=20, pady=5)
        self.btn_output_dir = ctk.CTkButton(
            frame_dir,
            text="📂 變更位置",
            command=self.select_output_dir,
            corner_radius=M3_BTN_RADIUS,
            # 【優化 1】：不使用 transparent，改用背後卡片的底色，能完美解決 1px 邊線破圖
            fg_color=M3_CARD_COLOR,
            # 【優化 2】：明確給予邊框顏色，讓淺色與深色模式都有乾淨的輪廓
            border_color=("gray75", "gray30"),
            border_width=2,
            text_color=("gray20", "gray80"),
            hover_color=("gray85", "gray25")  # 加上 hover 顏色讓點擊回饋更自然
        )
        self.btn_output_dir.pack(side="left", padx=(0, 10))
        self.lbl_output_dir = ctk.CTkLabel(frame_dir, text=f"路徑: {self.output_dir}", wraplength=300, justify="left",
                                           text_color=("gray50", "gray50"), font=ctk.CTkFont(size=12))
        self.lbl_output_dir.pack(side="left", fill="x", expand=True)

        frame_color = ctk.CTkFrame(frame_mid, fg_color="transparent")
        frame_color.pack(fill="x", padx=20, pady=(10, 20))
        ctk.CTkLabel(frame_color, text="PNG 透明色轉換 :").pack(side="left", padx=(0, 10))
        self.color_display = ctk.CTkLabel(frame_color, text="", width=36, height=28, fg_color=self.hex_color,
                                          corner_radius=8, cursor="hand2")
        self.color_display.pack(side="left", padx=(0, 10))
        self.color_display.bind("<Button-1>", self.choose_color)
        self.lbl_color_rgb = ctk.CTkLabel(frame_color, text="(255, 255, 255)", text_color=("gray50", "gray50"))
        self.lbl_color_rgb.pack(side="left")

        frame_bottom = ctk.CTkFrame(self.root, fg_color="transparent")
        frame_bottom.pack(fill="x", padx=25, pady=(10, 20))

        self.btn_convert = ctk.CTkButton(
            frame_bottom, text="開始轉換", command=self.start_conversion,
            height=54, font=ctk.CTkFont(size=18, weight="bold"),
            corner_radius=27,
            fg_color=M3_PRIMARY, text_color=M3_ON_PRIMARY
        )
        self.btn_convert.pack(fill="x")

    def update_listbox(self):
        for widget in self.list_frame.winfo_children():
            widget.destroy()
        self.list_btn_refs.clear()

        if not self.image_paths:
            lbl_drop = ctk.CTkLabel(
                self.list_frame, text="📥 拖放圖片至此",
                font=ctk.CTkFont(size=14, weight="bold"), text_color=("gray50", "gray50")
            )
            lbl_drop.pack(expand=True, fill="both", pady=25)
            self.selected_index = None
            return

        for i, path in enumerate(self.image_paths):
            is_selected = (i == self.selected_index)
            bg_color = M3_PRIMARY_CONTAINER if is_selected else ("gray90", "gray14")
            hover_col = M3_PRIMARY_CONTAINER if is_selected else ("gray85", "gray20")
            text_col = M3_ON_PRIMARY_CONTAINER if is_selected else ("black", "white")

            btn_item = ctk.CTkButton(
                self.list_frame, text=f"{i + 1}. {os.path.basename(path)}", anchor="w",
                fg_color=bg_color, hover_color=hover_col, text_color=text_col,
                corner_radius=8, height=32,
                command=lambda idx=i: self.select_item(idx)
            )
            btn_item.pack(fill="x", padx=4, pady=2)
            self.list_btn_refs.append(btn_item)

    def select_item(self, index):
        self.selected_index = index
        for i, btn in enumerate(self.list_btn_refs):
            is_selected = (i == self.selected_index)
            bg_color = M3_PRIMARY_CONTAINER if is_selected else ("gray90", "gray14")
            hover_col = M3_PRIMARY_CONTAINER if is_selected else ("gray85", "gray20")
            text_col = M3_ON_PRIMARY_CONTAINER if is_selected else ("black", "white")
            btn.configure(fg_color=bg_color, hover_color=hover_col, text_color=text_col)

    def move_up(self):
        if self.selected_index is not None and self.selected_index > 0:
            idx = self.selected_index
            self.image_paths[idx - 1], self.image_paths[idx] = self.image_paths[idx], self.image_paths[idx - 1]
            self.selected_index -= 1
            self.update_listbox()

    def move_down(self):
        if self.selected_index is not None and self.selected_index < len(self.image_paths) - 1:
            idx = self.selected_index
            self.image_paths[idx + 1], self.image_paths[idx] = self.image_paths[idx], self.image_paths[idx + 1]
            self.selected_index += 1
            self.update_listbox()

    def remove_selected(self):
        if self.selected_index is not None:
            del self.image_paths[self.selected_index]
            self.selected_index = None
            self.update_listbox()
            self.lbl_img_count.configure(text=f"目前加入: {len(self.image_paths)} 張")
            self.refresh_auto_filename()

    def setup_dnd(self):
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
            keep_duplicates = messagebox.askyesno("發現重複圖片",
                                                  f"有 {len(duplicates_found)} 張圖片已存在。\n\n是否重複加入？")
            if keep_duplicates:
                self.image_paths.extend(duplicates_found)
                added_count += len(duplicates_found)

        if added_count > 0:
            self.update_listbox()
            self.lbl_img_count.configure(text=f"目前加入: {len(self.image_paths)} 張")
            self.refresh_auto_filename()
            return True
        elif not new_to_add and not duplicates_found and paths:
            messagebox.showinfo("提示", "加入失敗，僅支援圖片格式！")
            return False
        return False

    def handle_drop(self, event):
        dropped_files = self.root.tk.splitlist(event.data)
        success = self.process_added_images(dropped_files)

    def refresh_auto_filename(self):
        current_val = self.custom_filename.get().strip()
        if current_val == self.last_auto_filename or not current_val:
            new_time_str = datetime.now().strftime("%Y%m%d_%H%M%S")
            self.custom_filename.set(new_time_str)
            self.last_auto_filename = new_time_str

    def add_images(self):
        paths = filedialog.askopenfilenames(title="選擇圖片",
                                            filetypes=[("Image Files", "*.png *.jpg *.jpeg *.bmp *.gif")])
        if paths:
            self.process_added_images(paths)

    def clear_images(self):
        if self.image_paths:
            self.image_paths.clear()
            self.selected_index = None
            self.update_listbox()
            self.lbl_img_count.configure(text="目前加入: 0 張")
            self.refresh_auto_filename()

    def toggle_filename_entry(self):
        if self.merge_mode.get():
            self.entry_filename.configure(state="normal")
            self.lbl_filename.configure(text_color=("black", "white"))
        else:
            self.entry_filename.configure(state="disabled")
            self.lbl_filename.configure(text_color="gray")

    def select_output_dir(self):
        dir_path = filedialog.askdirectory(title="選擇輸出位置")
        if dir_path:
            self.output_dir = dir_path
            self.lbl_output_dir.configure(text=f"路徑: {self.output_dir}")

    def choose_color(self, event=None):
        # 【修改這裡】：使用我們自製支援輸入的 EditableAskColor
        pick_color = EditableAskColor(title="選擇透明背景替換色", initial_color=self.hex_color)
        selected_hex = pick_color.get()

        if selected_hex:
            self.hex_color = selected_hex
            hex_str = self.hex_color.lstrip('#')
            self.bg_color = tuple(int(hex_str[i:i + 2], 16) for i in (0, 2, 4))
            self.color_display.configure(fg_color=self.hex_color)
            self.lbl_color_rgb.configure(text=str(self.bg_color))
            self.root.update_idletasks()

    def get_unique_filepath(self, filepath):
        """如果選擇「保留兩者」，自動在檔名後加上 (1), (2)..."""
        base, ext = os.path.splitext(filepath)
        counter = 1
        new_filepath = f"{base}({counter}){ext}"
        while os.path.exists(new_filepath):
            counter += 1
            new_filepath = f"{base}({counter}){ext}"
        return new_filepath

    def prompt_file_exists(self, filename):
        dialog = ctk.CTkToplevel(self.root)
        dialog.title("發現重複檔案")
        dialog.geometry("400x180")
        dialog.resizable(False, False)

        # 【優化】：視窗置中計算
        self.root.update_idletasks()  # 確保讀取到最新的視窗數值
        parent_x = self.root.winfo_x()
        parent_y = self.root.winfo_y()
        parent_width = self.root.winfo_width()
        parent_height = self.root.winfo_height()

        # 計算對話框應該出現的座標 (置於父視窗中心)
        x = parent_x + (parent_width // 2) - (400 // 2)
        y = parent_y + (parent_height // 2) - (180 // 2)
        dialog.geometry(f"400x180+{x}+{y}")

        dialog.attributes("-topmost", True)
        dialog.transient(self.root)
        dialog.grab_set()

        result = tk.StringVar(value="skip")

        ctk.CTkLabel(
            dialog,
            text=f"輸出目錄中已存在檔案：\n「{filename}」\n\n您想要如何處理？",
            justify="center"
        ).pack(pady=20)

        btn_frame = ctk.CTkFrame(dialog, fg_color="transparent")
        btn_frame.pack(pady=10)

        def set_res(val):
            result.set(val)
            dialog.destroy()

        ctk.CTkButton(btn_frame, text="覆蓋", width=90, fg_color="#D32F2F", hover_color="#B71C1C",
                      command=lambda: set_res("overwrite")).pack(side="left", padx=10)
        ctk.CTkButton(btn_frame, text="保留兩者", width=90, command=lambda: set_res("keep")).pack(side="left", padx=10)
        ctk.CTkButton(btn_frame, text="跳過", width=90, fg_color="gray", hover_color="darkgray",
                      command=lambda: set_res("skip")).pack(side="left", padx=10)

        self.root.wait_window(dialog)
        return result.get()

    def start_conversion(self):
        if not self.image_paths:
            messagebox.showwarning("警告", "請先新增圖片！")
            return

        is_merge = self.merge_mode.get()
        custom_name = self.custom_filename.get().strip()
        out_dir = self.output_dir
        bg_col = self.bg_color
        paths = list(self.image_paths)

        # 準備要轉換的任務清單 [(圖片路徑清單, 輸出PDF路徑), ...]
        conversion_tasks = []

        # --- 處理合併模式 ---
        if is_merge:
            filename = custom_name if custom_name else datetime.now().strftime("%Y%m%d_%H%M%S")
            if not filename.lower().endswith(".pdf"): filename += ".pdf"

            target_path = os.path.join(out_dir, filename)

            if os.path.exists(target_path):
                choice = self.prompt_file_exists(filename)
                if choice == "skip":
                    return
                elif choice == "keep":
                    target_path = self.get_unique_filepath(target_path)

            conversion_tasks.append((paths, target_path))

        # --- 處理獨立檔案模式 ---
        else:
            for p in paths:
                base_name = os.path.splitext(os.path.basename(p))[0]
                filename = f"{base_name}.pdf"
                target_path = os.path.join(out_dir, filename)

                if os.path.exists(target_path):
                    choice = self.prompt_file_exists(filename)
                    if choice == "skip":
                        continue  # 跳過這張圖，繼續下一張
                    elif choice == "keep":
                        target_path = self.get_unique_filepath(target_path)

                conversion_tasks.append(([p], target_path))

        if not conversion_tasks:
            return

        # UI 狀態鎖定
        self.btn_convert.configure(state="disabled", text="轉換中...")

        def run_tasks():
            success_count = 0
            all_ok = True

            for task_paths, target_file in conversion_tasks:
                success, _ = convert_images_to_pdf(task_paths, target_file, bg_col)
                if success:
                    success_count += 1
                else:
                    all_ok = False

            # 準備完成訊息
            if is_merge:
                msg = f"檔案已成功儲存於：\n{conversion_tasks[0][1]}" if all_ok else "轉換過程發生錯誤。"
            else:
                msg = f"完成 {success_count}/{len(conversion_tasks)} 張圖片轉換。"

            self.root.after(0, lambda: self.on_conversion_done(is_merge, all_ok, msg))

        threading.Thread(target=run_tasks, daemon=True).start()

    def on_conversion_done(self, is_merge, success, msg):
        if is_merge:
            if success:
                messagebox.showinfo("成功", f"轉換完成！\n\n{msg}")
            else:
                messagebox.showerror("錯誤", msg)
        else:
            if success:
                messagebox.showinfo("成功", msg)
            else:
                messagebox.showwarning("部分成功", msg)

        self.btn_convert.configure(state="normal", text="開始轉換")


# ==========================================
# 【新增】：快速模式 (Drop-and-Go) 邏輯
# ==========================================
def run_fast_mode(file_paths):
    """
    當使用者將檔案直接拖曳到 .exe 圖示上時觸發。
    不開啟主 UI，直接將圖片打包成 PDF，並儲存在程式所在目錄。
    """
    # 建立一個隱藏的 Tkinter 視窗，用來顯示完成後的 MessageBox 提示
    root = tk.Tk()
    root.withdraw()

    try:
        # 1. 過濾出有效的圖片檔案
        valid_exts = ('.png', '.jpg', '.jpeg', '.bmp', '.gif')
        images = [p for p in file_paths if p.lower().endswith(valid_exts)]

        if not images:
            messagebox.showerror("錯誤", "沒有找到支援的圖片格式！")
            return

        # 2. 決定輸出目錄 (取決於是否被 PyInstaller 打包)
        if getattr(sys, 'frozen', False):
            app_dir = os.path.dirname(sys.executable)
        else:
            app_dir = os.path.dirname(os.path.abspath(__file__))

        # 3. 智慧判斷檔名：單檔用原名，多檔用時間
        if len(images) == 1:
            base_name = os.path.splitext(os.path.basename(images[0]))[0]
            out_name = f"{base_name}.pdf"
        else:
            out_name = f"FastConvert_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf"

        out_path = os.path.join(app_dir, out_name)

        # 4. 檢查檔案是否已存在，並詢問是否覆蓋
        if os.path.exists(out_path):
            overwrite = messagebox.askyesno(
                "發現重複檔案",
                f"輸出目錄中已存在名為「{out_name}」的檔案。\n\n是否要覆蓋它？"
            )
            # 如果使用者點擊「否」，則取消轉換並結束
            if not overwrite:
                return

        # 5. 執行轉換 (預設背景色為白色)
        success, msg = convert_images_to_pdf(images, out_path, (255, 255, 255))
        if success:
            messagebox.showinfo("快速轉換成功", f"已成功將 {len(images)} 張圖片轉換為 PDF！\n\n儲存於：\n{out_path}")
        else:
            messagebox.showerror("轉換失敗", msg)

    except Exception as e:
        messagebox.showerror("系統錯誤", str(e))
    finally:
        # 確保隱藏視窗的進程會被乾淨地關閉
        root.destroy()


# ==========================================
# 程式進入點
# ==========================================
if __name__ == "__main__":
    # sys.argv 存放的是系統傳給這支程式的參數
    # sys.argv[0] 是程式自己本身的名字
    # 如果 len(sys.argv) > 1，代表有人拖曳檔案到它的圖示上

    if len(sys.argv) > 1:
        # 提取拖曳進來的檔案路徑 (排除第 0 個自己)
        dropped_files = sys.argv[1:]
        run_fast_mode(dropped_files)
    else:
        # 沒有帶參數，正常啟動 UI
        root = CTk_DnD()
        app = ImageToPdfApp(root)
        root.mainloop()