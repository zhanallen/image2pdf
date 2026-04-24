import os
from PIL import Image


def convert_images_to_pdf(image_paths, output_pdf_path, bg_color=(255, 255, 255)):
    """
    核心模組：將多張圖片轉換並合併為單一 PDF 檔案。支援自訂透明背景顏色。

    參數:
        image_paths (list): 包含圖片檔案路徑的字串列表。
        output_pdf_path (str): 欲輸出的 PDF 檔案名稱或路徑。
        bg_color (tuple): 替換透明通道的 RGB 顏色，預設為白色 (255, 255, 255)。

    回傳:
        bool: 轉換成功回傳 True，失敗回傳 False。
    """
    if not image_paths:
        print("錯誤：沒有提供任何圖片路徑。")
        return False

    image_list = []
    first_image = None

    try:
        for path in image_paths:
            if not os.path.exists(path):
                print(f"警告：找不到圖片檔案 '{path}'，將略過此檔案。")
                continue

            # 打開圖片
            img = Image.open(path)

            # 判斷圖片是否包含透明通道 (RGBA, LA 或是帶有透明度資訊的 P 模式)
            if img.mode in ('RGBA', 'LA') or (img.mode == 'P' and 'transparency' in img.info):
                # 統一轉換為 RGBA 模式以便處理
                img_rgba = img.convert('RGBA')

                # 建立一張指定 RGB 顏色的純色背景圖
                background = Image.new('RGB', img_rgba.size, bg_color)

                # 將原圖貼上背景，並使用原圖的 alpha 通道 (第 3 個 channel) 作為遮罩
                # 這樣透明的地方就會透出我們設定的背景色
                background.paste(img_rgba, mask=img_rgba.split()[3])
                img = background
            else:
                # 如果沒有透明通道，直接轉換為 RGB 即可
                img = img.convert('RGB')

            if first_image is None:
                first_image = img
            else:
                image_list.append(img)

        # 檢查是否至少有一張有效的圖片
        if first_image is None:
            print("錯誤：沒有找到任何有效的圖片可以轉換。")
            return False

        # 將第一張圖片儲存為 PDF，並將後續圖片附加進去
        first_image.save(
            output_pdf_path,
            "PDF",
            resolution=100.0,
            save_all=True,
            append_images=image_list
        )
        print(f"✅ 成功！PDF 已儲存至：{output_pdf_path}")
        return True

    except Exception as e:
        print(f"❌ 轉換過程中發生未預期的錯誤：{e}")
        return False


def main():
    """
    主程式：負責處理流程控制與使用者互動。
    """
    print("=== 圖片轉 PDF 小工具 ===")

    input_images = [
        "../test/images/sample1.jpg",
        "../test/images/sample2.jpg",
        "../test/images/sample16.png",
        "../test/images/test.png"
    ]
    output_file = "../test/output.pdf"

    print(f"準備處理 {len(input_images)} 張圖片...")

    # 呼叫獨立的核心模組進行轉換
    # 這裡我們將透明背景設定為淡黃色 (255, 255, 204) 作為測試
    # 如果不傳入 bg_color，就會使用預設的純白色 (255, 255, 255)
    success = convert_images_to_pdf(input_images, output_file, bg_color=(255, 255, 255))

    if success:
        print("程式執行完畢。")
    else:
        print("程式無法完成轉換，請檢查上述錯誤訊息。")


if __name__ == "__main__":
    main()