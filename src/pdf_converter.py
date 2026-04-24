# pdf_converter.py
import os
from PIL import Image


def convert_images_to_pdf(image_paths, output_pdf_path, bg_color=(255, 255, 255)):
    """
    將多張圖片轉換並合併為單一 PDF 檔案。
    """
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