import tkinter as tk
from tkinter import filedialog, messagebox
import qrcode
from PIL import Image, ImageDraw, ImageFont
import os
import zipfile

# 设置保存目录并生成二维码
def generate_qr_codes(start_code, end_code, save_directory, prefix_url):
    try:
        # 确保编号范围是整数
        start_code = int(start_code)
        end_code = int(end_code)
    except ValueError:
        messagebox.showerror("输入错误", "编号范围必须是整数！")
        return
    
    os.makedirs(save_directory, exist_ok=True)

    qr_files = []  # 用于存储所有生成的二维码文件路径

    for code in range(start_code, end_code + 1):
        charger_code = f"{code}-1"
        charger_url = f"{prefix_url}{charger_code}"

        # 生成二维码
        qr = qrcode.QRCode(
            version=1,
            error_correction=qrcode.constants.ERROR_CORRECT_L,
            box_size=10,
            border=4,
        )
        qr.add_data(charger_url)
        qr.make(fit=True)

        # 创建二维码图像
        img = qr.make_image(fill_color="black", back_color="white").convert("RGB")

        # 设置字体
        try:
            font = ImageFont.truetype("arialbd.ttf", 40)
        except IOError:
            font = ImageFont.truetype("arial.ttf", 40)

        # 计算文本尺寸
        bbox = font.getbbox(charger_code)
        text_width = bbox[2] - bbox[0]
        text_height = bbox[3] - bbox[1]

        # 创建组合图像
        combined_img = Image.new("RGB", (img.width, img.height + text_height + 10), "white")
        combined_img.paste(img, (0, 0))

        # 绘制文本
        draw_combined = ImageDraw.Draw(combined_img)
        text_position = ((img.width - text_width) // 2, img.height - 20)
        draw_combined.text(text_position, charger_code, font=font, fill="black")

        # 保存二维码
        file_name = f"qrcode_{charger_code}.bmp"
        file_path = os.path.join(save_directory, file_name)
        combined_img.save(file_path, format='BMP')

        qr_files.append(file_path)  # 将文件路径添加到列表

    # 使用起始编号创建压缩包名称
    zip_file_name = f"qrcodes_{start_code}.zip"
    zip_file_path = os.path.join(save_directory, zip_file_name)

    # 将二维码文件打包成 zip 文件
    with zipfile.ZipFile(zip_file_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
        for file in qr_files:
            zipf.write(file, os.path.basename(file))  # 将二维码文件写入 zip 压缩包中

    # 删除原始的二维码文件（如果不再需要）
    for file in qr_files:
        os.remove(file)

    messagebox.showinfo("完成", f"二维码已生成并打包成 ZIP 文件，保存到 {zip_file_path}！")

# 创建 GUI
def start_gui():
    def on_generate():
        start_code = entry_start.get()
        end_code = entry_end.get()
        prefix_url = entry_prefix.get().strip()
        if not prefix_url:
            prefix_url = "https://api.shangyucharge.com/1731594137129717760?code="
        save_directory = filedialog.askdirectory(title="选择保存目录")
        if not save_directory:
            return
        generate_qr_codes(start_code, end_code, save_directory, prefix_url)

    root = tk.Tk()
    root.title("商宇充电桩二维码生成器")

    tk.Label(root, text="起始桩编号:").grid(row=0, column=0, padx=10, pady=5)
    entry_start = tk.Entry(root, width=25)
    entry_start.grid(row=0, column=1, padx=10, pady=5)

    tk.Label(root, text="结束桩编号:").grid(row=1, column=0, padx=10, pady=5)
    entry_end = tk.Entry(root, width=25)
    entry_end.grid(row=1, column=1, padx=10, pady=5)

    tk.Label(root, text="前缀 URL:").grid(row=2, column=0, padx=10, pady=5)
    entry_prefix = tk.Entry(root, width=25)
    entry_prefix.insert(0, "https://api.shangyucharge.com/1731594137129717760?code=")
    entry_prefix.grid(row=2, column=1, padx=10, pady=5)

    tk.Button(root, text="生成二维码", command=on_generate).grid(row=3, column=0, columnspan=2, pady=10)

    root.mainloop()

if __name__ == "__main__":
    start_gui()
