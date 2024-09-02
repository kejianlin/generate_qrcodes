import pandas as pd
import qrcode
from PIL import Image, ImageDraw, ImageFont
import os

# 设置 Excel 文件路径和读取数据
excel_path = "chargers.xlsx"  # 修改为你的 Excel 文件路径
sheet_name = "deviceNumber"   # 修改为你的工作表名称
column_name = "设备编号"        # 修改为包含编号的列名称

# 读取 Excel 文件
df = pd.read_excel(excel_path, sheet_name=sheet_name)
charger_numbers = df[column_name].dropna().tolist()

# 设置保存目录
save_directory = "qr_codes_bmp"
os.makedirs(save_directory, exist_ok=True)

# 遍历内容生成二维码并保存为 BMP 格式
for index, charger_code in enumerate(charger_numbers):

    # 提取充电桩编号
    charger_code = f"{charger_code}-1"
    charger_url = f"https://api.shangyucharge.com/1731594137129717760?code={charger_code}"
    
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
    img = qr.make_image(fill_color="black", back_color="white")
    
    # 转换图像为RGB模式，方便后续处理
    img = img.convert("RGB")
    
    # 设置字体（可以使用粗体字体，例如 arialbd.ttf，如果存在）
    try:
        font = ImageFont.truetype("arialbd.ttf", 40)  # 尝试使用粗体字体
    except IOError:
        font = ImageFont.truetype("arial.ttf", 40)  # 退回到常规字体
        # 或者使用默认字体作为备选
        # font = ImageFont.load_default()

    # 计算文本尺寸
    charger_code = str(charger_code)  # 确保将其转换为字符串
    bbox = font.getbbox(charger_code)
    text_width = bbox[2] - bbox[0]
    text_height = bbox[3] - bbox[1]

    # 创建一个新的图像，将二维码和文本组合
    combined_img = Image.new("RGB", (img.width, img.height + text_height + 10), "white")
    combined_img.paste(img, (0, 0))

    # 在底部绘制文本
    draw_combined = ImageDraw.Draw(combined_img)
    text_position = ((img.width - text_width) // 2, img.height - 20)
    draw_combined.text(text_position, charger_code, font=font, fill="black")

    # 设置保存文件名
    file_name = f"qrcode_{charger_code}.bmp"
    file_path = os.path.join(save_directory, file_name)

    # 保存为 BMP 格式
    combined_img.save(file_path, format='BMP')

    print(f"二维码已保存到: {file_path}")

print("批量二维码生成完成！")
