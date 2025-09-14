import subprocess
import os
from datetime import datetime

# ---------- 输入 PNG 文件 ----------
input_png = r"D:\a数据库\有意思的东西\基因工具\image\ico (1).svg"

# ---------- 输出 ICO 文件 ----------
base_name = os.path.splitext(os.path.basename(input_png))[0]
timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
output_ico = os.path.join(os.path.dirname(input_png), f"{base_name}_{timestamp}.ico")

# ---------- ICO 尺寸 ----------
sizes = [256, 128, 64, 48, 32, 16]
size_str = ",".join(map(str, sizes))

# ---------- 构建命令 ----------
magick_path = r"D:\Program Files\ImageMagick-7.1.2-Q16-HDRI\magick.exe"

cmd = [
    magick_path,
    input_png,
    "-define", f"icon:auto-resize={size_str}",
    output_ico
]
subprocess.run(cmd, check=True)

# ---------- 执行 ----------
try:
    subprocess.run(cmd, check=True)
    print(f"成功生成透明 ICO: {output_ico}")
except subprocess.CalledProcessError as e:
    print("生成 ICO 出错:", e)
