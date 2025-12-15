import os
import glob
import shutil
from datetime import datetime

# 源目录（包含图片的文件夹）
source_dir = 'C:/Users/NERO/Downloads/temp'

# 目标目录（新文件夹，用于存放复制的图片）
target_dir = '../data/face/race/avatar'

# 确保目标目录存在，如果不存在则创建
if not os.path.exists(target_dir):
    os.makedirs(target_dir)

# 使用 glob 查找所有图片文件（例如：jpg, png, jpeg）
image_extensions = ['.jpg', '.jpeg', '.png', '.gif', '.bmp']

image_files = []
# 遍历源目录及其子目录
for root, dirs, files in os.walk(source_dir):
    print(f"Current directory: {root}")
    print(f"Files in this directory: {files}")
    for file in files:
        # 获取文件的扩展名
        ext = os.path.splitext(file)[1].lower()

        # 检查文件是否是图片
        if ext in image_extensions:
            # 获取文件的完整路径
            image_path = os.path.join(root, file)
            image_files.append(image_path)

# 按文件名排序（可以根据需要按其他方式排序）
image_files.sort()

# 按顺序重命名并复制图片
for index, image_path in enumerate(image_files, start=1):
    # 获取文件名和扩展名
    filename = os.path.basename(image_path)
    name, ext = os.path.splitext(filename)

    # 生成新的文件名（例如：image_001.jpg, image_002.jpg, ...）
    new_filename = f"image_{index:03d}{ext}"

    # 构建目标路径
    target_path = os.path.join(target_dir, new_filename)

    # 复制文件到目标目录
    shutil.copy(image_path, target_path)
    print(f"Copied {image_path} to {target_path}")

print("所有图片已复制到目标文件夹。")