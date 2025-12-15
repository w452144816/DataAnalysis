import os
import cv2

import NeDeepFace

# 源目录（包含图片的文件夹）
source_dir = '../data/face/race/avatar'

# 目标目录（用于存放检测到的人脸区域图片）
target_dir = '../data/face/race/avatar_race'

# 确保目标目录存在，如果不存在则创建
if not os.path.exists(target_dir):
    os.makedirs(target_dir)

# 加载 OpenCV 的人脸检测模型（Haar Cascade 或 DNN）
# 这里使用 Haar Cascade 模型
# face_cascade = cv2.CascadeClassifier(cv2.data.haarcascades + 'haarcascade_frontalface_default.xml')

ndf = NeDeepFace.NeDeepFace()

# 图片文件的扩展名
image_extensions = ['.jpg', '.jpeg', '.png', '.bmp']

# 遍历源目录及其所有子目录
for root, dirs, files in os.walk(source_dir):
    for file in files:
        # 获取文件的扩展名
        ext = os.path.splitext(file)[1].lower()

        # 检查文件是否是图片
        if ext in image_extensions:
            # 获取文件的完整路径
            image_path = os.path.join(root, file)

            # 读取图片
            image = cv2.imread(image_path)

            # 如果图片读取失败，跳过
            if image is None:
                print(f"无法读取图片: {image_path}")
                continue

            # # 将图片转换为灰度图（人脸检测需要灰度图）
            # gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)

            # 检测人脸
            # faces = face_cascade.detectMultiScale(gray, scaleFactor=1.1, minNeighbors=5, minSize=(30, 30))

            # faces = ndf.detect_face(image)
            race = ndf.detect_age_gender_race(image)
            if race is None:
                continue

            # 生成目标文件名
            filename = os.path.splitext(file)[0]
            face_filename = f"{filename}_face_{1}.jpg"
            target_path = os.path.join(target_dir, face_filename)
            cv2.putText(image, race['race'], (0,150), fontFace = cv2.FONT_HERSHEY_SIMPLEX, fontScale = 6, color = (129, 0, 255), thickness = 4)
            # 保存人脸区域图片
            try:
                cv2.imwrite(target_path, image)
                print(f"保存人脸区域: {target_path}")
            except:
                print(f"保存人脸失败: {target_path}")
                pass

            # 遍历检测到的人脸
            # for i, (x, y, w, h) in enumerate(faces):
            #     # 裁剪人脸区域
            #     face_roi = image[y:h, x:w]
            #
            #     # 生成目标文件名
            #     filename = os.path.splitext(file)[0]
            #     face_filename = f"{filename}_face_{i + 1}.jpg"
            #     target_path = os.path.join(target_dir, face_filename)
            #
            #     # 保存人脸区域图片
            #     try:
            #         cv2.imwrite(target_path, face_roi)
            #         print(f"保存人脸区域: {target_path}")
            #     except:
            #         print(f"保存人脸失败: {target_path}")
            #         pass

print("所有图片的人脸区域已保存到目标文件夹。")