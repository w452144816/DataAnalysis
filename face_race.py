import os
import pandas as pd
import numpy as np
from PIL import Image
import torch
from torch.utils.data import Dataset, DataLoader
from torchvision import transforms
from sklearn.model_selection import train_test_split
from sklearn.metrics import accuracy_score, classification_report

from deepface import DeepFace

# 自定义数据集类
class FairFaceDataset(Dataset):
    def __init__(self, csv_file, img_dir, transform=None):
        """
        Args:
            csv_file (str): 数据集的 CSV 文件路径。
            img_dir (str): 图像文件夹路径。
            transform (callable, optional): 图像预处理操作。
        """
        self.data = pd.read_csv(csv_file)
        self.img_dir = img_dir
        self.transform = transform

    def __len__(self):
        return len(self.data)

    def __getitem__(self, idx):
        if torch.is_tensor(idx):
            idx = idx.tolist()

        img_name = os.path.join(self.img_dir, self.data.iloc[idx, 0])
        image = Image.open(img_name).convert("RGB")
        label = self.data.iloc[idx, 3]  # 种族标签在第 7 列

        if self.transform:
            image = self.transform(image)

        # 将标签转换为类别索引
        label_map = {
            "White": 0,
            "Black": 1,
            "Latino_Hispanic": 2,
            "East Asian": 3,
            "Southeast Asian": 4,
            "Indian": 5,
            "Middle Eastern": 6,
        }
        label = label_map[label]

        return image, label, img_name

# 数据预处理
transform = transforms.Compose([
    transforms.Resize((224, 224)),
    transforms.ToTensor(),
    transforms.Normalize(mean=[0.485, 0.456, 0.406], std=[0.229, 0.224, 0.225]),
])

# 加载数据集
csv_file = "../data/face/race/fairface/fairface_label_train.csv"  # 验证集的 CSV 文件路径
img_dir = "../data/face/race/fairface/fairface-img-margin025-trainval"  # 图像文件夹路径

dataset = FairFaceDataset(csv_file=csv_file, img_dir=img_dir, transform=transform)
dataloader = DataLoader(dataset, batch_size=32, shuffle=False)

# 加载预训练模型（例如 ResNet）
from torchvision import models

fair_7 = models.resnet34()
fair_7.fc = torch.nn.Linear(fair_7.fc.in_features, 18)
fair_7.load_state_dict(
    torch.load('../aiface-engine/models/res34_fair_align_multi_7_20190809.pt', map_location=torch.device('cpu')))
fair_7.eval()

df_label_map = {
    "asian": 3,
    "indian": 5,
    "black": 1,
    "white": 0,
    "middle eastern": 6,
    "latino hispanic": 2
}

# 验证模型
all_preds = []
all_labels = []

with torch.no_grad():
    for images, labels, imagepath in dataloader:
        outputs = fair_7(images)
        outputs = outputs.detach().numpy()
        outputs = np.squeeze(outputs)
        race_outputs = outputs[:,:7]
        race_score = np.exp(race_outputs) / np.sum(np.exp(race_outputs))
        race_pred = np.argmax(race_score, axis=1)

        # df_res_l = []
        # for it in imagepath:
        #     objs = DeepFace.analyze(
        #         img_path=it,
        #         actions=['race'],
        #         enforce_detection = False
        #     )
        #     df_res = df_label_map[objs[0]['dominant_race']]
        #     df_res_l.append(df_res)

        # _, preds = torch.max(outputs, 1)
        all_preds.extend(race_pred)
        all_labels.extend(labels.cpu().numpy())

# 计算准确率和分类报告
accuracy = accuracy_score(all_labels, all_preds)
report = classification_report(all_labels, all_preds, target_names=[
    "White", "Black", "Latino_Hispanic", "East Asian", "Southeast Asian", "Indian", "Middle Eastern"
])

print(f"Accuracy: {accuracy:.4f}")
print("Classification Report:")
print(report)


#Accuracy: 0.7199  20ms
# Classification Report:
#                  precision    recall  f1-score   support
#
#           White       0.77      0.77      0.77      2085
#           Black       0.88      0.86      0.87      1556
# Latino_Hispanic       0.56      0.58      0.57      1623
#      East Asian       0.74      0.78      0.76      1550
# Southeast Asian       0.65      0.64      0.65      1415
#          Indian       0.77      0.72      0.75      1516
#  Middle Eastern       0.65      0.64      0.64      1209
#
#        accuracy                           0.72     10954
#       macro avg       0.72      0.71      0.72     10954
#    weighted avg       0.72      0.72      0.72     10954


# Accuracy: 0.4857  100ms
# Classification Report:
#                  precision    recall  f1-score   support
#
#           White       0.50      0.69      0.58      2085
#           Black       0.66      0.79      0.72      1556
# Latino_Hispanic       0.37      0.29      0.32      1623
#      East Asian       0.39      0.85      0.53      1550
# Southeast Asian       0.00      0.00      0.00      1415
#          Indian       0.70      0.29      0.41      1516
#  Middle Eastern       0.46      0.35      0.40      1209
#
#        accuracy                           0.49     10954
#       macro avg       0.44      0.47      0.42     10954
#    weighted avg       0.45      0.49      0.44     10954