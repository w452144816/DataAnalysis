import torch
import torch.nn as nn
import torchvision.models as models
from torchvision import transforms
import cv2
import numpy as np
from math import ceil
from itertools import product
import argparse
from enum import Enum
from PIL import Image
import math
from time import perf_counter
import os
import pprint

def conv_bn(inp, oup, stride = 1, leaky = 0):
    return nn.Sequential(
        nn.Conv2d(inp, oup, 3, stride, 1, bias=False),
        nn.BatchNorm2d(oup),
        nn.LeakyReLU(negative_slope=leaky, inplace=True)
    )

def conv_bn_no_relu(inp, oup, stride):
    return nn.Sequential(
        nn.Conv2d(inp, oup, 3, stride, 1, bias=False),
        nn.BatchNorm2d(oup),
    )

def conv_bn1X1(inp, oup, stride, leaky=0):
    return nn.Sequential(
        nn.Conv2d(inp, oup, 1, stride, padding=0, bias=False),
        nn.BatchNorm2d(oup),
        nn.LeakyReLU(negative_slope=leaky, inplace=True)
    )

def conv_dw(inp, oup, stride, leaky=0.1):
    return nn.Sequential(
        nn.Conv2d(inp, inp, 3, stride, 1, groups=inp, bias=False),
        nn.BatchNorm2d(inp),
        nn.LeakyReLU(negative_slope= leaky,inplace=True),

        nn.Conv2d(inp, oup, 1, 1, 0, bias=False),
        nn.BatchNorm2d(oup),
        nn.LeakyReLU(negative_slope= leaky,inplace=True),
    )

class SSH(nn.Module):
    def __init__(self, in_channel, out_channel):
        super(SSH, self).__init__()
        assert out_channel % 4 == 0
        leaky = 0
        if (out_channel <= 64):
            leaky = 0.1
        self.conv3X3 = conv_bn_no_relu(in_channel, out_channel//2, stride=1)

        self.conv5X5_1 = conv_bn(in_channel, out_channel//4, stride=1, leaky = leaky)
        self.conv5X5_2 = conv_bn_no_relu(out_channel//4, out_channel//4, stride=1)

        self.conv7X7_2 = conv_bn(out_channel//4, out_channel//4, stride=1, leaky = leaky)
        self.conv7x7_3 = conv_bn_no_relu(out_channel//4, out_channel//4, stride=1)

    def forward(self, input):
        conv3X3 = self.conv3X3(input)

        conv5X5_1 = self.conv5X5_1(input)
        conv5X5 = self.conv5X5_2(conv5X5_1)

        conv7X7_2 = self.conv7X7_2(conv5X5_1)
        conv7X7 = self.conv7x7_3(conv7X7_2)

        out = torch.cat([conv3X3, conv5X5, conv7X7], dim=1)
        out = torch.nn.functional.relu(out)
        return out

class FPN(nn.Module):
    def __init__(self,in_channels_list,out_channels):
        super(FPN,self).__init__()
        leaky = 0
        if (out_channels <= 64):
            leaky = 0.1
        self.output1 = conv_bn1X1(in_channels_list[0], out_channels, stride = 1, leaky = leaky)
        self.output2 = conv_bn1X1(in_channels_list[1], out_channels, stride = 1, leaky = leaky)
        self.output3 = conv_bn1X1(in_channels_list[2], out_channels, stride = 1, leaky = leaky)

        self.merge1 = conv_bn(out_channels, out_channels, leaky = leaky)
        self.merge2 = conv_bn(out_channels, out_channels, leaky = leaky)

    def forward(self, input):
        # names = list(input.keys())
        input = list(input.values())

        output1 = self.output1(input[0])
        output2 = self.output2(input[1])
        output3 = self.output3(input[2])

        up3 = torch.nn.functional.interpolate(output3, size=[output2.size(2), output2.size(3)], mode="nearest")
        output2 = output2 + up3
        output2 = self.merge2(output2)

        up2 = torch.nn.functional.interpolate(output2, size=[output1.size(2), output1.size(3)], mode="nearest")
        output1 = output1 + up2
        output1 = self.merge1(output1)

        out = [output1, output2, output3]
        return out



class ClassHead(nn.Module):
    def __init__(self,inchannels=512,num_anchors=3):
        super(ClassHead,self).__init__()
        self.num_anchors = num_anchors
        self.conv1x1 = nn.Conv2d(inchannels,self.num_anchors*2,kernel_size=(1,1),stride=1,padding=0)

    def forward(self,x):
        out = self.conv1x1(x)
        out = out.permute(0,2,3,1).contiguous()
        
        return out.view(out.shape[0], -1, 2)

class BboxHead(nn.Module):
    def __init__(self,inchannels=512,num_anchors=3):
        super(BboxHead,self).__init__()
        self.conv1x1 = nn.Conv2d(inchannels,num_anchors*4,kernel_size=(1,1),stride=1,padding=0)

    def forward(self,x):
        out = self.conv1x1(x)
        out = out.permute(0,2,3,1).contiguous()

        return out.view(out.shape[0], -1, 4)

class LandmarkHead(nn.Module):
    def __init__(self,inchannels=512,num_anchors=3):
        super(LandmarkHead,self).__init__()
        self.conv1x1 = nn.Conv2d(inchannels,num_anchors*10,kernel_size=(1,1),stride=1,padding=0)

    def forward(self,x):
        out = self.conv1x1(x)
        out = out.permute(0,2,3,1).contiguous()

        return out.view(out.shape[0], -1, 10)

class RetinaFace(nn.Module):
    def __init__(self):
        super(RetinaFace,self).__init__()
            
        #backbone = models.resnet50(weights=models.ResNet50_Weights.IMAGENET1K_V1)
        backbone = models.resnet50()
        self.body = models._utils.IntermediateLayerGetter(backbone, {'layer2': 1, 'layer3': 2, 'layer4': 3})
        in_channels_stage2 = 256
        in_channels_list = [
            in_channels_stage2 * 2,
            in_channels_stage2 * 4,
            in_channels_stage2 * 8,
        ]
        out_channels = 256
        self.fpn = FPN(in_channels_list,out_channels)
        self.ssh1 = SSH(out_channels, out_channels)
        self.ssh2 = SSH(out_channels, out_channels)
        self.ssh3 = SSH(out_channels, out_channels)

        self.ClassHead = self._make_class_head(fpn_num=3, inchannels=256)
        self.BboxHead = self._make_bbox_head(fpn_num=3, inchannels=256)
        self.LandmarkHead = self._make_landmark_head(fpn_num=3, inchannels=256)

    def _make_class_head(self,fpn_num=3,inchannels=64,anchor_num=2):
        classhead = nn.ModuleList()
        for i in range(fpn_num):
            classhead.append(ClassHead(inchannels,anchor_num))
        return classhead
    
    def _make_bbox_head(self,fpn_num=3,inchannels=64,anchor_num=2):
        bboxhead = nn.ModuleList()
        for i in range(fpn_num):
            bboxhead.append(BboxHead(inchannels,anchor_num))
        return bboxhead

    def _make_landmark_head(self,fpn_num=3,inchannels=64,anchor_num=2):
        landmarkhead = nn.ModuleList()
        for i in range(fpn_num):
            landmarkhead.append(LandmarkHead(inchannels,anchor_num))
        return landmarkhead

    def forward(self,inputs):
        out = self.body(inputs)

        # FPN
        fpn = self.fpn(out)

        # SSH
        feature1 = self.ssh1(fpn[0])
        feature2 = self.ssh2(fpn[1])
        feature3 = self.ssh3(fpn[2])
        features = [feature1, feature2, feature3]

        bbox_regressions = torch.cat([self.BboxHead[i](feature) for i, feature in enumerate(features)], dim=1)
        classifications = torch.cat([self.ClassHead[i](feature) for i, feature in enumerate(features)],dim=1)
        ldm_regressions = torch.cat([self.LandmarkHead[i](feature) for i, feature in enumerate(features)], dim=1)
        
        output = (bbox_regressions, torch.nn.functional.softmax(classifications, dim=-1), ldm_regressions)
        return output

class PriorBox(object):
    def __init__(self, image_size=None):
        super(PriorBox, self).__init__()
        self.min_sizes = [[16, 32], [64, 128], [256, 512]]
        self.steps = [8, 16, 32]
        self.clip = False
        self.image_size = image_size
        self.feature_maps = [[ceil(self.image_size[0]/step), ceil(self.image_size[1]/step)] for step in self.steps]
        self.name = "s"

    def forward(self):
        anchors = []
        for k, f in enumerate(self.feature_maps):
            min_sizes = self.min_sizes[k]
            for i, j in product(range(f[0]), range(f[1])):
                for min_size in min_sizes:
                    s_kx = min_size / self.image_size[1]
                    s_ky = min_size / self.image_size[0]
                    dense_cx = [x * self.steps[k] / self.image_size[1] for x in [j + 0.5]]
                    dense_cy = [y * self.steps[k] / self.image_size[0] for y in [i + 0.5]]
                    for cy, cx in product(dense_cy, dense_cx):
                        anchors += [cx, cy, s_kx, s_ky]

        # back to torch land
        output = torch.Tensor(anchors).view(-1, 4)
        if self.clip:
            output.clamp_(max=1, min=0)
        return output


# Adapted from https://github.com/Hakuyume/chainer-ssd
def decode(loc, priors, variances):
    """Decode locations from predictions using priors to undo
    the encoding we did for offset regression at train time.
    Args:
        loc (tensor): location predictions for loc layers,
            Shape: [num_priors,4]
        priors (tensor): Prior boxes in center-offset form.
            Shape: [num_priors,4].
        variances: (list[float]) Variances of priorboxes
    Return:
        decoded bounding box predictions
    """

    boxes = torch.cat((
        priors[:, :2] + loc[:, :2] * variances[0] * priors[:, 2:],
        priors[:, 2:] * torch.exp(loc[:, 2:] * variances[1])), 1)
    boxes[:, :2] -= boxes[:, 2:] / 2
    boxes[:, 2:] += boxes[:, :2]
    return boxes

def decode_landm(pre, priors, variances):
    """Decode landm from predictions using priors to undo
    the encoding we did for offset regression at train time.
    Args:
        pre (tensor): landm predictions for loc layers,
            Shape: [num_priors,10]
        priors (tensor): Prior boxes in center-offset form.
            Shape: [num_priors,4].
        variances: (list[float]) Variances of priorboxes
    Return:
        decoded landm predictions
    """
    landms = torch.cat((priors[:, :2] + pre[:, :2] * variances[0] * priors[:, 2:],
                        priors[:, :2] + pre[:, 2:4] * variances[0] * priors[:, 2:],
                        priors[:, :2] + pre[:, 4:6] * variances[0] * priors[:, 2:],
                        priors[:, :2] + pre[:, 6:8] * variances[0] * priors[:, 2:],
                        priors[:, :2] + pre[:, 8:10] * variances[0] * priors[:, 2:],
                        ), dim=1)
    return landms

# --------------------------------------------------------
# Fast R-CNN
# Copyright (c) 2015 Microsoft
# Licensed under The MIT License [see LICENSE for details]
# Written by Ross Girshick
# --------------------------------------------------------

import numpy as np

def py_cpu_nms(dets, thresh):
    """Pure Python NMS baseline."""
    x1 = dets[:, 0]
    y1 = dets[:, 1]
    x2 = dets[:, 2]
    y2 = dets[:, 3]
    scores = dets[:, 4]

    areas = (x2 - x1 + 1) * (y2 - y1 + 1)
    order = scores.argsort()[::-1]

    keep = []
    while order.size > 0:
        i = order[0]
        keep.append(i)
        xx1 = np.maximum(x1[i], x1[order[1:]])
        yy1 = np.maximum(y1[i], y1[order[1:]])
        xx2 = np.minimum(x2[i], x2[order[1:]])
        yy2 = np.minimum(y2[i], y2[order[1:]])

        w = np.maximum(0.0, xx2 - xx1 + 1)
        h = np.maximum(0.0, yy2 - yy1 + 1)
        inter = w * h
        ovr = inter / (areas[i] + areas[order[1:]] - inter)

        inds = np.where(ovr <= thresh)[0]
        order = order[inds + 1]

    return keep

def remove_prefix(state_dict, prefix):
    ''' Old style model is stored with all names of parameters sharing common prefix 'module.' '''
    #print('remove prefix \'{}\''.format(prefix))
    f = lambda x: x.split(prefix, 1)[-1] if x.startswith(prefix) else x
    return {f(key): value for key, value in state_dict.items()}

def findEuclideanDistance(source_representation, test_representation):
    euclidean_distance = source_representation - test_representation
    euclidean_distance = np.sum(np.multiply(euclidean_distance, euclidean_distance))
    euclidean_distance = np.sqrt(euclidean_distance)
    return euclidean_distance

#this function copied from the deepface repository: https://github.com/serengil/deepface/blob/master/deepface/commons/functions.py
def alignment_procedure(img, left_eye, right_eye, nose):

    #this function aligns given face in img based on left and right eye coordinates

    #left eye is the eye appearing on the left (right eye of the person)
    #left top point is (0, 0)

    left_eye_x, left_eye_y = left_eye
    right_eye_x, right_eye_y = right_eye

    #-----------------------
    #decide the image is inverse

    center_eyes = (int((left_eye_x + right_eye_x) / 2), int((left_eye_y + right_eye_y) / 2))
    
    if False:

        img = cv2.circle(img, (int(left_eye[0]), int(left_eye[1])), 2, (0, 255, 255), 2)
        img = cv2.circle(img, (int(right_eye[0]), int(right_eye[1])), 2, (255, 0, 0), 2)
        img = cv2.circle(img, center_eyes, 2, (0, 0, 255), 2)
        img = cv2.circle(img, (int(nose[0]), int(nose[1])), 2, (255, 255, 255), 2)

    #-----------------------
    #find rotation direction

    if left_eye_y > right_eye_y:
        point_3rd = (right_eye_x, left_eye_y)
        direction = -1 #rotate same direction to clock
    else:
        point_3rd = (left_eye_x, right_eye_y)
        direction = 1 #rotate inverse direction of clock

    #-----------------------
    #find length of triangle edges

    a = findEuclideanDistance(np.array(left_eye), np.array(point_3rd))
    b = findEuclideanDistance(np.array(right_eye), np.array(point_3rd))
    c = findEuclideanDistance(np.array(right_eye), np.array(left_eye))

    #-----------------------

    #apply cosine rule

    if b != 0 and c != 0: #this multiplication causes division by zero in cos_a calculation

        cos_a = (b*b + c*c - a*a)/(2*b*c)
        
        #PR15: While mathematically cos_a must be within the closed range [-1.0, 1.0], floating point errors would produce cases violating this
        #In fact, we did come across a case where cos_a took the value 1.0000000169176173, which lead to a NaN from the following np.arccos step
        cos_a = min(1.0, max(-1.0, cos_a))
        
        
        angle = np.arccos(cos_a) #angle in radian
        angle = (angle * 180) / math.pi #radian to degree

        #-----------------------
        #rotate base image

        if direction == -1:
            angle = 90 - angle

        img = Image.fromarray(img)
        img = np.array(img.rotate(direction * angle))

    #-----------------------

    return img #return img anyway

class Gender(Enum):
    Male = 1
    Female = 2


class Race(Enum):
    White = 1
    Black = 2
    Latino_Hispanic = 3
    East_Asian = 4
    Southeast_Asian = 5
    Indian = 6
    Middle_Eastern = 7

class Age(Enum):
    AGE_0_2 = 1
    AGE_3_9 = 2
    AGE_10_19 = 3
    AGE_20_29 = 4
    AGE_30_39 = 5
    AGE_40_49 = 6
    AGE_50_59 = 7
    AGE_60_69 = 8
    AGE_70_plus = 9

class NeDeepFace:
    def __init__(self, threshold = 0.95):
        pretrained_path = 'models/face/Retinaface_Resnet50_Final.pth'
        self.detector = RetinaFace()
        pretrained_dict = torch.load(pretrained_path, map_location=lambda storage, loc: storage)
        pretrained_dict = remove_prefix(pretrained_dict, 'module.')
        self.detector.load_state_dict(
            pretrained_dict, 
            strict=False)
        self.detector.eval()
        self.threshold = threshold
        self.top_k = 5000
        self.nms_threshold = 0.4
        self.keep_top_k = 750
        self.visualization_threshold = 0.6
        self.max_side = 1280
        self.fair_7 = models.resnet34()
        self.fair_7.fc = nn.Linear(self.fair_7.fc.in_features, 18)
        self.fair_7.load_state_dict(torch.load('models/face/res34_fair_align_multi_7_20190809.pt', map_location=torch.device('cpu')))
        self.fair_7.eval()

    def __check_max_side(self, img):
        max_len = max(img.shape[0], img.shape[1])
        if max_len > self.max_side:
            scale = self.max_side / max_len            
            return cv2.resize(img, (0,0), fx=scale, fy=scale)
        return img

    def __draw_face(self, img, bbox):
        text = "{:.4f}".format(bbox[4])
        b = list(map(int, bbox))
        cv2.rectangle(img, (b[0], b[1]), (b[2], b[3]), (0, 0, 255), 2)
        cx = b[0]
        cy = b[1] + 12
        cv2.putText(img, text, (cx, cy),
                    cv2.FONT_HERSHEY_DUPLEX, 0.5, (255, 255, 255))
        # landms
        cv2.circle(img, (b[5], b[6]), 1, (0, 0, 255), 4)
        cv2.circle(img, (b[7], b[8]), 1, (0, 255, 255), 4)
        cv2.circle(img, (b[9], b[10]), 1, (255, 0, 255), 4)
        cv2.circle(img, (b[11], b[12]), 1, (0, 255, 0), 4)
        cv2.circle(img, (b[13], b[14]), 1, (255, 0, 0), 4)

    def detect_face(self, image_path, draw = False):
        torch.set_grad_enabled(False)
        resize = 1  
        img_raw = None
        if isinstance(image_path, np.ndarray):
            img_raw = image_path
        else:
            img_raw = cv2.imread(image_path)
            img_raw = self.__check_max_side(img_raw)
        img = np.float32(img_raw)

        im_height, im_width, _ = img.shape
        scale = torch.Tensor([img.shape[1], img.shape[0], img.shape[1], img.shape[0]])
        img -= (104, 117, 123)
        img = img.transpose(2, 0, 1)
        img = torch.from_numpy(img).unsqueeze(0)
        
        loc, conf, landms = self.detector(img)  # forward pass

        priorbox = PriorBox(image_size=(im_height, im_width))
        priors = priorbox.forward()        
        prior_data = priors.data
        variance = [0.1, 0.2]
        boxes = decode(loc.data.squeeze(0), prior_data, variance)
        boxes = boxes * scale / resize
        boxes = boxes.cpu().numpy()
        scores = conf.squeeze(0).data.cpu().numpy()[:, 1]
        landms = decode_landm(landms.data.squeeze(0), prior_data, variance)
        scale1 = torch.Tensor([img.shape[3], img.shape[2], img.shape[3], img.shape[2],
                                img.shape[3], img.shape[2], img.shape[3], img.shape[2],
                                img.shape[3], img.shape[2]])
        
        landms = landms * scale1 / resize
        landms = landms.cpu().numpy()

        # ignore low scores
        inds = np.where(scores > self.threshold)[0]
        boxes = boxes[inds]
        landms = landms[inds]
        scores = scores[inds]

        # keep top-K before NMS
        order = scores.argsort()[::-1][:self.top_k]
        boxes = boxes[order]
        landms = landms[order]
        scores = scores[order]

        # do NMS
        dets = np.hstack((boxes, scores[:, np.newaxis])).astype(np.float32, copy=False)
        keep = py_cpu_nms(dets, self.nms_threshold)
        # keep = nms(dets, args.nms_threshold,force_cpu=args.cpu)
        dets = dets[keep, :]
        landms = landms[keep]

        # keep top-K faster NMS
        dets = dets[:self.keep_top_k, :]
        landms = landms[:self.keep_top_k, :]

        dets = np.concatenate((dets, landms), axis=1)

        #print(dets.shape)
        # show image
        vis_cnt = 0
        target = []
        for b in dets:
            if b[4] < self.visualization_threshold:
                continue
            vis_cnt += 1

            temp = [int(item) for item in b]
            target.append(temp[:4])

        if draw:
            for b in dets:
                if b[4] < self.visualization_threshold:
                    continue
                self.__draw_face(img_raw, b)                

            cv2.imshow('result', img_raw)
            cv2.waitKey(0)
        if vis_cnt == 1:
            return target
        return None

    def detect_age_gender_race(self, img_path, draw = False):
        start = perf_counter()
        # img_raw = cv2.imread(img_path)
        img_raw = np.array(img_path)
        img_raw = self.__check_max_side(img_raw)
        bbox = self.detect_face(img_raw, False)
        bbox = bbox[0] if bbox is not None else None
        if not bbox or not all(x >= 0 for x in bbox):
            return None
        img_face = img_raw[round(bbox[1]) : round(bbox[3]), round(bbox[0]) : round(bbox[2])]
        img = cv2.cvtColor(img_face, cv2.COLOR_BGR2RGB)
        trans = transforms.Compose([
            transforms.ToPILImage(),
            transforms.Resize((224, 224)),
            transforms.ToTensor(),
            transforms.Normalize(mean=[0.485, 0.456, 0.406], std=[0.229, 0.224, 0.225])
        ])

        image = trans(img)
        image = image.view(1, 3, 224, 224)  # reshape image to match model dimensions (1 batch size)
            
        outputs = self.fair_7(image)
        outputs = outputs.detach().numpy()
        outputs = np.squeeze(outputs)

        race_outputs = outputs[:7]
        gender_outputs = outputs[7:9]
        age_outputs = outputs[9:18]

        race_score = np.exp(race_outputs) / np.sum(np.exp(race_outputs))
        gender_score = np.exp(gender_outputs) / np.sum(np.exp(gender_outputs))
        age_score = np.exp(age_outputs) / np.sum(np.exp(age_outputs))

        race_pred = np.argmax(race_score)
        gender_pred = np.argmax(gender_score)
        age_pred = np.argmax(age_score)
        end = perf_counter()

        result = {}
        result['gender'] = Gender(gender_pred + 1).name
        result['race'] = Race(race_pred + 1).name
        result['Age'] = Age(age_pred + 1).name
        # if draw:
        #     pprint.pprint(result)
        #     print('Time spent {:.2f} s'.format(end - start))
        #     self.__draw_face(img_raw, bbox)
        #     # print(Race(race_pred + 1))
        #     # print(Gender(gender_pred + 1))
        #     # print(Age(age_pred + 1))
        #     x0 = round(bbox[0])
        #     y0 = round(bbox[3] + 30)
        #     cv2.putText(img_raw, str(Race(race_pred + 1)), (x0, y0),
        #             cv2.FONT_HERSHEY_DUPLEX, 1.0, (0, 0, 255))
        #     y0 += 30
        #     cv2.putText(img_raw, str(Gender(gender_pred + 1)), (x0, y0),
        #             cv2.FONT_HERSHEY_DUPLEX, 1.0, (0, 0, 255))
        #     y0 += 30
        #     cv2.putText(img_raw, str(Age(age_pred + 1)), (x0, y0),
        #             cv2.FONT_HERSHEY_DUPLEX, 1.0, (0, 0, 255))
        #     cv2.imshow('result', img_raw)
        #     cv2.waitKey(0)
        #     cv2.destroyAllWindows()
        return result

def det_sextype(_pathl:list):
    ndf = NeDeepFace()
    result = {"Male": 0, "Female": 0, "Unknow": 0}
    samplefile = None
    for root, dirs, files in os.walk(_pathl):
        for file_name in files:
            tfile_name = os.path.join(root, file_name)
            if os.path.isfile(tfile_name) == True:
                try:
                    res = ndf.detect_age_gender_race(tfile_name, True)
                    if res:
                        if samplefile is None:
                            samplefile = tfile_name
                        res_gender = res['gender']
                        if res_gender == "Male":
                            result["Male"] = result["Male"] + 1
                        elif res_gender == "Female":
                            result["Female"] = result["Female"] + 1
                    else:
                        result["Unknow"] = result["Unknow"] + 1
                except:
                    pass
            else:
                print("not a file")
            print('')
            shared.state.nextjob()

    r = max(result.keys(), key=(lambda x: result[x]))
    return r, samplefile

def loop_test():    
    ndf = NeDeepFace()
    while True:
        print('Please input file name, or q to exit:')
        file_name = input()
        if file_name == 'q':
            break
        
        file_name = file_name.strip('"')

        if os.path.isfile(file_name) == True:
            ndf.detect_age_gender_race(file_name, True)
        else:
            print("not a file")        
        print('')

if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='NeDeepFace')
    parser.add_argument('-i', '--input', help='input image path')
    parser.add_argument('--gender', action='store_true', default=True)
    parser.add_argument('--loop', action='store_true')
    args = parser.parse_args()
    
    if args.loop:
        loop_test()
    else:    
        ndf = NeDeepFace()
        if args.gender:
            ndf.detect_age_gender_race(args.input, True)
        else:
            ndf.detect_face(args.input, True)
    