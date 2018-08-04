import sys
import os
import cv2
import numpy as np
from PIL import ImageGrab

class Capture:
    def __init__(self, cam_id, save_path, size=[]):
        self.save_path = save_path
        self.size = size
        self.width = size[0]
        self.height = size[1]
        self.fps = 20
        self.capture = cv2.VideoCapture(cam_id)
        self.codec = cv2.VideoWriter_fourcc(*"XVID")        

    def get_writer(self):
        return cv2.VideoWriter(self.save_path, self.codec, self.fps, (self.width * 2, self.height))

    def get_output(self, first_frame, second_frame):
        out = np.zeros((self.height, self.width * 2, 3), dtype="uint8")
        out[0:self.height, 0:self.width] = first_frame
        out[0:self.height, self.width:self.width * 2] = second_frame
        return out
            
    def resize(self, frame):
        w = self.size[0]
        h = self.size[1]
        return cv2.resize(frame, dsize=(w, h), interpolation=cv2.INTER_CUBIC)

    def split_rgb(self, frame):
        return cv2.split(frame)

    def merge_frame(self, frame):
        (B, G, R) = self.split_rgb(frame)
        return cv2.merge([B, G, R])




