import sys
import win32com.client
import os
import cv2
import webbrowser
import winshell
import numpy as np
from PIL import ImageGrab
from selenium import webdriver
from win32api import GetSystemMetrics
from datetime import datetime, timedelta, time
from time import sleep

FILENAME_CAM = "output"
EXTENSION_CAM = ".mp4"
FULL_PATH_CAM = FILENAME_CAM + EXTENSION_CAM

FILENAME_SCR = "scrCap"
EXTENSION_SCR = ".mp4"
FULL_PATH_SCR = FILENAME_SCR + EXTENSION_SCR

FILENAME_COMB = "combined"
EXTENSION_COMB = ".mp4"
FULL_PATH_COMB = FILENAME_COMB + EXTENSION_COMB

FPS = 10
RESOLUTION = "480p"

DESKTOP_PATH = winshell.desktop(common=False)
MAX_MINS = 1

def is_time_eq(start, end):
    start_h = start.hour
    start_m = start.minute
    start_s = start.second
    end_h = end.hour
    end_m = end.minute
    end_s = end.second
    return time(start_h, start_m, start_s) == time(end_h, end_m, end_s)

def get_monitor_res():
    return GetSystemMetrics(0), GetSystemMetrics(1)

def main():  
    current_time = datetime.now()
    added_time = current_time + timedelta(minutes=MAX_MINS)

    # video capture from camera
    cap = cv2.VideoCapture(-1)

    # needed dimensions
    scr_w, scr_h = get_monitor_res()
    w = int(cap.get(3))
    h = int(cap.get(4))

    # codec
    fourcc = cv2.VideoWriter_fourcc(*'XVID')
    
    out_cam = cv2.VideoWriter(FULL_PATH_CAM, fourcc, FPS, (w, h))
    out_scr = cv2.VideoWriter(FULL_PATH_SCR, fourcc, FPS, (scr_w, scr_h))

    writer = None

    while cap.isOpened():
        img = ImageGrab.grab()
        img_np = np.array(img)
        screen_frame = cv2.cvtColor(img_np, cv2.COLOR_BGR2RGB)
        
        # Capture frame-by-frame
        ret, cam_frame = cap.read()
        if not ret:
            break

        writer = None
       # screen_frame = resize(screen_frame, width=240, height=320)
        screen_frame_resized = cv2.resize(screen_frame, dsize=(w, h), interpolation=cv2.INTER_CUBIC)

        if writer is None:
            writer = cv2.VideoWriter(FULL_PATH_COMB, fourcc, FPS, (w * 2, h))

            #show frame in window
            cv2.imshow("webcam cap", cam_frame)

            # write frame to output file
            out_cam.write(cam_frame)
            out_scr.write(screen_frame)

            (B_C, G_C, R_C) = cv2.split(cam_frame)
            (B_S, G_S, R_S) = cv2.split(screen_frame_resized)
            cam_frame_merged = cv2.merge([B_C, G_C, R_C])
            screen_frame_merged = cv2.merge([B_S, G_S, R_S])
            
            final_output = np.zeros((h, w * 2, 3), dtype="uint8")
            final_output[0:h, 0:w] = cam_frame_merged
            final_output[0:h, w:w*2] = screen_frame_merged

            writer.write(final_output)

            key = cv2.waitKey(1) & 0xFF
            #check for keypress (q to quit)
            if key == ord('q'):
                break

    # When everything done, release the capture
    cap.release()
    out_cam.release()
    out_scr.release()
    writer.release()
    cv2.destroyAllWindows()

if __name__ == "__main__":
        main()
