import sys
import win32com.client
import os
import cv2
import webbrowser
import winshell
import numpy as np
import psutil
from time import time
from EmailSender import EmailSender
from PIL import ImageGrab
from selenium import webdriver
from win32api import GetSystemMetrics
from datetime import datetime, timedelta

WANTED_PROCESSES = ["fortniteclient-win64-shipping", "chrome"]

FPS = 20

DATE_FILE_FORMAT = "%d-%m-%Y %H.%M"

RESOLUTION = "480p"
FILE_EXT = ".mp4"

DESKTOP_PATH = winshell.desktop(common=False)
SETTINGS_FILE = DESKTOP_PATH + "//settings.txt"

SET_DAY = 23

def open_chrome(url="www.google.com", path="C://Program Files (x86)//Google//Chrome//Application//chrome.exe"):
    webbrowser.get(path).open(url)

def process_running(process_name, extension=".exe"):
    full_process_name = (process_name + extension).lower()
    for proc in psutil.process_iter():
        proc_name = proc.name().lower()
        if proc_name == full_process_name:
            return True
    return False

def make_directory(desktop, dt):
    start_directory = desktop + "\\Captures"
    new_dir = start_directory + "\\" + dt
    if not os.path.exists(new_dir):
        os.makedirs(new_dir)
    return new_dir

def get_settings(filename=SETTINGS_FILE):
    with open(filename, "r") as settings_file:
        for settings in settings_file:
            host = settings.split()[0]
            port = settings.split()[1]
            address = settings.split()[2]
            password = settings.split()[3]
    return host, port, address, password

def get_monitor_res():
    return GetSystemMetrics(0), GetSystemMetrics(1)


def main():  
    print(process_running(WANTED_PROCESSES[0]))

    current_time = datetime.now()  
    current_month_and_yr = datetime.now().strftime("%B %Y")

    new_dir = make_directory(DESKTOP_PATH, current_month_and_yr)

    capture = cv2.VideoCapture(-1)

    # needed dimensions
    scr_w, scr_h = get_monitor_res() 

    # codec
    fourcc = cv2.VideoWriter_fourcc(*"XVID")

    # path
    combined_path = new_dir + "\\combined_" + current_time.strftime(DATE_FILE_FORMAT) + FILE_EXT

    #writer
    writer = cv2.VideoWriter(combined_path, fourcc, FPS, (scr_w * 2, scr_h))

    while True:
        img = ImageGrab.grab()
        img_np = np.array(img)
        screen_frame = cv2.cvtColor(img_np, cv2.COLOR_BGR2RGB)
        
        # Capture frame-by-frame
        ret, cam_frame = capture.read()
        if not ret:
            break

        cam_frame_resized = cv2.resize(cam_frame, dsize=(scr_w, scr_h), interpolation=cv2.INTER_CUBIC)          

        (B_C, G_C, R_C) = cv2.split(cam_frame_resized)
        (B_S, G_S, R_S) = cv2.split(screen_frame)

        cam_frame_merged = cv2.merge([B_C, G_C, R_C])
        screen_frame_merged = cv2.merge([B_S, G_S, R_S])
            
        final_output = np.zeros((scr_h, scr_w * 2, 3), dtype="uint8")
        final_output[0:scr_h, 0:scr_w] = cam_frame_merged
        final_output[0:scr_h, scr_w:scr_w * 2] = screen_frame_merged

        writer.write(final_output)

        key = cv2.waitKey(1) & 0xFF
        if key == ord("q"):
            break

    # When everything done, release the capture
    capture.release()
    writer.release()
    cv2.destroyAllWindows()

    # send email on certain date
    if datetime.now().day == SET_DAY:
        host, port, address, password = get_settings(SETTINGS_FILE)
        sender = EmailSender(host, port, address, password)
        message = "Monthly Email of Video Capture"
        sender.send_message(message=message, dt_file_path=new_dir, attach=True) 

if __name__ == "__main__":
        main()
