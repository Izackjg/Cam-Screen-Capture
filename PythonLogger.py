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
from Capture import Capture

WANTED_PROCESSES = ["fortniteclient-win64-shipping", "chrome"]

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
    current_time = datetime.now()  
    current_month_and_yr = datetime.now().strftime("%B %Y")

    new_dir = make_directory(DESKTOP_PATH, current_month_and_yr)

    #capture = cv2.VideoCapture(-1)

    # needed dimensions
    scr_w, scr_h = get_monitor_res() 

    # codec
    #fourcc = cv2.VideoWriter_fourcc(*"XVID")

    # path
    combined_save_path = new_dir + "\\combined_" + current_time.strftime(DATE_FILE_FORMAT) + FILE_EXT

    cap_class = Capture(-1, combined_save_path, [scr_w, scr_h])

    #writer
    capture = cap_class.capture
    writer = cap_class.get_writer()

    while True:
        img = ImageGrab.grab()
        img_np = np.array(img)
        screen_frame = cv2.cvtColor(img_np, cv2.COLOR_BGR2RGB)
        
        # Capture frame-by-frame
        ret, cam_frame = capture.read()
        if not ret:
            break

        cv2.imshow("Frame", cam_frame)

        cam_frame_resized = cap_class.resize(cam_frame)
        
        cam_frame_merged = cap_class.merge_frame(cam_frame_resized)
        screen_frame_merged = cap_class.merge_frame(screen_frame)       
            
        output = cap_class.get_output(cam_frame_merged, screen_frame_merged)

        writer.write(output)

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