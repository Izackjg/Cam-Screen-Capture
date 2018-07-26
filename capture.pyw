import sys
import win32com.client
import os
import cv2
import webbrowser
import winshell
import numpy as np
import smtplib
import psutil
from EmailSender import EmailSender
from PIL import ImageGrab
from selenium import webdriver
from win32api import GetSystemMetrics
from datetime import datetime, timedelta, time

DATE_FILE_FORMAT = "%d-%m-%Y %H.%M"

FPS = 5
RESOLUTION = "480p"
FILE_EXT = ".mp4"

DESKTOP_PATH = winshell.desktop(common=False)
SET_DAY = 23

def process_running(process_name, extension):
    full_process_name = process_name + extension
    for proc in psutil.process_iter():
        proc_name = proc.name()
        if proc_name == full_process_name:
            return True
    return False

def make_directory(desktop, dt):
    start_directory = desktop + "\\Captures"
    new_dir = start_directory + "\\" + dt
    if not os.path.exists(new_dir):
        os.makedirs(new_dir)
    return new_dir

def get_settings(filename):
    with open(filename, "r") as settings_file:
        for settings in settings_file:
            host = settings.split()[0]
            port = settings.split()[1]
            address = settings.split()[2]
            password = settings.split()[3]
    return host, port, address, password

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
    current_month_and_yr = datetime.now().strftime("%B %Y")

    new_dir = make_directory(DESKTOP_PATH, current_month_and_yr)

    chrome_running = process_running("chrome", ".exe")

    if chrome_running:
        cap = cv2.VideoCapture(-1)

        # needed dimensions
        scr_w, scr_h = get_monitor_res() 
        w = int(cap.get(3))
        h = int(cap.get(4))

        # codec
        fourcc = cv2.VideoWriter_fourcc(*"XVID")
    
        cam_path = new_dir + "\\cam_capture_" + current_time.strftime(DATE_FILE_FORMAT) + FILE_EXT
        scr_path = new_dir + "\\screen_capture" + current_time.strftime(DATE_FILE_FORMAT) + FILE_EXT
        combined_path = new_dir + "\\combined" + current_time.strftime(DATE_FILE_FORMAT) + FILE_EXT

        out_cam = cv2.VideoWriter(cam_path, fourcc, FPS, (w, h))       
        writer = cv2.VideoWriter(combined_path, fourcc, FPS, (w * 2, h))
        out_scr = cv2.VideoWriter(scr_path, fourcc, FPS, (scr_w, scr_h))

        while chrome_running or cap.isOpened():
            chrome_running = process_running("chrome", ".exe")

            img = ImageGrab.grab()
            img_np = np.array(img)
            screen_frame = cv2.cvtColor(img_np, cv2.COLOR_BGR2RGB)
        
            # Capture frame-by-frame
            ret, cam_frame = cap.read()
            if not ret:
                break

            screen_frame_resized = cv2.resize(screen_frame, dsize=(w, h), interpolation=cv2.INTER_CUBIC)

            # write frame to output file
            out_cam.write(cam_frame)
            out_scr.write(screen_frame)

            (B_C, G_C, R_C) = cv2.split(cam_frame)
            (B_S, G_S, R_S) = cv2.split(screen_frame_resized)

            cam_frame_merged = cv2.merge([B_C, G_C, R_C])
            screen_frame_merged = cv2.merge([B_S, G_S, R_S])
            
            final_output = np.zeros((h, w * 2, 3), dtype="uint8")
            final_output[0:h, 0:w] = cam_frame_merged
            final_output[0:h, w:w * 2] = screen_frame_merged

            #cv2.imshow("frame", cam_frame)
            cv2.imshow("output", final_output)

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

    # send email on certain date
    if datetime.now().day == SET_DAY:
        host, port, address, password = get_settings(DESKTOP_PATH + "\\settings.txt")
        sender = EmailSender(host, port, address, password)
        message = "Monthly Email of Video Capture"
        sender.send_message(message=message, dt_file_path=new_dir, attach=True) 

if __name__ == "__main__":
        main()
