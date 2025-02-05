# with delay 

import win32com.client
import cv2
import os
import threading
import tkinter as tk
from tkinter import filedialog, messagebox
import pyautogui
from PIL import Image, ImageTk
from cvzone.HandTrackingModule import HandDetector

# Function to open file dialog for selecting PPT
def select_ppt_file():
    file_path = filedialog.askopenfilename(title="Select PowerPoint File", filetypes=[("PowerPoint files", "*.pptx")])
    return file_path

# Function to start the presentation
def start_ppt():
    ppt_file = ppt_path.get()

    if ppt_file and os.path.exists(ppt_file):  # Ensure file was selected and exists
        global Presentation
        Application = win32com.client.Dispatch("PowerPoint.Application")
        Presentation = Application.Presentations.Open(ppt_file)
        print(f"Opening PowerPoint file: {Presentation.Name}")
        Presentation.SlideShowSettings.Run()

        # Start the gesture control
        start_camera()
    else:
        messagebox.showerror("Error", "No valid PowerPoint file selected or file does not exist!")

# Function to start the camera and track hand gestures
def start_camera():
    width, height = 640, 480

    # Camera Setup
    cap = cv2.VideoCapture(0)
    cap.set(3, width)
    cap.set(4, height)

    # Hand Detector
    detectorHand = HandDetector(detectionCon=0.7, maxHands=1)

    # Variables
    buttonPressed = False
    counter = 0
    drawing = False  
    delay_active = False  # New variable to manage delay

    # Get PowerPoint slideshow view
    slideShowView = Presentation.SlideShowWindow.View

    # Function to reset delay flag
    def reset_delay():
        nonlocal delay_active
        delay_active = False

    # Non-blocking PowerPoint navigation with delay
    def next_slide():
        nonlocal delay_active
        if not delay_active:
            root.after(0, slideShowView.Next)
            delay_active = True
            threading.Timer(1, reset_delay).start()  # 1-second delay

    def prev_slide():
        nonlocal delay_active
        if not delay_active:
            root.after(0, slideShowView.Previous)
            delay_active = True
            threading.Timer(1, reset_delay).start()  # 1-second delay

    def update_frame():
        nonlocal buttonPressed, counter, drawing

        success, img = cap.read()
        if not success:
            camera_label.after(10, update_frame)
            return

        img = cv2.flip(img, 1)
        hands, img = detectorHand.findHands(img)

        if hands:
            hand = hands[0]
            cx, cy = hand["center"]
            lmList = hand["lmList"]
            fingers = detectorHand.fingersUp(hand)

            print(f"Fingers: {fingers}")  # Debugging output

            if not buttonPressed:
                # **GESTURE CONTROL**

                # 1. **Little Finger (Last Finger) is Up -> Previous Slide**
                if fingers == [0, 0, 0, 0, 1]:  
                    print("🔙 Previous Slide")
                    buttonPressed = True
                    prev_slide()

                # 2. **Thumb (First Finger) is Up -> Next Slide**
                elif fingers == [1, 0, 0, 0, 0]:  
                    print("➡ Next Slide")
                    buttonPressed = True
                    next_slide()

                # 3. **Index & Middle Finger Up -> Activate PPT Marker**
                elif fingers == [0, 1, 1, 0, 0]:  
                    if not drawing:
                        print("🖊️ Activating PowerPoint Marker")
                        slideShowView.PointerType = 1  
                        drawing = True
                    else:
                        print("📝 Writing on PowerPoint")

                        screen_width, screen_height = pyautogui.size()
                        screen_x = int(cx / width * screen_width)
                        screen_y = int(cy / height * screen_height)

                        pyautogui.moveTo(screen_x, screen_y, _pause=False)
                        pyautogui.mouseDown()

                # 4. **Only Index Finger Up & Thumb Folded -> Activate Pointer**
                elif fingers == [0, 1, 0, 0, 0]:  
                    print("📍 Pointer Mode (Red Dot)")
                    cv2.circle(img, (cx, cy), 10, (0, 0, 255), cv2.FILLED)

                else:
                    if drawing:
                        print("❌ Stopping Drawing Mode")
                        pyautogui.mouseUp()
                        slideShowView.PointerType = 0
                        drawing = False

        if buttonPressed:
            counter += 1
            if counter > 5:  # Reduced debounce time
                counter = 0
                buttonPressed = False

        img = cv2.cvtColor(img, cv2.COLOR_BGR2RGB)
        img = Image.fromarray(img)
        img = ImageTk.PhotoImage(img)

        camera_label.img = img
        camera_label.config(image=img)
        camera_label.after(10, update_frame)

    update_frame()

# Main GUI setup
root = tk.Tk()
root.title("PowerPoint Control System")
root.geometry("700x600")

label = tk.Label(root, text="Select PowerPoint file to start the presentation", font=("Arial", 14))
label.pack(pady=10)

ppt_path = tk.StringVar()

def choose_file():
    file = select_ppt_file()
    if file:
        ppt_path.set(file)
        file_button.config(state="disabled")  
        start_button.config(state="normal")  
        label.config(text=f"Selected File: {os.path.basename(file)}")  

file_button = tk.Button(root, text="Choose PowerPoint File", font=("Arial", 12), command=choose_file)
file_button.pack(pady=10)

start_button = tk.Button(root, text="Start PPT & Control", font=("Arial", 12), state="disabled", command=start_ppt)
start_button.pack(pady=10)

camera_label = tk.Label(root)
camera_label.pack()

root.mainloop()
