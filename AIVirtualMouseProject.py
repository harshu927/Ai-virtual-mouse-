import cv2
import numpy as np
import HandTrackingModule12 as htm
import time
import autopy
import pyttsx3
from pynput.mouse import Controller as MouseController
import speech_recognition as sr
import pyautogui
import pytesseract
import os
from datetime import datetime
import ctypes
import psutil
from threading import  Thread
import pygame

# Initialize pygame mixer
pygame.mixer.init()

# Tesseract OCR Path
pytesseract.pytesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

if not os.path.exists(pytesseract.pytesseract_cmd):
    print("Error: Tesseract OCR is not installed or the path is incorrect.")
    exit()

# Sound File Paths
thumbs_up_sound = "C:\\Users\\KIIT\\Desktop\\Ai important virtual mouse2\\thumbs_up.wav"
funny_on_sound = "C:\\Users\\KIIT\\Desktop\\Ai important virtual mouse2\\funny_on.wav"

##########################
wCam, hCam = 640, 480
frameR = 100
smoothening = 5
skip_frames = 2
##########################

# Initialize Variables
pTime = 0
plocX, plocY = 0, 0
clocX, clocY = 0, 0
drawing_function = False
text_function = False
prev_x, prev_y = 0, 0
drawing_colour = (0, 255, 0)
shape_function = None
save_path = "screenshots"
command_result = None
show_performance = False
last_action_time = time.time()
auto_lock_enabled = False

# Ensure screenshot folder exists
if not os.path.exists(save_path):
    os.makedirs(save_path)

# Initialize Controllers
mouse = MouseController()
engine = pyttsx3.init()
recognizer = sr.Recognizer()

# Initialize Camera
cap = cv2.VideoCapture(0)
cap.set(3, wCam)
cap.set(4, hCam)

if not cap.isOpened():
    print("Error: Webcam not accessible.")
    exit()

detector = htm.handDetector(maxHands=1)
wScr, hScr = autopy.screen.size()

# Flags for Functions
mouse_control_enabled = True
gesture_function = "Normal"
frame_counter = 0

# Voice Feedback Function
def speak(text):
    engine.say(text)
    engine.runAndWait()

# Multi-threaded Voice Command Function
def listen_command():
    global command_result
    try:
        with sr.Microphone() as source:
            print("Listening for command...")
            recognizer.adjust_for_ambient_noise(source)
            audio = recognizer.listen(source, timeout=5, phrase_time_limit=5)
        command_result = recognizer.recognize_google(audio).lower()
        print(f"Command received: {command_result}")
    except Exception as e:
        print(f"Voice recognition error: {e}")
        command_result = None

# OCR Function
def perform_ocr(image):
    try:
        # Step 1: Convert the image to grayscale
        gray_image = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)

        # Step 2: Apply bilateral filtering to remove noise while preserving edges
        denoised_image = cv2.bilateralFilter(gray_image, 9, 75, 75)

        # Step 3: Sharpen the image for better edge detection
        kernel = np.array([[0, -1, 0], [-1, 5, -1], [0, -1, 0]])  # Sharpen filter
        sharpened_image = cv2.filter2D(denoised_image, -1, kernel)

        # Step 4: Adaptive thresholding for binarization
        binary_image = cv2.adaptiveThreshold(
            sharpened_image,
            255,
            cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
            cv2.THRESH_BINARY_INV,  # Use inverted binary for better OCR results
            11,
            2
        )

        # Step 5: Debugging - Save intermediate processed images
        cv2.imwrite("debug_gray.png", gray_image)
        cv2.imwrite("debug_denoised.png", denoised_image)
        cv2.imwrite("debug_sharpened.png", sharpened_image)
        cv2.imwrite("debug_binary.png", binary_image)

        # Step 6: Perform OCR using Tesseract
        custom_config = r'--oem 3 --psm 6'  # OCR Engine Mode and Page Segmentation Mode
        text = pytesseract.image_to_string(binary_image, config=custom_config)

        # Check for extracted text
        if text.strip():
            print(f"OCR Result:\n{text}")
            speak("Text extracted successfully.")
        else:
            print("No text detected. Ensure the text is clear and visible.")
            speak("No text detected in the given image.")

    except Exception as e:
        print(f"Error during OCR: {e}")
        speak("An error occurred while performing OCR.")

# Function to Draw Bounding Boxes
def draw_bounding_boxes(img, bbox):
    x, y, w, h = bbox
    # Draw pink bounding box
    cv2.rectangle(img, (x - 20, y - 20), (x + w + 20, y + h + 20), (255, 0, 255), 2)
    # Draw green bounding box
    cv2.rectangle(img, (x, y), (x + w, y + h), (0, 255, 0), 2)

# Execute Commands
def execute_voice_command(command):
    global mouse_control_enabled, gesture_function, drawing_function, text_function, drawing_colour, shape_function, show_performance, auto_lock_enabled

    try:
        if "enable mouse" in command:
            mouse_control_enabled = True
            speak("Mouse control enabled.")
        elif "disable mouse" in command:
            mouse_control_enabled = False
            speak("Mouse control disabled.")
        elif "zoom function" in command:
            gesture_function = "Zoom"
            speak("Zoom function activated.")
        elif "scroll function" in command:
            gesture_function = "Scroll"
            speak("Scroll function activated.")
        elif "normal function" in command:
            gesture_function = "Normal"
            speak("Normal function activated.")
        elif "drawing function" in command:
            drawing_function = not drawing_function
            speak("Drawing function activated." if drawing_function else "Drawing function deactivated.")
        elif "text function" in command:
            text_function = not text_function
            speak("Text function activated." if text_function else "Text function deactivated.")
        elif "draw rectangle" in command:
            shape_function = "rectangle"
            speak("Rectangle drawing function activated.")
        elif "draw circle" in command:
            shape_function = "circle"
            speak("Circle drawing function activated.")
        elif "draw ellipse" in command:
            shape_function = "ellipse"
            speak("Ellipse drawing function activated.")
        elif "clear shapes" in command:
            shape_function = None
            speak("Shape function cleared.")
        elif "clear canvas" in command:
            canvas = np.zeros((hCam, wCam, 3), dtype=np.uint8)
            speak("Canvas cleared.")
        elif "take screenshot" in command:
            filename = f"{save_path}/screenshot_{datetime.now().strftime('%Y%m%d_%H%M%S')}.png"
            pyautogui.screenshot(filename)
            speak(f"Screenshot saved as {filename}.")
        elif "increase volume" in command:
            ctypes.windll.winmm.waveOutSetVolume(0, 65535)
            speak("Volume increased to maximum.")
        elif "decrease volume" in command:
            ctypes.windll.winmm.waveOutSetVolume(0, 0)
            speak("Volume decreased to minimum.")
        elif "open browser" in command:
            os.system("start chrome")
            speak("Browser opened.")
        elif "lock screen" in command:
            ctypes.windll.user32.LockWorkStation()
            speak("Screen locked.")
        elif "perform ocr" in command:
            perform_ocr(img.copy())
        elif "show performance" in command:
            show_performance = True
            speak("Performance overlay enabled.")
        elif "hide performance" in command:
            show_performance = False
            speak("Performance overlay disabled.")
        elif "help" in command:
            speak("Here are the available commands: enable mouse, disable mouse, zoom function, scroll function, normal function, drawing function, text function, draw rectangle, draw circle, draw ellipse, clear shapes, take screenshot, perform ocr, increase volume, decrease volume, open browser, lock screen, show performance, hide performance, help menu, exit.")
        elif "exit" in command:
            speak("Exiting program.")
            exit()

        elif "change colour" in command:  # Handle color change commands
            command = command.lower().strip()  # Normalize the command for consistency

            if "red" in command:
                drawing_colour = (0, 0, 255)
                speak("Drawing colour changed to red.")
            elif "blue" in command:
                drawing_colour = (255, 0, 0)
                speak("Drawing colour changed to blue.")
            elif "green" in command:
                drawing_colour = (0, 255, 0)
                speak("Drawing colour changed to green.")
            else:
                speak("Colour not recognized.")

        else:
            speak("Command not recognized.")
    except Exception as e:
        print(f"Error executing voice command: {e}")

# Function to Play Sound
def play_sound(file_path):
    try:
        if os.path.exists(file_path):
            pygame.mixer.music.load(file_path)
            pygame.mixer.music.play()
        else:
            print(f"Sound file not found: {file_path}")
            speak("Sound file not found.")
    except Exception as e:
        print(f"Error playing sound: {e}")

# Function to Display System Performance
def display_system_performance(img):
    try:
        cpu_usage = psutil.cpu_percent()
        memory_info = psutil.virtual_memory()
        battery = psutil.sensors_battery()
        battery_percent = battery.percent if battery else "N/A"

        performance_text = f"CPU: {cpu_usage}% | Memory: {memory_info.percent}% | Battery: {battery_percent}%"
        cv2.putText(img, performance_text, (10, hCam - 10), cv2.FONT_HERSHEY_PLAIN, 1, (0, 255, 0), 1)
    except Exception as e:
        print(f"Error displaying performance: {e}")
last_action_time = 0
canvas = np.zeros((hCam, wCam, 3), dtype=np.uint8)  # Persistent drawing canvas
# Main Loop
while True:
    success, img = cap.read()
    if not success:
        print("Frame capture failed.")
        continue

    frame_counter += 1
    if frame_counter % skip_frames != 0:
        continue

    img = detector.findHands(img, draw=True)
    lmList, bbox = detector.findPosition(img, draw=True)

    if bbox:
        draw_bounding_boxes(img, bbox)

    if lmList:
        x1, y1 = lmList[8][1:]
        fingers = detector.fingersUp()

        # Mouse Movement
        if mouse_control_enabled and fingers[1] == 1 and fingers[2] == 0:
            x3 = np.interp(x1, (frameR, wCam - frameR), (0, wScr))
            y3 = np.interp(y1, (frameR, hCam - frameR), (0, hScr))
            clocX = plocX + (x3 - plocX) / smoothening
            clocY = plocY + (y3 - plocY) / smoothening
            autopy.mouse.move(wScr - clocX, clocY)
            plocX, plocY = clocX, clocY
            cv2.circle(img, (x1, y1), 15, (255, 0, 255), cv2.FILLED)

        # Scroll Function
        if gesture_function == "Scroll" and fingers[1] == 1 and fingers[2] == 1:
            length, img, lineInfo = detector.findDistance(8, 12, img)  # Distance between Index and Middle fingers
            if length < 30:  # Scroll Down
                pyautogui.scroll(-100)
                speak("Scrolling down.")
            elif length > 50:  # Scroll Up
                pyautogui.scroll(100)
                speak("Scrolling up.")

        # Mouse Click
        if fingers[1] == 1 and fingers[2] == 1:
            length, img, lineInfo = detector.findDistance(8, 12, img)
            if length < 40:
                cv2.circle(img, (lineInfo[4], lineInfo[5]), 15, (0, 255, 0), cv2.FILLED)
                autopy.mouse.click()
                speak("Mouse clicked.")

        # Gesture Recognition
        if fingers == [1, 1, 0, 0, 1]:
            play_sound(thumbs_up_sound)
        elif fingers == [0, 1, 1, 1, 1]:
            play_sound(funny_on_sound)

        # Text Function: Display text on the screen
        if text_function and fingers[1] == 1 and fingers[2] == 0:
            cv2.putText(img, "Text Function Active", (x1, y1), cv2.FONT_HERSHEY_SIMPLEX, 1, (255, 255, 255), 2)

        # Zoom Function
        if gesture_function == "Zoom" and fingers[1] == 1 and fingers[2] == 1:
            # Calculate distance between Thumb and Index fingers
            length_thumb_index, img, lineInfo = detector.findDistance(4, 8, img)
            # Provide thresholds for zoom in and zoom out
            if length_thumb_index < 40:  # Zoom Out
                pyautogui.hotkey("ctrl", "-")
                speak("Zooming out.")
            elif length_thumb_index > 120:  # Zoom In
                pyautogui.hotkey("ctrl", "+")
                speak("Zooming in.")

        # Drawing Shapes on Canvas
        if shape_function and fingers[1] == 1:
            if shape_function == "rectangle":
                cv2.rectangle(canvas, (x1, y1), (x1 + 100, y1 + 50), drawing_colour, -1)  # Draw filled rectangle
            elif shape_function == "circle":
                cv2.circle(canvas, (x1, y1), 50, drawing_colour, -1)  # Draw filled circle
            elif shape_function == "ellipse":
                cv2.ellipse(canvas, (x1, y1), (50, 30), 0, 0, 360, drawing_colour, -1)  # Draw filled ellipse

    # Display Active Drawing Color
    cv2.putText(img, f"Color: {'Red' if drawing_colour == (0, 0, 255) else 'Blue' if drawing_colour == (255, 0, 0) else 'Green'}", (10, 80), cv2.FONT_HERSHEY_PLAIN, 2, drawing_colour, 2)

    # Display System Performance Overlay
    if show_performance:
        display_system_performance(img)

    # Auto Lock Feature
    if auto_lock_enabled and time.time() - last_action_time > 60:
        ctypes.windll.user32.LockWorkStation()

    # Display FPS and Current Function
    cTime = time.time()
    fps = 1 / (cTime - pTime)
    pTime = cTime
    function_text = f"Function: {gesture_function} | FPS: {int(fps)}"
    cv2.putText(img, function_text, (10, 50), cv2.FONT_HERSHEY_PLAIN, 2, (255, 0, 0), 2)

    # Show Image
    cv2.putText(img,
                f"Active Color: {'Red' if drawing_colour == (0, 0, 255) else 'Blue' if drawing_colour == (255, 0, 0) else 'Green'}",
                (10, 120), cv2.FONT_HERSHEY_SIMPLEX, 1, drawing_colour, 2)
    img = cv2.addWeighted(img, 0.5, canvas, 0.5, 0)  # Merge canvas with the frame
    cv2.imshow("AI Virtual Mouse", img)

    key = cv2.waitKey(1) & 0xFF
    if key == ord('v'):
        Thread(target=listen_command).start()
        if command_result:
            print(f"processing command: {command_result}")
            execute_voice_command(command_result)
            command_result = None
    elif key == ord('q'):
        break

    last_action_time = time.time()

cap.release()
cv2.destroyAllWindows()