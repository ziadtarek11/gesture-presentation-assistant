import cv2
import mediapipe as mp
import numpy as np
import time
import pyautogui
import os
import google.generativeai as genai
import win32com.client
import pythoncom
import pickle
import tkinter as tk
from tkinter import simpledialog, messagebox
import glob
from PIL import Image
import io
import math
import threading

mp_drawing = mp.solutions.drawing_utils
mp_drawing_styles = mp.solutions.drawing_styles
mp_hands = mp.solutions.hands
mp_face_mesh = mp.solutions.face_mesh

hands = mp_hands.Hands(
    model_complexity=0,
    min_detection_confidence=0.5,
    min_tracking_confidence=0.5)

face_mesh = mp_face_mesh.FaceMesh(
    max_num_faces=1,
    refine_landmarks=True,
    min_detection_confidence=0.5,
    min_tracking_confidence=0.5)


WINDOW_NAME = 'Gesture Presentation Assistant'
WINDOW_WIDTH = 480 
WINDOW_HEIGHT = 360 
WINDOW_POS_X = 1920 - WINDOW_WIDTH - 30 
WINDOW_POS_Y = 30

FONT = cv2.FONT_HERSHEY_SIMPLEX
FONT_SCALE = 0.7
FONT_THICKNESS = 2
COLOR_WHITE = (255, 255, 255)
COLOR_RED = (0, 0, 255)
COLOR_GREEN = (0, 255, 0)
COLOR_BLUE = (255, 0, 0)
COLOR_YELLOW = (0, 255, 255)
COLOR_ORANGE = (0, 165, 255)
TEXT_ALPHA = 0.6 


GEMINI_API_KEY = None
GEMINI_CONFIGURED = False
GEMINI_MODEL_NAME = "gemini-1.5-flash-latest" 
GEMINI_EMOTION_PROMPT = """Analyze the person in this image carefully.
Does the person have a clearly furrowed brow (eyebrows drawn together and possibly lowered)?
    - If YES, respond 'CONFUSED'.
    - If NO or unclear, respond 'NEUTRAL'.
Respond ONLY with 'CONFUSED' or 'NEUTRAL'.
""" 

# Initialize a global speaker object for SAPI
try:
    speaker = win32com.client.Dispatch("SAPI.SpVoice")
    speak_enabled = True
    print("SAPI voice initialized.")
except pythoncom.com_error:
    print("Could not initialize SAPI voice. Voice feedback will be disabled.")
    speak_enabled = False


def get_gemini_api_key():
    """Prompts the user for the Gemini API key using a Tkinter dialog."""
    global GEMINI_API_KEY, GEMINI_CONFIGURED
    if GEMINI_CONFIGURED: 
        return GEMINI_API_KEY

    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True) 

    api_key = simpledialog.askstring("Gemini API Key", "Please enter your Gemini API Key:", show='*')

    root.destroy() 

    if api_key:
        GEMINI_API_KEY = api_key
        try:
            genai.configure(api_key=GEMINI_API_KEY)
            print("Gemini API Key configured successfully.")
            GEMINI_CONFIGURED = True
            return GEMINI_API_KEY
        except Exception as e:
            print(f"Error configuring Gemini API: {e}. Emotion analysis might be disabled.")
            GEMINI_CONFIGURED = False
            return None
    else:
        print("No Gemini API Key entered. Emotion analysis will be disabled.")
        GEMINI_CONFIGURED = False
        return None


def analyze_emotion(frame_bgr):
    """Analyzes the emotion in the frame using Gemini.
       Designed to be run in a separate thread. Updates global current_emotion.
    """
    global GEMINI_CONFIGURED, GEMINI_MODEL_NAME, GEMINI_EMOTION_PROMPT, current_emotion 
    if not GEMINI_CONFIGURED:
        return 

    print(f"Analyzing frame {frame_count} for emotion (Gemini Thread)...") 
    try:
        frame_rgb = cv2.cvtColor(frame_bgr, cv2.COLOR_BGR2RGB)
        img_pil = Image.fromarray(frame_rgb)

        model = genai.GenerativeModel(GEMINI_MODEL_NAME)
        response = model.generate_content(
            [GEMINI_EMOTION_PROMPT, img_pil],
            generation_config=genai.types.GenerationConfig(
                candidate_count=1,
                max_output_tokens=50,
                temperature=0.1)
            )


        if response.parts:
            detected_emotion = response.text.strip().upper()
            valid_emotions = {"CONFUSED", "NEUTRAL"}
            if detected_emotion in valid_emotions:
                current_emotion = detected_emotion 
                print(f"Emotion Analysis Result (Thread): {current_emotion}")
            else:
                print(f"Warning: Unexpected emotion format from Gemini: {response.text.strip()}. Keeping previous: {current_emotion}")
        else:
             print(f"Gemini response empty or blocked (Thread). Reason: {response.prompt_feedback}")

    except Exception as e:
        print(f"Error during Gemini analysis (Thread): {e}")
        if "quota" in str(e).lower():
            print("Gemini API quota likely exceeded. Disabling emotion analysis temporarily.")


AUTHORIZED_PRESENTERS = {} 
FACE_MESH_THRESHOLD = 0.08

def get_image_bytes(image_path):
    """Loads an image and returns its bytes."""
    try:
        img = Image.open(image_path)
        byte_arr = io.BytesIO()
        img.save(byte_arr, format='JPEG')
        return byte_arr.getvalue()
    except Exception as e:
        print(f"Error loading image {image_path}: {e}")
        return None

def get_face_landmarks_from_image(image_path):
    """Processes an image file and returns MediaPipe face landmarks.
       Uses the globally initialized face_mesh object.
    """
    try:
        print(f"DEBUG: Attempting to read image: {image_path}")
        img = cv2.imread(image_path)
        if img is None:
            print(f"Error: Could not read image {image_path}")
            return None
        print(f"DEBUG: Image {os.path.basename(image_path)} loaded. Shape: {img.shape}")

        img_rgb = cv2.cvtColor(img, cv2.COLOR_BGR2RGB)
        print(f"DEBUG: Processing {os.path.basename(image_path)} with FaceMesh...")
        results = face_mesh.process(img_rgb)

        if results.multi_face_landmarks:
            print(f"DEBUG: Face landmarks FOUND for {os.path.basename(image_path)}.")
            return results.multi_face_landmarks[0].landmark
        else:
            print(f"DEBUG: No face landmarks found in {os.path.basename(image_path)} by face_mesh.process(). results.multi_face_landmarks is None or empty.")
            return None
    except Exception as e:
        print(f"Error processing face landmarks for {image_path}: {e}")
        return None

def load_authorized_presenters(data_folder="data"):
    """Loads presenter images, extracts face landmarks using MediaPipe Face Mesh."""
    print("Loading authorized presenters using MediaPipe Face Mesh...")
    AUTHORIZED_PRESENTERS.clear()
    image_paths = glob.glob(os.path.join(data_folder, "*.jpg")) + \
                  glob.glob(os.path.join(data_folder, "*.png"))
    if not image_paths:
        print(f"Warning: No presenter images found in '{data_folder}'. Authentication will not work.")
        if speak_enabled:
            speaker.Speak("Warning: No presenter images found. Authentication will not work.")
        return

    for img_path in image_paths:
        presenter_name = os.path.splitext(os.path.basename(img_path))[0]
        print(f"Processing presenter: {presenter_name} from image: {img_path}...")
        landmarks = get_face_landmarks_from_image(img_path)
        if landmarks:
            AUTHORIZED_PRESENTERS[presenter_name] = landmarks
            print(f"  - Face landmarks extracted for {presenter_name}. Number of landmarks: {len(landmarks)}")
        else:
            print(f"  - Failed to extract face landmarks for {presenter_name}.")
    print(f"Loaded {len(AUTHORIZED_PRESENTERS)} authorized presenters with face landmarks.")
    if speak_enabled and AUTHORIZED_PRESENTERS:
        speaker.Speak(f"Loaded {len(AUTHORIZED_PRESENTERS)} authorized presenters.")


def detect_gesture_mediapipe(frame_rgb, hand_landmarks, handedness_label): 
    """Interprets hand landmarks to detect gestures with improved distinction."""
    if not hand_landmarks or not handedness_label: 
        return "NONE"

    landmarks = hand_landmarks.landmark
    h, w, _ = frame_rgb.shape 

    def is_finger_up(tip_idx, pip_idx):
        return landmarks[tip_idx].y < landmarks[pip_idx].y - 0.015 

    def is_finger_down(tip_idx, dip_idx):
        return landmarks[tip_idx].y > landmarks[dip_idx].y + 0.005

    def is_thumb_up(thumb_tip_idx, thumb_ip_idx):
        return landmarks[thumb_tip_idx].y < landmarks[thumb_ip_idx].y - 0.03 

    def is_thumb_out(thumb_tip_idx, index_mcp_idx, handedness_label):
        thumb_x = landmarks[thumb_tip_idx].x
        index_mcp_x = landmarks[index_mcp_idx].x
        is_out = False
        threshold = 0.05 

        if handedness_label == 'Right':
            is_out = (index_mcp_x - thumb_x) > threshold
        elif handedness_label == 'Left':
            is_out = (thumb_x - index_mcp_x) > threshold
        return is_out

    index_up = is_finger_up(mp_hands.HandLandmark.INDEX_FINGER_TIP, mp_hands.HandLandmark.INDEX_FINGER_PIP)
    middle_up = is_finger_up(mp_hands.HandLandmark.MIDDLE_FINGER_TIP, mp_hands.HandLandmark.MIDDLE_FINGER_PIP)
    ring_up = is_finger_up(mp_hands.HandLandmark.RING_FINGER_TIP, mp_hands.HandLandmark.RING_FINGER_PIP)
    pinky_up = is_finger_up(mp_hands.HandLandmark.PINKY_TIP, mp_hands.HandLandmark.PINKY_PIP)

    index_down = is_finger_down(mp_hands.HandLandmark.INDEX_FINGER_TIP, mp_hands.HandLandmark.INDEX_FINGER_DIP)
    middle_down = is_finger_down(mp_hands.HandLandmark.MIDDLE_FINGER_TIP, mp_hands.HandLandmark.MIDDLE_FINGER_DIP)
    ring_down = is_finger_down(mp_hands.HandLandmark.RING_FINGER_TIP, mp_hands.HandLandmark.RING_FINGER_DIP)
    pinky_down = is_finger_down(mp_hands.HandLandmark.PINKY_TIP, mp_hands.HandLandmark.PINKY_DIP)

    thumb_up = is_thumb_up(mp_hands.HandLandmark.THUMB_TIP, mp_hands.HandLandmark.THUMB_IP)
    thumb_out = is_thumb_out(mp_hands.HandLandmark.THUMB_TIP, mp_hands.HandLandmark.INDEX_FINGER_MCP, handedness_label)


    if thumb_up and index_down and middle_down and ring_down and pinky_down:
        return "THUMBS_UP" 

    elif index_up and middle_down and ring_down and pinky_down and not thumb_up:
        return "NEXT"

    elif index_up and middle_up and ring_down and pinky_down and not thumb_up:
        return "PREVIOUS"

    elif index_up and middle_up and ring_up and pinky_up and thumb_out:
         return "START"

    elif index_down and middle_down and ring_down and pinky_down and not thumb_up:
         return "END"

    return "NONE"


def authenticate_presenter_mediapipe(current_landmarks):
    """Compares current face landmarks to authorized ones using MSE.

    Args:
        current_landmarks: A list (or list-like object) of landmark objects from MediaPipe Face Mesh for the current frame.

    Returns:
        The name of the authenticated presenter or None.
    """
    if not current_landmarks:
        print("Auth: No current face landmarks detected.")
        return None
    if not AUTHORIZED_PRESENTERS:
        print("Auth: No authorized presenters loaded.")
        return None

    min_avg_distance = float('inf')
    matched_presenter = None
    # print(f"Auth: Current FACE_MESH_THRESHOLD: {FACE_MESH_THRESHOLD}") 

    try:
        current_vecs = np.array([[lm.x, lm.y, lm.z] for lm in current_landmarks])
        if current_vecs.shape[0] != 478:
             print(f"Auth Warning: Unexpected number of current landmarks: {current_vecs.shape[0]}")
             return None
    except TypeError as e:
        print(f"Auth Error: Error converting current_landmarks to array: {e}. Landmarks type: {type(current_landmarks)}")
        return None
    except Exception as e:
        print(f"Auth Error: Unexpected error processing current_landmarks: {e}")
        return None

    # print(f"Auth: Comparing with {len(AUTHORIZED_PRESENTERS)} authorized presenters.")
    for name, stored_landmarks_list in AUTHORIZED_PRESENTERS.items():
        # print(f"Auth: Checking against presenter: {name}") 
        if not stored_landmarks_list:
             print(f"Auth Warning: Stored landmarks for {name} are empty.")
             continue

        try:
            stored_vecs = np.array([[lm.x, lm.y, lm.z] for lm in stored_landmarks_list])
        except Exception as e:
            print(f"Auth Error: Error converting stored landmarks for {name}: {e}")
            continue

        if stored_vecs.shape[0] != current_vecs.shape[0]:
            print(f"Auth Warning: Landmark count mismatch for {name}. Stored: {stored_vecs.shape[0]}, Current: {current_vecs.shape[0]}")
            continue

        try:
            distances = np.linalg.norm(current_vecs - stored_vecs, axis=1)
            avg_distance = np.mean(distances)
            print(f"Auth: Calculated distance for {name}: {avg_distance:.4f}") # Log distance for each
        except Exception as e:
            print(f"Auth Error: Error calculating distance for {name}: {e}")
            continue

        if avg_distance < min_avg_distance:
            min_avg_distance = avg_distance
            matched_presenter = name

    if matched_presenter is not None:
        print(f"Auth: Closest match is {matched_presenter} with distance: {min_avg_distance:.4f}. Threshold is {FACE_MESH_THRESHOLD}")
        if min_avg_distance < FACE_MESH_THRESHOLD:
            print(f"Auth: SUCCESS - {matched_presenter} (Distance: {min_avg_distance:.4f})")
            return matched_presenter
        else:
            print(f"Auth: FAILED - Distance {min_avg_distance:.4f} for {matched_presenter} is NOT less than threshold {FACE_MESH_THRESHOLD}")
            return None
    else:
        print("Auth: No presenter was matched (min_avg_distance remained inf or no presenters processed).")
        return None

def control_presentation(command):
    """Controls PowerPoint via COM API or simulates key presses.

    Requires PowerPoint to be running.
    """
    global speak_enabled, speaker
    print(f"Executing command: {command}")
    feedback_speech = None

    if command == "PAUSE_CONFUSION":
        try:
            pyautogui.press('b')
            print("Presentation paused (black screen toggle)")
            feedback_speech = "Presentation paused due to confusion."
            time.sleep(0.5)
        except Exception as e:
            print(f"Error executing {command} with pyautogui: {e}")
            feedback_speech = "Error pausing presentation."
        if speak_enabled and feedback_speech:
            speaker.Speak(feedback_speech)
        return 
    elif command == "RESUME_CONFUSION":
        try:
            pyautogui.press('b')
            print("Presentation resumed (black screen toggle)")
            feedback_speech = "Resuming presentation."
            time.sleep(0.5)
        except Exception as e:
            print(f"Error executing {command} with pyautogui: {e}")
            feedback_speech = "Error resuming presentation."
        if speak_enabled and feedback_speech:
            speaker.Speak(feedback_speech)
        return 

    try:
        pythoncom.CoInitialize()
        ppt_app = win32com.client.GetActiveObject("PowerPoint.Application")

        if not ppt_app.Presentations.Count > 0:
            print("PowerPoint is running, but no presentation is open.")
            if speak_enabled:
                speaker.Speak("PowerPoint is running, but no presentation is open.")
            return

        try:
            slideshow_view = ppt_app.ActivePresentation.SlideShowWindow.View
            in_slideshow = True
        except Exception:
            in_slideshow = False
            slideshow_view = None 

        if command == "NEXT":
            if in_slideshow:
                slideshow_view.Next()
                feedback_speech = "Next slide."
            else:
                print("Not in slideshow mode. Cannot go to next slide.")
                feedback_speech = "Not in slideshow mode."
        elif command == "PREVIOUS":
            if in_slideshow:
                slideshow_view.Previous()
                feedback_speech = "Previous slide."
            else:
                print("Not in slideshow mode. Cannot go to previous slide.")
                feedback_speech = "Not in slideshow mode."
        elif command == "START":
            if not in_slideshow:
                if ppt_app.ActiveWindow:
                    ppt_app.ActivePresentation.SlideShowSettings.Run()
                    feedback_speech = "Starting presentation."
                else:
                    print("No active presentation window found to start slideshow.")
                    feedback_speech = "No active presentation to start."
            else:
                print("Already in slideshow mode.")
                feedback_speech = "Already in slideshow mode."
        elif command == "END":
            if in_slideshow:
                slideshow_view.Exit()
                feedback_speech = "Ending presentation."
            else:
                print("Not in slideshow mode.")
                feedback_speech = "Not in slideshow mode."
        
        if speak_enabled and feedback_speech:
            speaker.Speak(feedback_speech)

        time.sleep(0.5)

    except pythoncom.com_error as e:
        error_message = "PowerPoint application not found or not running."
        if e.hresult == -2147221021 or "GetActiveObject" in str(e):
            print(error_message)
        else:
            print(f"PowerPoint COM Error: {e}")
            error_message = "PowerPoint error."
        if speak_enabled:
            speaker.Speak(error_message)

    except Exception as e:
        print(f"Error controlling PowerPoint: {e}")
        if speak_enabled:
            speaker.Speak("Error controlling presentation.")
    finally:
        pythoncom.CoUninitialize()


def draw_ui(frame, gesture, auth_user, emotion, is_paused, hand_landmarks, face_landmarks):
    """Draws the UI elements directly onto the frame."""
    h, w, _ = frame.shape 
    y_pos = 30 

    text_lines = 3
    max_text_width = 300
    text_bg_height = text_lines * 35 + 10
    text_bg_y_end = y_pos + text_bg_height

    y_start_bg = max(0, y_pos - 25)
    y_end_bg = min(h, text_bg_y_end)
    x_start_bg = 10
    x_end_bg = min(w, 10 + max_text_width)

    if y_start_bg < y_end_bg and x_start_bg < x_end_bg: 
        sub_img = frame[y_start_bg : y_end_bg, x_start_bg : x_end_bg]
        black_rect = np.zeros(sub_img.shape, dtype=np.uint8)
        res = cv2.addWeighted(sub_img, 1 - TEXT_ALPHA, black_rect, TEXT_ALPHA, 0)
        frame[y_start_bg : y_end_bg, x_start_bg : x_end_bg] = res 

    gesture_text = f"Gesture: {gesture}"
    gesture_color = COLOR_YELLOW if gesture != "NONE" else COLOR_WHITE
    cv2.putText(frame, gesture_text, (10, y_pos), FONT, FONT_SCALE, gesture_color, FONT_THICKNESS, cv2.LINE_AA) 
    y_pos += 30

    auth_text = f"Auth: {auth_user}" if auth_user else "Auth: -"
    auth_color = COLOR_GREEN if auth_user else COLOR_RED
    cv2.putText(frame, auth_text, (10, y_pos), FONT, FONT_SCALE, auth_color, FONT_THICKNESS, cv2.LINE_AA) 
    y_pos += 30

    emotion_text = f"Emotion: {emotion}" if emotion else "Emotion: -"
    emotion_color = COLOR_ORANGE if emotion == "CONFUSED" else COLOR_WHITE
    cv2.putText(frame, emotion_text, (10, y_pos), FONT, FONT_SCALE, emotion_color, FONT_THICKNESS, cv2.LINE_AA) 
    y_pos += 30

    if is_paused:
        pause_text = "PAUSED (Confused)"
        text_size, _ = cv2.getTextSize(pause_text, FONT, FONT_SCALE * 1.2, FONT_THICKNESS + 1)
        text_x = (w - text_size[0]) // 2
        text_y = h - 30
        cv2.putText(frame, pause_text, (text_x, text_y), FONT, FONT_SCALE * 1.2, COLOR_RED, FONT_THICKNESS + 1, cv2.LINE_AA) 

    if hand_landmarks:
        mp_drawing.draw_landmarks(
            frame, 
            hand_landmarks,
            mp_hands.HAND_CONNECTIONS,
            mp_drawing_styles.get_default_hand_landmarks_style(),
            mp_drawing_styles.get_default_hand_connections_style())

    if face_landmarks:
        mp_drawing.draw_landmarks(
            image=frame,
            landmark_list=face_landmarks,
            connections=mp_face_mesh.FACEMESH_CONTOURS,
            landmark_drawing_spec=None,
            connection_drawing_spec=mp_drawing_styles.get_default_face_mesh_contours_style())


def main():
    global GEMINI_API_KEY, GEMINI_CONFIGURED, AUTHORIZED_PRESENTERS
    global frame_count, last_authenticated_user, last_auth_check_frame, last_emotion_analysis_frame, current_emotion, emotion_analysis_thread
    global speak_enabled, speaker # Add speaker to globals for main

    EMOTION_ANALYSIS_INTERVAL = 30 
    AUTH_CHECK_INTERVAL = 60 

    frame_count = 0
    last_authenticated_user = None
    last_auth_check_frame = -AUTH_CHECK_INTERVAL 
    last_emotion_analysis_frame = -EMOTION_ANALYSIS_INTERVAL 
    current_emotion = "NEUTRAL" 
    emotion_analysis_thread = None 
    user_just_authenticated = False # To announce authentication only once

    get_gemini_api_key()
    load_authorized_presenters() 

    cap = cv2.VideoCapture(0)
    if not cap.isOpened():
        print("Error: Could not open webcam.")
        return

    cv2.namedWindow(WINDOW_NAME, cv2.WINDOW_NORMAL)
    cv2.resizeWindow(WINDOW_NAME, WINDOW_WIDTH, WINDOW_HEIGHT)
    cv2.moveWindow(WINDOW_NAME, WINDOW_POS_X, WINDOW_POS_Y)

    current_gesture = "NONE"
    last_gesture_time = time.time()
    gesture_debounce_time = 1.0 
    presentation_paused_by_emotion = False

    print("Starting video capture... Press 'q' to quit.")

    while cap.isOpened():
        success, frame = cap.read()
        if not success:
            print("Ignoring empty camera frame.")
            continue

        frame_count += 1
        frame_rgb = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
        frame_rgb.flags.writeable = False 

        hand_results = hands.process(frame_rgb)
        face_results = face_mesh.process(frame_rgb)
        frame_rgb.flags.writeable = True 

        detected_gesture_this_frame = "NONE"
        processed_hand_landmarks = None
        processed_face_landmarks = None

        current_face_landmarks = None
        if face_results.multi_face_landmarks:
            current_face_landmarks = face_results.multi_face_landmarks[0].landmark
            processed_face_landmarks = face_results.multi_face_landmarks[0]

            if not last_authenticated_user or (frame_count - last_auth_check_frame >= AUTH_CHECK_INTERVAL):
                authenticated_user = authenticate_presenter_mediapipe(current_face_landmarks)
                last_auth_check_frame = frame_count
                if authenticated_user:
                    if not last_authenticated_user and speak_enabled : # New authentication
                        speaker.Speak(f"Welcome, {authenticated_user}.")
                        user_just_authenticated = True 
                    last_authenticated_user = authenticated_user
                elif last_authenticated_user:
                    print("Authentication lost. Re-checking...")
                    if speak_enabled:
                        speaker.Speak("Authentication lost.")
                    last_authenticated_user = None
                    user_just_authenticated = False
        else:
            if last_authenticated_user:
                 print("Face lost. Resetting authentication.")
                 if speak_enabled:
                     speaker.Speak("Face lost. Authentication reset.")
            last_authenticated_user = None
            user_just_authenticated = False
            last_auth_check_frame = frame_count

        if last_authenticated_user and GEMINI_CONFIGURED and \
           (frame_count - last_emotion_analysis_frame >= EMOTION_ANALYSIS_INTERVAL) and \
           (emotion_analysis_thread is None or not emotion_analysis_thread.is_alive()):

            last_emotion_analysis_frame = frame_count
            frame_copy_for_thread = frame.copy()
            emotion_analysis_thread = threading.Thread(target=analyze_emotion, args=(frame_copy_for_thread,))
            emotion_analysis_thread.start()

        if hand_results.multi_hand_landmarks:
            for i, hand_landmarks in enumerate(hand_results.multi_hand_landmarks):
                processed_hand_landmarks = hand_landmarks
                handedness = hand_results.multi_handedness[i].classification[0].label
                detected_gesture_this_frame = detect_gesture_mediapipe(frame_rgb, hand_landmarks, handedness)
                if detected_gesture_this_frame != "NONE":
                    break

        if detected_gesture_this_frame != "NONE" and last_authenticated_user:
            current_time = time.time()
            if detected_gesture_this_frame != current_gesture or (current_time - last_gesture_time > gesture_debounce_time):
                command_to_run = detected_gesture_this_frame
                if command_to_run == "THUMBS_UP": command_to_run = "START"

                # Voice feedback for gestures is now handled within control_presentation
                if not presentation_paused_by_emotion or command_to_run in ["START", "END"]:
                     control_presentation(command_to_run) # control_presentation will handle speech
                     current_gesture = detected_gesture_this_frame
                     last_gesture_time = current_time
                else:
                    print("Ignoring gesture, presentation paused due to confusion.")
        elif detected_gesture_this_frame == "NONE":
             current_gesture = "NONE"


        if last_authenticated_user and current_emotion == "CONFUSED" and not presentation_paused_by_emotion:
            print("CONFUSED emotion detected. Pausing presentation.")
            control_presentation("PAUSE_CONFUSION") # control_presentation will handle speech
            presentation_paused_by_emotion = True
        elif presentation_paused_by_emotion and current_emotion != "CONFUSED":
            print("Emotion no longer CONFUSED. Resuming presentation.")
            control_presentation("RESUME_CONFUSION") # control_presentation will handle speech
            presentation_paused_by_emotion = False

        draw_ui(frame, detected_gesture_this_frame, last_authenticated_user, current_emotion, presentation_paused_by_emotion, processed_hand_landmarks, processed_face_landmarks)

        cv2.imshow('Gesture Presentation Assistant', frame)

        if cv2.waitKey(5) & 0xFF == ord('q'):
            break

    cap.release()
    cv2.destroyAllWindows()
    if hands: hands.close()
    if face_mesh: face_mesh.close()
    if emotion_analysis_thread and emotion_analysis_thread.is_alive():
        print("Waiting for emotion analysis thread to finish...")
        emotion_analysis_thread.join()
    print("Video capture stopped.")

if __name__ == '__main__':
    main()
