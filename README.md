# Gesture-Based Smart Presentation Assistant

A computer-vision system that lets a presenter control a slideshow hands-free — no clicker, no keyboard. It uses hand-gesture recognition to navigate slides, facial recognition to make sure only the authorized presenter is in control, and an optional emotion-detection feature that can auto-pause the presentation if the audience looks confused.

## Why

Manual clickers and keyboard shortcuts interrupt the flow of a presentation and reduce audience engagement. This project replaces that with a natural, camera-based interface.

## Features

- **Presenter authentication** — uses MediaPipe Face Mesh (478 landmarks) to verify the presenter's identity before accepting any gesture commands.
- **Gesture-based slide control**:
  - `START` — four fingers up, thumb out
  - `NEXT` — index finger up only
  - `PREVIOUS` — index and middle fingers up
  - `END` — closed fist
- **Real-time on-screen feedback** — shows authentication status, current gesture, and (optionally) detected emotion, with hand/face landmark overlays.
- **Emotion-driven auto-pause (optional)** — sends frames to the Gemini API to detect a confused expression and automatically blanks/pauses the slideshow.
- **PowerPoint integration** — controls PowerPoint directly via the Windows COM API, with a keyboard-simulation fallback (via PyAutoGUI) for broader compatibility.
- **Voice feedback** — spoken confirmations (e.g. "Next slide activated") via Windows SAPI text-to-speech.

## Tech stack

| Purpose | Library |
|---|---|
| Video capture & image processing | OpenCV |
| Hand & face landmark detection | MediaPipe (Hands, Face Mesh) |
| Numerical / distance calculations | NumPy |
| Keyboard simulation fallback | PyAutoGUI |
| PowerPoint control & text-to-speech | pywin32 (COM API, SAPI) |
| Emotion analysis | Google Generative AI (Gemini 1.5 Flash) |
| Image handling for the Gemini API | Pillow |

## How it works

1. On startup, the app loads reference photos of authorized presenters from the `data/` folder and extracts their face landmarks.
2. Each webcam frame is checked against those landmarks; if the mean landmark distance is below a set threshold, the current person is authenticated.
3. Once authenticated, hand landmarks are analyzed each frame to detect one of the defined gestures, which is mapped to a slide command.
4. Commands are sent to PowerPoint via the COM API, falling back to simulated key presses if that isn't available.
5. If enabled, a background thread periodically sends frames to the Gemini API to check for a confused expression and pauses the show if detected.

## Setup

```bash
git clone https://github.com/<your-username>/gesture-presentation-assistant.git
cd gesture-presentation-assistant
pip install -r requirements.txt
```

Add one or more clear, front-facing reference photos of authorized presenters to the `data/` folder (filename, minus extension, becomes the presenter's display name — e.g. `data/john.jpg`).

Run it:

```bash
python main.py
```

If you want the emotion-detection feature, you'll be prompted for a Gemini API key on startup (get one from [Google AI Studio](https://aistudio.google.com/)). Leave it blank to skip emotion detection. This project was built and tested on **Windows** (it uses `pywin32`/COM for PowerPoint control and SAPI voice feedback).

## Known limitations / future work

- Camera calibration (lens-distortion correction) is scoped but not yet wired into the main loop.
- Gesture set is currently fixed; a natural next step is letting users define custom gestures.
- Authentication currently uses a single reference photo per person; enrolling multiple images per presenter would improve reliability.
- PowerPoint control is Windows-only; broader cross-platform support is a planned improvement.

## Team

Built as a graduation/team project by Ziad Tarek, Ismaaiel Hossam Eldien, Abdelrahman Sayed, Ali Abdelrahim Elsaid, Islam Abdelhady, and Ahmed Zaky.
