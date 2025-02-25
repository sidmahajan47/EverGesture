import cv2
import time
import math
import numpy as np
import HandControlModule as htm
import pycaw
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume

# Camera Settings
wCam, hCam = 640, 480
cap = cv2.VideoCapture(0)
cap.set(3, wCam)
cap.set(4, hCam)
volBar = 0
vol = 0

# Initialize Pycaw for volume control
try:
    devices = AudioUtilities.GetSpeakers()
    interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
    volume = interface.QueryInterface(IAudioEndpointVolume)
    volRange = volume.GetVolumeRange()
    minvol = volRange[0]
    maxvol = volRange[1]
except Exception as e:
    print(f"Error accessing system volume: {e}")
    exit()

pTime = 0
detector = htm.handDetector(detectionCon=0.7)
lastTimeHandDetected = time.time()
pauseThreshold = 5  # Pause script if no hand detected for X seconds

while True:
    success, img = cap.read()
    if not success:
        continue  # Skip if the frame is not read properly

    img = detector.findHands(img)
    lmList, bbox = detector.findPosition(img, draw=False)
    length = 0  # Initialize length to prevent errors

    if lmList:
        lastTimeHandDetected = time.time()
        if len(lmList) > 8:
            x1, y1 = lmList[4][1], lmList[4][2]  # Thumb tip
            x2, y2 = lmList[8][1], lmList[8][2]  # Index finger tip
            cx, cy = (x1 + x2) // 2, (y1 + y2) // 2  # Midpoint

            # Draw landmarks and line
            cv2.circle(img, (x1, y1), 15, (0, 255, 0), cv2.FILLED)
            cv2.circle(img, (x2, y2), 15, (0, 255, 0), cv2.FILLED)
            cv2.circle(img, (cx, cy), 15, (255, 0, 255), cv2.FILLED)
            cv2.line(img, (x1, y1), (x2, y2), (0, 255, 0), 3)

            # Calculate distance between thumb and index finger
            length = math.hypot(x2 - x1, y2 - y1)
            vol = np.interp(length, [10, 150], [minvol, maxvol])
            volBar = np.interp(length, [10, 150], [400, 150])
            volume.SetMasterVolumeLevel(vol, None)

            # Change color if distance is below threshold
            if length < 50:
                cv2.circle(img, (cx, cy), 15, (0, 0, 255), cv2.FILLED)

            # Mute if all fingers are down (closed fist)
            fingers = detector.fingersUp()
            if fingers.count(1) == 0:
                volume.SetMasterVolumeLevel(-65.0, None)

    # Volume Bar UI
    cv2.rectangle(img, (50, 150), (85, 400), (0, 255, 0), 3)
    cv2.rectangle(img, (50, int(volBar)), (85, 400), (0, 255, 0), cv2.FILLED)
    cv2.putText(img, f'{int(np.interp(volBar, [400, 150], [100, 0]))}%', (40, 420), cv2.FONT_HERSHEY_COMPLEX, 0.7, (0, 255, 0), 2)

    # Pause script if no hand detected for a while
    if time.time() - lastTimeHandDetected > pauseThreshold:
        cv2.putText(img, "No Hand Detected. Paused.", (50, 50), cv2.FONT_HERSHEY_COMPLEX, 1, (0, 0, 255), 2)

    # FPS Calculation
    cTime = time.time()
    fps = 1 / (cTime - pTime)
    pTime = cTime
    cv2.putText(img, f'FPS: {int(fps)}', (40, 50), cv2.FONT_HERSHEY_COMPLEX, 1, (255, 0, 0), 3)

    # Developer watermark
    cv2.putText(img, "Created by Siddhesh Mahajan (AKA ANONSID)", (10, hCam - 10), cv2.FONT_HERSHEY_COMPLEX, 0.5, (255, 255, 255), 1)

    cv2.imshow("Hand Tracking", img)
    if cv2.waitKey(5) & 0xFF == ord('q'):
        break

cap.release()
cv2.destroyAllWindows()
