import cv2
import time
import numpy as np
import HandTrackingModule as ht
import math
import platform

if platform.system() == 'Windows':
    from ctypes import cast, POINTER
    from comtypes import CLSCTX_ALL
    from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
elif platform.system() == 'Linux':
    import subprocess
elif platform.system() == 'Darwin':
    pass
    # Need to work on Darwin
    # pip install appscript == 1.0.0
    # from osax import *

wCam, hCam = 640, 480

cap = cv2.VideoCapture(0)
cap.set(3, wCam)
cap.set(4, hCam)
pTime = 0
length = 0
volumePercent = 0
detector = ht.handDetector(detectionCon=0.7)

if platform.system() == 'Windows':
    devices = AudioUtilities.GetSpeakers()
    interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
    volume = cast(interface, POINTER(IAudioEndpointVolume))
    volRange = volume.GetVolumeRange()
    minVolRange, maxVolRange = volRange[0], volRange[1]
    length = volRange[0]
elif platform.system() == 'Linux' or platform.system() == 'Darwin':
    minVolRange, maxVolRange = 0, 100
    # if platform.system() == 'Darwin':
    #     sa = OSAX()
else:
    print("Cannot access to Operating System's Volume Control.")
    exit()

while True:
    success, img = cap.read()
    img = detector.findHands(img)
    lmList = detector.findPosition(img, draw=False)
    if len(lmList) != 0:
        x1, y1 = lmList[4][1], lmList[4][2]
        x2, y2 = lmList[8][1], lmList[8][2]
        cx, cy = (x1 + x2) // 2, (y1 + y2) // 2
        cv2.circle(img, (x1, y1), 15, (0, 255, 255), cv2.FILLED)
        cv2.circle(img, (x2, y2), 15, (0, 255, 255), cv2.FILLED)
        cv2.line(img, (x1, y1), (x2, y2), (0, 255, 255), 3)
        cv2.circle(img, (cx, cy), 15, (0, 255, 0), cv2.FILLED)
        length = math.hypot(x2 - x1, y2 - y1)
        vol = np.interp(length, [40, 200], [minVolRange, maxVolRange])
        volumePercent = int(np.interp(vol, [0, 100], [minVolRange, maxVolRange]))
        if platform.system() == 'Windows':
            volume.SetMasterVolumeLevel(vol, None)
        elif platform.system() == 'Linux':
            subprocess.call(["amixer", "-D", "pulse", "sset", "Master", str(vol) + "%"], stdout=subprocess.DEVNULL,
                            stderr=subprocess.STDOUT)
        elif platform.system() == 'Darwin':
            # sa.set_volume(i*2)
            pass

        if length < 50:
            cv2.circle(img, (cx, cy), 15, (255, 0, 0), cv2.FILLED)

    cv2.rectangle(img, (50, 150), (85, 400), (0, 255, 0), 3)
    cv2.rectangle(img, (50, int(np.interp(length, [40, 200], [400, 150]))), (85, 400), (0, 255, 0), cv2.FILLED)
    cv2.putText(img, f'{volumePercent} %', (40, 450), cv2.FONT_HERSHEY_PLAIN, 2, (0, 255, 0), 3)

    cTime = time.time()
    fps = 1 / (cTime - pTime)
    pTime = cTime

    cv2.putText(img, f'FPS: {int(fps)}', (10, 30), cv2.FONT_HERSHEY_PLAIN, 2, (255, 0, 0), 3)
    cv2.imshow("Img", img)
    cv2.waitKey(1)
