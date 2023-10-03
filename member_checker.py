import time

import win32api as wa
import win32con as wc
import win32gui as wg
import win32process as wp
import imagehash as ih
import pyautogui as ag
import pytesseract as ts

WINDOW_NAME = 'LINE'
WINDOW_SIZE = (504, 896)
WINDOW_CROP = (100, 110, 5, 5)
SCROLL_TICKS = 3
SCROLL_SLEEP = 0.3
SCREENSHOT_SLEEP = 1
MAX_SCREENSHOTS = 20

def winEnumHandler(hwnd, coordHandler):
    if wg.IsWindowVisible(hwnd) and not wg.IsIconic(hwnd) and wg.GetWindowText(hwnd) == WINDOW_NAME:
        windowRec = wg.GetWindowRect(hwnd)
        coordHandler[hwnd] = {
            'position': [windowRec[0], windowRec[1]],
            'size': [windowRec[2]-windowRec[0], windowRec[3]-windowRec[1]],
            'cropped': [
                windowRec[0]+WINDOW_CROP[0],
                windowRec[1]+WINDOW_CROP[1],
                windowRec[2]-windowRec[0]-WINDOW_CROP[0]-WINDOW_CROP[2],
                windowRec[3]-windowRec[1]-WINDOW_CROP[1]-WINDOW_CROP[3],
            ],
        }

appCoord = {}
wg.EnumWindows(winEnumHandler, appCoord)

targetCoord = dict(filter(lambda x: x[1]['size'] == list(WINDOW_SIZE), appCoord.items()))
if len(targetCoord) != 1:
    raise Exception(f"Failed to find the target window with given size: {WINDOW_SIZE}")

# Focus on the target window
targetHandle = list(targetCoord.keys())[0]
remoteThread, _ = wp.GetWindowThreadProcessId(targetHandle)
wp.AttachThreadInput(wa.GetCurrentThreadId(), remoteThread, True)
wg.SetFocus(targetHandle)
target = list(targetCoord.values())[0]

# Move to the center of the target window
targetCenter = target['position'][0]+target['size'][0]/2, target['position'][1]+target['size'][1]/2
ag.moveTo(*targetCenter, duration=0.5)

targetImgs = []
targetHashes = []
for i in range(MAX_SCREENSHOTS):
    # Get current screenshot and hash
    time.sleep(SCREENSHOT_SLEEP)
    currentImg = ag.screenshot(region=target['cropped'])
    currentHash = ih.colorhash(currentImg)

    if len(targetHashes) >= 1 and currentHash-targetHashes[-1] == 0:
        print("Scroll end detected")
        break
    else:
        print(f"Taking screenshot {i:02d}")
        targetImgs.append(currentImg)
        targetHashes.append(currentHash)

    # Scroll down N ticks
    for s in range(SCROLL_TICKS):
        time.sleep(SCROLL_SLEEP)
        ag.scroll(-1)

# for img in targetImgs:
#     img.show()
