# title         : member_checker.py
# description   : This is a tool for verifying OpenChat member against a member list
# author        : Eric
# date          : 10/03/2023
# version       : 1.0.0
####################################################################################

import cv2
import time
import requests

import numpy as np
import pandas as pd
import win32api as wa
import win32con as wc
import win32gui as wg
import win32process as wp
import imagehash as ih
import pyautogui as ag
import pytesseract as ts

from thefuzz import fuzz

# Constants
WINDOW_NAME = 'LINE'
# Setting 1
# WINDOW_SIZE = (288, 512)
# WINDOW_CROP = (60, 65, 3, 3)
# Setting 2
WINDOW_SIZE = (504, 896)
WINDOW_CROP = (100, 110, 5, 5)

SCROLL_TICKS = 3
SCROLL_SLEEP = 0.3
SCREENSHOT_SLEEP = 1
MAX_SCREENSHOTS = 20
FUZZY_THRESHOLD = 75
IGNORE_MEMBERS = {'', '管理員1', '管理員2'}
MEMBERS_API = 'https://docs.google.com/spreadsheets/d/e/2PACX-1vTixPZ1SIc1duIOOeMCF8a8x753GXLDzCAuVXRpSXQ9mtJQcb3tnSbJkLC38KdM6OXohcGLQMtRAZg3/pub?gid=620929327&single=true&output=csv'

# Settings
ts.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"


def win_enum_handler(hwnd, coord_handler):
    if wg.IsWindowVisible(hwnd) and not wg.IsIconic(hwnd) and wg.GetWindowText(hwnd) == WINDOW_NAME:
        window_rec = wg.GetWindowRect(hwnd)
        coord_handler[hwnd] = {
            'position': [window_rec[0], window_rec[1]],
            'size': [window_rec[2]-window_rec[0], window_rec[3]-window_rec[1]],
            'cropped': [
                window_rec[0]+WINDOW_CROP[0],
                window_rec[1]+WINDOW_CROP[1],
                window_rec[2]-window_rec[0]-WINDOW_CROP[0]-WINDOW_CROP[2],
                window_rec[3]-window_rec[1]-WINDOW_CROP[1]-WINDOW_CROP[3],
            ],
        }


def capture_screenshots():
    print("\nStart finding target window")
    app_coord = {}
    wg.EnumWindows(win_enum_handler, app_coord)

    target_coord = dict(filter(lambda x: x[1]['size'] == list(WINDOW_SIZE), app_coord.items()))
    if len(target_coord) != 1:
        raise Exception(f"Failed to find the target window with given size: {WINDOW_SIZE}")

    # Focus on the target window
    target_handle = list(target_coord.keys())[0]
    remote_thread, _ = wp.GetWindowThreadProcessId(target_handle)
    wp.AttachThreadInput(wa.GetCurrentThreadId(), remote_thread, True)
    wg.SetFocus(target_handle)
    target = list(target_coord.values())[0]

    # Move to the center of the target window
    target_center = target['position'][0]+target['size'][0]/2, target['position'][1]+target['size'][1]/2
    ag.moveTo(*target_center, duration=0.2)

    target_images = []
    target_hashes = []
    for i in range(MAX_SCREENSHOTS):
        # Get current screenshot and hash
        time.sleep(SCREENSHOT_SLEEP)
        current_image = ag.screenshot(region=target['cropped'])
        current_hash = ih.average_hash(current_image, 64)

        if len(target_hashes) >= 1 and current_hash-target_hashes[-1] < 5:
            print("Automatic scrolling completed\n")
            break
        else:
            print(f"  Taking screenshot {i+1:02d}")
            target_images.append(current_image)
            target_hashes.append(current_hash)

        # Scroll down N ticks
        for _ in range(SCROLL_TICKS):
            time.sleep(SCROLL_SLEEP)
            ag.scroll(-1)

    return target_images

def process_screenshots(screenshots):
    member_list = []
    print("Performing OCR on the screenshots")

    for idx, image in enumerate(screenshots):
        # Convert to grayscale
        print(f"  Processing screenshot {idx+1:02d}")
        image_size = image.size
        preprocessed = cv2.cvtColor(np.array(image), cv2.COLOR_BGR2GRAY)
        preprocessed = cv2.resize(preprocessed, list(map(lambda x: x*2, image_size)), interpolation=cv2.INTER_LANCZOS4)
        # cv2.imshow('pic', preprocessed)
        # cv2.waitKey(0)

        # Perform text extraction
        data = ts.image_to_string(
            preprocessed,
            lang='eng+chi_tra',
            config=r'--psm 6 configfile ./tesseract.conf',
        )
        # print(data)
        member_list.extend(map(lambda x: x.replace(' ', ''), data.split('\n')))

    print("Member list parsed from screenshot\n")
    return set(member_list)-IGNORE_MEMBERS

def fetch_members():
    print("Start fetching valid members")
    member_csv = pd.read_csv(MEMBERS_API)
    print("Valid members fetched\n")
    return set(member_csv['用戶名'].values.tolist())

def validate_members(members, valid_members):
    # Remove identical members
    print("Start validating members")
    remain_members = members - valid_members
    remain_valid = valid_members - members
    print("  Performing exact check")

    print("  Performing fuzzy check")
    fuzzy_pairs = []
    not_matched = []
    for valid in remain_valid:
        best_match = []
        for member in remain_members:
            score = fuzz.ratio(valid, member)
            if len(best_match) != 2 or score > best_match[1]:
                best_match = [member, score]

        if best_match[1] >= FUZZY_THRESHOLD:
            fuzzy_pairs.append((valid, best_match[0], best_match[1]))
            remain_members.remove(best_match[0])
        else:
            not_matched.append(valid)
    print("Validation completed")
    print("\n---")

    print("\nValidated data:")
    print(f"  Total: {len(valid_members)}")
    print(f"  Exact: {len(valid_members)-len(remain_valid)}")
    print(f"  Fuzzy: {len(fuzzy_pairs)}")
    for pair in sorted(fuzzy_pairs, key=lambda x: x[2], reverse=True):
        print(f"    {pair[0]:<15}\t/\t{pair[1]:<17}\tscore: {pair[2]:3d}")

    print("\nNon-validated data:")
    print(f"  Valid:  {sorted(not_matched)}")
    print(f"  Member: {sorted(remain_members)}")


if __name__ == '__main__':
    screenshots = capture_screenshots()
    members = process_screenshots(screenshots)
    valid_members = fetch_members()
    validate_members(members, valid_members)