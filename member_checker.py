# title         : member_checker.py
# description   : This is a tool for verifying OpenChat member against a member list
# author        : Eric
# date          : 10/03/2023
# version       : 1.0.0
####################################################################################

import os
import re
import cv2
import time

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
from datetime import datetime as dt

# Constants
WINDOW_NAME = 'LINE'
# Setting 1
# WINDOW_SIZE = (288, 512)
# WINDOW_CROP = (60, 65, 90, 3)
# TEXT_VERT = [15, 30]
# Setting 2
WINDOW_SIZE = (504, 896)
WINDOW_CROP = (100, 110, 150, 5)
TEXT_VERT = [30, 55]

SCROLL_TICKS = 6
SCROLL_SLEEP = 0.2
SCREENSHOT_SLEEP = 1
MAX_SCREENSHOTS = 20

COUNT_THRESHOLD = 4
VALIDATE_KEY = '用戶名'
CHAR_BLACKLIST = '[ \\n\\t\\/\\\\.*?:<>"|]'
FUZZY_THRESHOLD = 50

IGNORE_MEMBERS = {'管理員1', '管理員2'}
MEMBERS_API = 'https://docs.google.com/spreadsheets/d/e/2PACX-1vTixPZ1SIc1duIOOeMCF8a8x753GXLDzCAuVXRpSXQ9mtJQcb3tnSbJkLC38KdM6OXohcGLQMtRAZg3/pub?gid=620929327&single=true&output=csv'
EXPORT_ROOT = '.\\export\\'
EXPORT_TYPE = '.png'
TESSERACT_PATH = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
TESSERACT_CONF = r'.\tesseract.conf'

# Settings
ts.pytesseract.tesseract_cmd = TESSERACT_PATH


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


def get_hotspots(data, gap_threshold, count_threshold):
    previous = -99999
    groups = []
    for entry in sorted(data):
        if entry - previous <= gap_threshold/2:
            groups[-1].append(entry)
        else:
            groups.append([entry])
        previous = entry
    hotspots = []
    for group in groups:
        if len(group) >= count_threshold:
            hotspots.append(int(np.quantile(group, [0.75])[0]))
    return hotspots


def parse_text_center(data, shape):
    box_coord = []
    for entry in data.split('\n'):
        if entry:
            elements = entry.split(' ')
            box_coord.append({
                'key': elements[0],
                'rect': [
                    int(elements[1]),
                    shape[0] - int(elements[4]),
                    int(elements[3]),
                    shape[0] - int(elements[2]),
                ],
            })
    filtered_coord = map(lambda x: x['rect'][1], filter(
        lambda x: x['key'] != '-', box_coord))
    return get_hotspots(filtered_coord, sum(TEXT_VERT), COUNT_THRESHOLD)


def capture_screenshots():
    print("\nStart finding target window")
    app_coord = {}
    wg.EnumWindows(win_enum_handler, app_coord)

    target_coord = dict(
        filter(lambda x: x[1]['size'] == list(WINDOW_SIZE), app_coord.items()))
    if len(target_coord) != 1:
        raise Exception(
            f"Failed to find the target window with given size: {WINDOW_SIZE}, possible window sizes: {[c[1]['size'] for c in app_coord.items()]}")

    # Focus on the target window
    target_handle = list(target_coord.keys())[0]
    remote_thread, _ = wp.GetWindowThreadProcessId(target_handle)
    wp.AttachThreadInput(wa.GetCurrentThreadId(), remote_thread, True)
    wg.SetFocus(target_handle)
    target = list(target_coord.values())[0]

    # Move to the center of the target window
    target_center = target['position'][0]+target['size'][0] / \
        2, target['position'][1]+target['size'][1]/2
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
            ag.scroll(-100)

    return target_images


def process_screenshots(screenshots):
    print("Performing OCR on the screenshots")
    name_images = []
    for idx, image in enumerate(screenshots):
        # Convert to grayscale
        print(f"  Extracting screenshot {idx+1:02d}")
        image_size = image.size
        preprocessed = cv2.cvtColor(np.array(image), cv2.COLOR_BGR2GRAY)
        preprocessed = cv2.resize(preprocessed, list(
            map(lambda x: x*2, image_size)), interpolation=cv2.INTER_LANCZOS4)

        # Perform text extraction
        data = ts.image_to_boxes(
            preprocessed,
            lang='eng+chi_tra',
            config=r'--psm 6 -c tessedit_char_blacklist=£€¢!@#$%^&*()[]{}_+=/\\:;~\"',
        )
        text_centers = parse_text_center(data, preprocessed.shape[:2])
        for idx, center in enumerate(text_centers):
            name_images.append(
                preprocessed[center-TEXT_VERT[0]:center+TEXT_VERT[1], :])

    # Filter edge images
    name_images = list(filter(
        lambda x: x.shape[0] == TEXT_VERT[0]+TEXT_VERT[1], name_images))
    print(f"Total {len(name_images)} names extracted")

    member_list = []
    member_images = {}
    for idx, image in enumerate(name_images):
        if idx % 20 == 0:
            print(f"  Processing name {idx+1:03d} - {idx+20:03d}")
        parsed = ts.image_to_string(
            image,
            lang='eng+chi_tra',
            config=r'--psm 7 -c tessedit_char_blacklist=£€¢!@#$%^&*()[]{}_+=/\\:;~\"',
        )
        parsed = re.sub(CHAR_BLACKLIST, '', parsed)
        member_images[parsed] = cv2.imencode(EXPORT_TYPE, image)[1]
        member_list.append(parsed)
    member_set = set(member_list)-IGNORE_MEMBERS
    print(f"Total {len(member_set)} members parsed from screenshots\n")
    return {
        'members': member_set,
        'images': member_images,
    }


def fetch_members():
    print("Start fetching valid members")
    member_csv = pd.read_csv(MEMBERS_API)
    print("Valid members fetched\n")
    return set(member_csv[VALIDATE_KEY].values.tolist())


def validate_members(members, valid_members):
    # Remove identical members
    print("Start validating members")
    print("  Performing exact check")
    exact_matches = members['members'].intersection(valid_members)
    remain_members = members['members'] - valid_members
    remain_valid = valid_members - members['members']

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
            # member_images[member]
        else:
            not_matched.append(valid)

    print("  Exporting non-validated images")
    # Prepare export folder
    folder_name = dt.now().strftime('%Y%m%d_%H%M%S')

    member_images = members['images']
    export_path = os.path.join(EXPORT_ROOT, folder_name)
    fuzzy_path = os.path.join(export_path, 'fuzzy')
    remain_path = os.path.join(export_path, 'remain')
    os.mkdir(export_path)
    os.mkdir(fuzzy_path)
    os.mkdir(remain_path)
    for pair in fuzzy_pairs:
        member_images[pair[1]].tofile(os.path.join(
            fuzzy_path, f"{pair[2]:02d}_{pair[1]}_{pair[0]}{EXPORT_TYPE}"))
    for member in remain_members:
        member_images[member].tofile(os.path.join(
            remain_path, f"{member}{EXPORT_TYPE}"))

    print("Validation completed")
    print("\n---")

    print("\nValidated data:")
    print(f"  Total: {len(valid_members)}")
    print(f"  Exact: {len(exact_matches)}")
    for member in sorted(exact_matches):
        print(f"    {member}")
    print(f"  Fuzzy: {len(fuzzy_pairs)}")
    for pair in sorted(fuzzy_pairs, key=lambda x: x[2], reverse=True):
        print(f"    {pair[0]:<15}\t/\t{pair[1]:<17}\tscore: {pair[2]:02d}")

    print("\nNon-validated data:")
    print(f"  Valid:  {len(not_matched)}")
    for member in sorted(not_matched):
        print(f"    {member}")
    print(f"  Remain: {len(remain_members)}")
    for member in sorted(remain_members):
        print(f"    {member}")


if __name__ == '__main__':
    screenshots = capture_screenshots()
    members = process_screenshots(screenshots)
    valid_members = fetch_members()
    validate_members(members, valid_members)
