# LINE Open Chat Automation

This tool is designed to automate the process of checking member names in LINE Open Chat. It compares the names of chat members with a predefined member list, ensuring accuracy and organization.

## Requirements

* Tesseract OCR
* Python 3.x
* Member list API

## Installation

1. Clone the repository:
    ```bash
    git clone https://github.com/weiyuericlee/open-chat-automation.git
    ```
2. Navigate to the project directory:
    ```bash
    cd open-chat-automation
    ```
3. Create Python venv with the name `py3_env`:
    ```bash
    python -m venv py3_env
    ```
4. Activate Python venv:
    ```bash
    ./py3_env/Scripts/Activate.ps1
    ```
5. Install required Python dependencies
    ```bash
    pip install -r requirements.txt
    ```
6. Download Tesseract OCR from:
    [Tesseract at UB Mannheim](https://github.com/UB-Mannheim/tesseract/wiki) or any preferred source
7. Install Tesseract OCR and keep note of the executable location (e.g. `C:\Program Files\Tesseract-OCR\tesseract.exe`)
8. Update the `TESSERACT_PATH` to the executable location

## Usage

1. Open `LINE for Windows`
2. Open the member page
3. Run the Python script with:
    ```bash
    python member_checker.py
    ```
4. The script will automatically scroll and take screenshots, then perform OCR on the list

## Notes

Tested in Python 3.11.4.

## License

Distributed under the MIT License. See `LICENSE.txt` for more information.

## Contact

* Eric Lee - weiyu.ericlee@gmail.com
* Ariel Huang - yinx8306@gmail.com
