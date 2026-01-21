import os
import sys
import subprocess  # nosec B404
from subprocess import Popen  # nosec B404
import socket
import pathlib
import time
from datetime import datetime
import platform
import msvcrt
import ctypes
import win32com.client

from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.common.exceptions import WebDriverException

# from selenium.webdriver.chrome.service import Service
# from webdriver_manager.chrome import ChromeDriverManager

from colorist import ColorRGB
from colorama import just_fix_windows_console

# Fix for colors for CMD (win)
just_fix_windows_console()

# Global processes
mxp_pcsc_process: Popen[bytes]
browser: webdriver.Chrome

# Static required paths (change if it gets updated)
MXP_PCSC_PATH = "C:\\some path to MXP PCSC Keyboard.exe"
URL = "https://link-to-be-operated-on"

# Variables
OS_USERNAME = os.environ.get('USERNAME')
OS_NET_NAME = platform.node()
OS_NAME = platform.system()
OS_VER = platform.version()

# Task scheduler sevice
scheduler = win32com.client.gencache.EnsureDispatch('Schedule.Service')

# Colors
red = ColorRGB(255, 0, 0)
green = ColorRGB(0, 255, 0)
yellow = ColorRGB(229, 229, 16)
cyan = ColorRGB(0, 255, 255)
purple = ColorRGB(128, 0, 128)
RESET = ColorRGB.OFF


def disable_quickedit():
    """
    Disable quickedit mode on Windows terminal. quickedit prevents script to
    run without user pressing keys
    """
    if not os.name == 'posix':
        try:
            kernel32 = ctypes.WinDLL('kernel32', use_last_error=True)
            device = r'\\.\CONIN$'
            with open(device, 'r', encoding='utf-8') as con:
                h_con = msvcrt.get_osfhandle(con.fileno())
                kernel32.SetConsoleMode(h_con, 0x0080)
        except Exception as e:
            print(f'{red}[SYS ERR]:{RESET} Cannot disable QuickEdit mode ' + str(e))
            print(f'{yellow}[SYS INFO]:{RESET} The script might be automatically\
            paused on Windows terminal')


def show_task_scheduler_info():
    """
    Displays Task scheduler schedule for system reboot time
    """
    scheduler.Connect()
    folders = [scheduler.GetFolder('\\')]
    while folders:
        folder = folders.pop(0)
        folders += list(folder.GetFolders(0))
        for task in folder.GetTasks(0):
            if task.Name == "PC REBOOT":
                next_run_time = datetime.replace(task.NextRunTime, tzinfo=None)
                last_run_time = datetime.replace(task.LastRunTime, tzinfo=None)
                print(f"{yellow}[NEXT SYS REBOOT] : {next_run_time}{RESET}")
                print(f"{yellow}[LAST SYS REBOOT] : {last_run_time}{RESET}\n")
                break


def get_time() -> str:
    """
    Checks current local time and returns it.

    Returns:
        str: current formatted time
    """
    t = time.time()
    return time.strftime("%H:%M", time.localtime(t))


def get_date() -> str:
    """
    Checks current local date and returns it.

    Returns:
        str: current formatted date
    """
    return datetime.today().strftime('%Y-%m-%d %H-%M-%S')


def clear_console():
    """
    Clears all text in the console
    """
    os.system("cls||clear")  # nosec


def process(process_name: str) -> bool:
    """
    Checks processes image name if its still running or not.

    Args:
        process_name (str): image name for the process (e.g. 'Chrome.exe')

    Returns:
        bool: True if its still active/running, otherwise False
    """
    call = 'TASKLIST', '/FI', f'imagename eq {process_name}'
    # use buildin check_output right away
    output = subprocess.check_output(call).decode()  # nosec
    # check in last line for process name
    last_line = output.strip().rsplit('\r\n', maxsplit=1)[-1]
    # because Fail message could be translated
    return last_line.lower().startswith(process_name.lower())


def internet(host="8.8.8.8", port=53, timeout=3) -> bool:
    """
    Heartbeat for internet connection.

    Host: 8.8.8.8 (google-public-dns-a.google.com)
    OpenPort: 53/tcp
    Service: domain (DNS/TCP)
    """
    try:
        socket.setdefaulttimeout(timeout)
        socket.socket(socket.AF_INET, socket.SOCK_STREAM).connect((host, port))
        return True
    except socket.error as ex:
        print(f"{red}[NET ERR]: {ex}{RESET}")
        return False


def restart_all():
    """
    Restart all given processes and attempts to restart them.
    """
    print(f"{yellow}[SYS INFO]: SELF-RESTART INITIALIZED{RESET}")
    browser.quit()
    mxp_pcsc_process.kill()
    print(f"{green}[SYS OK]: PROCESSES SUCCESSFULLY TERMINATED{RESET}")
    print(f"{yellow}[SYS INFO]: RESTARTING REQUIRED PROCESSES{RESET}")
    time.sleep(5)
    main()


def err_handler(e: Exception | str):
    """
    | When an error is encountered, it will write all errors in a txt file,
    | then call func to restart all required tasks for TAR clock in and out to work properly.

    Args:
        e (Exception | str): any error exception
    """
    date_now = get_date()
    err_log_path = pathlib.Path(__file__).parent.resolve()
    try:
        with open(f"{err_log_path}\\err_logs\\{date_now} - error_log.txt", "w+", encoding="utf-8") as f:
            f.write(f"DATE & TIME: {date_now}\n")
            f.write("--------------------------------\n")
            f.write(f"[SYS ERR]: {str(e)}")
        print(f"{cyan}[SYS INFO]: ERROR LOGS CREATED >{RESET} {yellow}{err_log_path}\\err_logs\\{date_now} - error_log.txt{RESET}")
        browser.quit()
        mxp_pcsc_process.kill()
        time.sleep(5)
        os.execv(sys.executable, ['python'] + sys.argv)  # nosec
    except FileNotFoundError as be:
        print(f"{red}{be}{RESET}")
        os.makedirs(err_log_path.joinpath("err_logs"))
        print(f"{cyan}[SYS INFO]: ERR_LOGS FOLDER CREATED{RESET}")
        time.sleep(5)
        err_handler(e)
    print(f"{red}[SYS ERR]: {e}{RESET}")
    restart_all()


def show_sys_info():
    """
    Display basic system and platform info
    """
    now = get_time()
    print(f"{yellow}[TIME] - {now}{RESET}")
    print("--------------\n")
    print(f"{cyan}[OS NAME]    : {OS_NAME}{RESET}")
    print(f"{cyan}[OS VERSION] : {OS_VER}{RESET}")
    print(f"{cyan}[OS USER]    : {OS_USERNAME}{RESET}\n")
    show_task_scheduler_info()


def process_handler():
    """
    | Checks every 60 seconds if ```Chrome``` and ```MXP PCSC Keyboard``` processes are running,
    | if one of them is not (e.g. crashes, someone closes it), it attempt to terminate
    | the other process and rerun them again in order (MXP PCSC > Chrome) for proper
    | card reading functionallity.
    |
    | It also checks if device is connected to the internet and displays confirmation status,
    | if not, an error message will be displayed internet connection is lost and
    | the duration of how long has it been offline once it connects back to the internet.
    """
    attempt_num = 0
    while True:
        is_mxp_pcsc_running = process("MXP PCSC Keyboard.exe")
        is_browser_running = process("chrome.exe")
        if is_browser_running is False or is_mxp_pcsc_running is False:
            print(f"{red}[SYS ERR]: ONE OR MORE OF THE REQUIRED PROCESSES IS NOT RUNNING{RESET}")
            restart_all()

        is_device_connected = internet()
        if is_device_connected is True:
            if attempt_num > 0:
                print(f"{green}[NET STATUS]: ONLINE{RESET}")
                print(f"{cyan}[NET INFO]: TOTAL ATTEMPTS => {attempt_num - 1}{RESET}")
                attempt_num = 0
                restart_all()
            else:
                if browser.current_url != URL:
                    print(f"{red}[WEB ERR]: BAD URL{RESET}")
                    restart_all()
                clear_console()
                show_sys_info()
                print(f"{green}[NET PC NAME] : {OS_NET_NAME}{RESET}")
                print(f"{green}[NET STATUS]  : ONLINE{RESET}")
                print(f"{green}----------------------{RESET}")
                time.sleep(60)
        else:
            clear_console()
            show_sys_info()
            print(f"{green}[NET PC NAME] : {OS_NET_NAME}{RESET}")
            print(f"{yellow}[NET STATUS]  : OFFLINE{RESET}")
            print(f"{red}[NET ERR]     : NO INTERNET CONNECTION{RESET}")
            print(f"{red}[NET PROCESS] : ATTEMPTING TO RECONNECT{RESET}")
            print(f"{red}---------------------------------------{RESET}\n")
            print(f"{purple}[RECONNECTION ATTEMPTS]:{RESET} {yellow}{attempt_num}{RESET}")
            attempt_num += 1
            time.sleep(15)


def run_browser() -> bool:
    """
    | Opens ``Chrome`` browser and accepts appearing alert on page load,
    | then injects additional JS on inital page load to switch focus
    | from PIN input field to Employee ID field after given time.

        Additional browser arguments:
            --kiosk
            --fullscreen
            --disable-pinch

        Experimental options (disabling headliner info):
            - useAutomationExtension: False
            - excludeSwitches: ["enable-autmation"]

    """
    global browser
    browser_options = Options()
    browser_options.add_argument("--kiosk")
    browser_options.add_argument("--fullscreen")
    browser_options.add_argument("--disable-pinch")
    browser_options.add_experimental_option("useAutomationExtension", False)
    browser_options.add_experimental_option("excludeSwitches", ["enable-automation"])
    browser_options.add_experimental_option("detach", True)
    try:
        browser = webdriver.Chrome(options=browser_options)
        browser.get(URL)
        wait = WebDriverWait(browser, 30)
    except WebDriverException as wde:
        print(f"{red}[WEB ERR]: {wde}{RESET}")
        err_handler(wde)
        return False
    try:
        wait.until(EC.alert_is_present(), f'{red}[WEB ERR]: ALERT TIMED OUT{RESET}')
        alert = browser.switch_to.alert
        alert.accept()
        print(f"{purple}[WEB INFO]: ALERT ACCEPTED{RESET}")
        id_input = wait.until(EC.element_to_be_clickable((By.ID, 'edLoginUser')))
        if id_input:
            browser.execute_script('''
                const id_div = document.getElementById("edLoginUser")
                const pin_div = document.getElementById("edLoginPin")
                const close_btn = document.getElementById("ext-button-24")
                const id_input = document.getElementById("ext-element-1018")
                const pin_input = document.getElementById("ext-element-1026")
                function returnFocusOnIdInput() {
                    id_div.classList.add("x-field-focused")
                    id_div.classList.add("selected")
                    id_input.focus()
                    pin_div.classList.remove("x-field-focused")
                    pin_div.classList.remove("selected")
                    pin_input.blur()
                }
                close_btn.addEventListener("touchend", (e) => {
                    e.preventDefault();
                    setTimeout(() => {
                        returnFocusOnIdInput()
                    }, 1000)
                })
                close_btn.addEventListener("click", (e) => {
                    e.preventDefault();
                    setTimeout(() => {
                        returnFocusOnIdInput()
                    }, 1000)
                })
                setInterval(returnFocusOnIdInput, 60000)
            ''')
            return True
        return False
    except Exception as e:
        print(f"{red}[WEB ERR]: {e}{RESET}")
        err_handler(e)
        return False


def run_mxp_pcsc() -> int:
    """
    Opens ``MXP PCSC Keyboard.exe`` from temp folder for card reader functionality.

    Returns:
        int: returns processes current running code
    """
    global mxp_pcsc_process
    mxp_pcsc_process = Popen(MXP_PCSC_PATH)  # nosec
    return mxp_pcsc_process.returncode


def main():
    """
    | Init function once exe is booted.
    |
    | Runs ```MXP PCSC Keyboard.exe``` PCSC first for card reading functionality,
    | if it returns code 0 (only given when the program is executed and closed),
    | it will attempt to open browser and run process handler for regular checking.
    """
    try:
        disable_quickedit()
        return_code = run_mxp_pcsc()
        if return_code != 0:
            print(f"{green}[SYS OK]: MXP PCSC KEYBOARD SUCCESSFULLY RUNNING{RESET}")
            print(f"{cyan}[SYS INFO]: OPENING CHROME BROWSER{RESET}")
            is_open = run_browser()
            if is_open is True:
                process_handler()
            else:
                restart_all()

    except FileNotFoundError as fnfe:
        print(f"{red}[SYS ERR]: MXP PCSC KEYBOARD FAILED TO EXECUTE{RESET}")
        err_handler(fnfe)


main()
