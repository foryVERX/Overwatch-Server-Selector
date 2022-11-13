from tkinter import *
from tkinter.font import Font
from tkinter import ttk, filedialog
import win32com.shell.shell as shell
from subprocess import run
from PIL import ImageTk, Image
from io import BytesIO
import pic2str
import base64
import ctypes
from os.path import exists, isdir, isfile, join
from os import getenv, path, mkdir, listdir, linesep, startfile, remove
import webbrowser
import socket
import threading
import requests
import logging
import datetime
from urllib3 import Retry
from requests.adapters import HTTPAdapter

__version__ = '5.1.0'

# Create window object
app = Tk()
# Set Properties
app.title('MINA Overwatch 2 Server Selector')
app.resizable(False, False)
app.geometry('500x600')
app.configure(bg='#282828')
frame = Frame(app, width=500, height=94)
frame.pack()
frame.place(x=-2, y=0)

# For images
# Load byte data
byte_LOGO_SMALL_APPLICATION = base64.b64decode(pic2str.LOGO_SMALL_APPLICATION)
byte_SQUARE_BACKGROUND_MINA_TEST = base64.b64decode(pic2str.SQUARE_BACKGROUND_MINA_TEST)
byte_play_on_eu = base64.b64decode(pic2str.play_on_eu)
byte_programmable_button = base64.b64decode(pic2str.PROGRAMABLE_BUTTON)
byte_play_on_na_east = base64.b64decode(pic2str.play_on_na_east)
byte_play_on_na_west = base64.b64decode(pic2str.play_on_na_west)
byte_BLOCK_MIDDLE_EAST = base64.b64decode(pic2str.BLOCK_MIDDLE_EAST)
byte_play_on_australia = base64.b64decode(pic2str.play_on_australia)
byte_donation = base64.b64decode(pic2str.Donation_Button)
byte_CUSTOM_SETTINGS = base64.b64decode(pic2str.CUSTOM_SETTINGS)
byte_UNBLOCK_ALL_MAIN = base64.b64decode(pic2str.UNBLOCK_ALL_MAIN)
byte_RESET_BUTTON = base64.b64decode(pic2str.RESET_BUTTON)
byte_OPEN_IP_LIST_BUTTON = base64.b64decode(pic2str.OPEN_IP_LIST_BUTTON)
byte_APPLY_BUTTON = base64.b64decode(pic2str.APPLY_BUTTON)
byte_CUSTOM_SETTINGS_BACKGROUND = base64.b64decode(pic2str.CUSTOM_SETTINGS_BACKGROUND)

image_data_SMALL_APPLICATION = BytesIO(byte_LOGO_SMALL_APPLICATION)
image_data_SQUARE_BACKGROUND_MINA_TEST = BytesIO(byte_SQUARE_BACKGROUND_MINA_TEST)
image_play_on_eu = BytesIO(byte_play_on_eu)
image_programmable_button = BytesIO(byte_programmable_button)
image_play_on_na_east = BytesIO(byte_play_on_na_east)
image_play_on_na_west = BytesIO(byte_play_on_na_west)
image_BLOCK_MIDDLE_EAST = BytesIO(byte_BLOCK_MIDDLE_EAST)
image_play_on_australia = BytesIO(byte_play_on_australia)
image_donation = BytesIO(byte_donation)
image_CUSTOM_SETTINGS = BytesIO(byte_CUSTOM_SETTINGS)
image_UNBLOCK_ALL_MAIN = BytesIO(byte_UNBLOCK_ALL_MAIN)
image_data_RESET_BUTTON = BytesIO(byte_RESET_BUTTON)
image_data_OPEN_IP_LIST_BUTTON = BytesIO(byte_OPEN_IP_LIST_BUTTON)
image_data_APPLY_BUTTON = BytesIO(byte_APPLY_BUTTON)
image_data_CUSTOM_SETTINGS_BACKGROUND = BytesIO(byte_CUSTOM_SETTINGS_BACKGROUND)

# Add images
background = ImageTk.PhotoImage(Image.open(image_data_SQUARE_BACKGROUND_MINA_TEST))
logo = Label(frame, image=background)
logo.pack()

CUSTOM_SETTINGS_BACKGROUND = ImageTk.PhotoImage(Image.open(image_data_CUSTOM_SETTINGS_BACKGROUND))

button_img_RESET_BUTTON = ImageTk.PhotoImage(Image.open(image_data_RESET_BUTTON))
button_img_OPEN_IP_LIST_BUTTON = ImageTk.PhotoImage(Image.open(image_data_OPEN_IP_LIST_BUTTON))
button_img_APPLY_BUTTON = ImageTk.PhotoImage(Image.open(image_data_APPLY_BUTTON))
button_img_programmable_button = ImageTk.PhotoImage(Image.open(image_programmable_button))
button_img_ME = ImageTk.PhotoImage(Image.open(image_BLOCK_MIDDLE_EAST))
button_img_EU = ImageTk.PhotoImage(Image.open(image_play_on_eu))
button_img_NA_WEST = ImageTk.PhotoImage(Image.open(image_play_on_na_west))
button_img_ME_EAST = ImageTk.PhotoImage(Image.open(image_play_on_na_east))
button_img_Australia = ImageTk.PhotoImage(Image.open(image_play_on_australia))
button_img_donation = ImageTk.PhotoImage(Image.open(image_donation))
button_img_Default = ImageTk.PhotoImage(Image.open(image_UNBLOCK_ALL_MAIN))
button_img_CUSTOM_SETTINGS = ImageTk.PhotoImage(Image.open(image_CUSTOM_SETTINGS))

# Add font
futrabook_font = Font(family="Futura PT Demi", size=10)

# Global variables
localappdata_path = getenv('APPDATA') + '\\OverwatchServerBlocker'
ip_version_path = localappdata_path + '\\IP_version.txt'
overwatch_path = 'C:\\Program Files (x86)\\Overwatch\\_retail_\\Overwatch.exe'
customConfig_path = localappdata_path + '\\customConfig.txt'
ip_version_url = 'https://raw.githubusercontent.com/foryVERX/Overwatch-Server-Selector/main/ip_lists/IP_version.txt'
Ip_ranges_ME_url = 'https://raw.githubusercontent.com/foryVERX/Overwatch-Server-Selector/main/ip_lists/Ip_ranges_ME.txt'
Ip_ranges_EU_url = 'https://raw.githubusercontent.com/foryVERX/Overwatch-Server-Selector/main/ip_lists/Ip_ranges_EU.txt'
Ip_ranges_NA_East_url = 'https://raw.githubusercontent.com/foryVERX/Overwatch-Server-Selector/main/ip_lists/Ip_ranges_NA_East.txt'
Ip_ranges_NA_central_url = 'https://raw.githubusercontent.com/foryVERX/Overwatch-Server-Selector/main/ip_lists/Ip_ranges_NA_central.txt'
Ip_ranges_NA_West_url = 'https://raw.githubusercontent.com/foryVERX/Overwatch-Server-Selector/main/ip_lists/Ip_ranges_NA_West.txt'
Ip_ranges_AS_Korea_url = 'https://raw.githubusercontent.com/foryVERX/Overwatch-Server-Selector/main/ip_lists/Ip_ranges_AS_Korea.txt'
Ip_ranges_AS_1_url = 'https://raw.githubusercontent.com/foryVERX/Overwatch-Server-Selector/main/ip_lists/Ip_ranges_AS_1.txt'
Ip_ranges_AS_Singapore_url = 'https://raw.githubusercontent.com/foryVERX/Overwatch-Server-Selector/main/ip_lists/Ip_ranges_AS_Singapore.txt'
Ip_ranges_AS_Taiwan_url = 'https://raw.githubusercontent.com/foryVERX/Overwatch-Server-Selector/main/ip_lists/Ip_ranges_AS_Taiwan.txt'
Ip_ranges_AS_Japan_url = 'https://raw.githubusercontent.com/foryVERX/Overwatch-Server-Selector/main/ip_lists/Ip_ranges_AS_Japan.txt'
Ip_ranges_Australia_url = 'https://raw.githubusercontent.com/foryVERX/Overwatch-Server-Selector/main/ip_lists/Ip_ranges_Australia.txt'
Ip_ranges_Brazil_url = 'https://raw.githubusercontent.com/foryVERX/Overwatch-Server-Selector/main/ip_lists/Ip_ranges_Brazil.txt'
BlockingConfig_url = 'https://raw.githubusercontent.com/foryVERX/Overwatch-Server-Selector/main/ip_lists/BlockingConfig.txt'

updating_state = False
internet_initialization = False
sorter_initialization = False
checkForUpdate_initialization = False
tunnel_option = False
isUpdated = ''
internetConnection = ''
update_time = 0

# IP Ranges
Ip_ranges_dic = {}
blockingConfigDic = {}
ip_range_checkboxes = {}
ip_ranges_files = []
customIpRanges = []
customConfig = []

# Logger
if not exists(localappdata_path):
    mkdir(localappdata_path)  # Make directory
logging.basicConfig(
    level=logging.DEBUG,
    filename=localappdata_path + '\\OVERWATCH SERVER SELECTOR LOG.log',
    filemode='w',
    format='[%(asctime)s] %(levelname)s - %(message)s',
    datefmt='%H:%M:%S'
)


# Functions
def check_admin():
    if ctypes.windll.shell32.IsUserAnAdmin() == 0:
        adminLabel.config(text='Restart(run as administrator)', bg='#282828', fg='#ef2626', font=futrabook_font)
        logging.info("USER IS NOT AN ADMIN")
    else:
        adminLabel.config(text='Running as administrator', bg='#282828', fg='#26ef4c', font=futrabook_font)
        logging.info("USER IS ADMIN")


def is_connect():
    global internetConnection, internet_initialization
    try:
        socket.create_connection(("www.google.com", 80))
        internetConnection = True
    except OSError:
        internetConnection = False


def updateIp():
    global updating_state, sorter_initialization, checkForUpdate_initialization
    controlButtons('disabled')
    if internetConnection:
        updating_state = True
        update_text = "UPDATING..."
        internetLabel.config(text=update_text, fg='#ddee4a')
        msg_fail = "CONNECTION FAILED... Trying to update"
        start = datetime.datetime.now()
        with requests.Session() as s:  # Create a session and use it for all requests
            adapter = HTTPAdapter(max_retries=Retry(total=4, backoff_factor=1, allowed_methods=None,
                                                    status_forcelist=[429, 500, 502, 503, 504]))
            # This is to allow retry on a new connection if it fails
            s.mount("http://", adapter)
            s.mount("https://", adapter)
            createTextFile('BlockingConfig', request_raw_file(BlockingConfig_url, msg_fail, s), progressbar=True)
            createTextFile('Ip_ranges_EU', request_raw_file(Ip_ranges_EU_url, msg_fail, s), progressbar=True)
            createTextFile('Ip_ranges_ME', request_raw_file(Ip_ranges_ME_url, msg_fail, s), progressbar=True)
            createTextFile('Ip_ranges_AS_Singapore', request_raw_file(Ip_ranges_AS_Singapore_url, msg_fail, s),
                           progressbar=True)
            createTextFile('Ip_ranges_Australia', request_raw_file(Ip_ranges_Australia_url, msg_fail, s),
                           progressbar=True)
            createTextFile('Ip_ranges_Brazil', request_raw_file(Ip_ranges_Brazil_url, msg_fail, s), progressbar=True)
            createTextFile('Ip_ranges_NA_East', request_raw_file(Ip_ranges_NA_East_url, msg_fail, s), progressbar=True)
            createTextFile('Ip_ranges_NA_West', request_raw_file(Ip_ranges_NA_West_url, msg_fail, s), progressbar=True)
            createTextFile('Ip_ranges_NA_central', request_raw_file(Ip_ranges_NA_central_url, msg_fail, s),
                           progressbar=True)
            createTextFile('Ip_ranges_AS_Japan', request_raw_file(Ip_ranges_AS_Japan_url, msg_fail, s),
                           progressbar=True)
            createTextFile('Ip_ranges_AS_1', request_raw_file(Ip_ranges_AS_1_url, msg_fail, s), progressbar=True)
            createTextFile('Ip_ranges_AS_Taiwan', request_raw_file(Ip_ranges_AS_Taiwan_url, msg_fail, s),
                           progressbar=True)
            createTextFile('Ip_ranges_AS_Korea', request_raw_file(Ip_ranges_AS_Korea_url, msg_fail, s),
                           progressbar=True)
            createTextFile('IP_version', request_raw_file(ip_version_url, msg_fail, s), progressbar=True)
        finish = datetime.datetime.now() - start
        logging.info('UPDATING TOOK: ' + str(finish))
        progressBar.lower()
        logging.info("UPDATED")
        updating_state = False
        checkForUpdate_initialization = False
        ipSorter()
    else:
        update_text = "Please check your internet connection to download servers ip"
        internetLabel.config(text=update_text, fg='#ddee4a')


def check_ip_update():  # A function called at the start of the program to check for update
    global isUpdated, updating_state, checkForUpdate_initialization
    if updating_state:
        return
    updating_state = True
    is_connect()
    logging.info("Checking for updates")
    if internetConnection:
        if isdir(localappdata_path) and path.exists(ip_version_path):
            logging.info(localappdata_path + ' FOUND')
            logging.info(ip_version_path + ' FOUND')
            with open(ip_version_path, "r") as reader:  # Read Ip_version.txt from GitHub and analyze
                logging.info('Reading Ip_version.txt')
                for line in reader.readlines():
                    if len(line) > 1:
                        msg_fail = "Update check failed"
                        with requests.Session() as s:
                            adapter = HTTPAdapter(max_retries=Retry(total=4, backoff_factor=1, allowed_methods=None,
                                                                    status_forcelist=[429, 500, 502, 503, 504]))
                            s.mount("http://", adapter)
                            s.mount("https://", adapter)
                            ip_version_request = request_raw_file(ip_version_url, msg_fail, s)
                        ip_version_request = linesep.join([s for s in ip_version_request.splitlines() if s])
                        line = linesep.join([s for s in line.splitlines() if s])
                        logging.debug("VERSION REQUEST DURING CHECK FOR UPDATE: " + str(ip_version_request))
                        if ip_version_request == line:
                            update_text = "UPDATED"
                            app.after(250, internetLabel.config(text=update_text, fg='#26ef4c'))
                        else:
                            update_text = "NOT UPDATED"
                            app.after(250, internetLabel.config(text=update_text, fg='#ef2626'))
                            threading.Thread(target=updateIp, daemon=True).start()
        else:  # Make directory and call updateIp
            update_text = "FIRST TIME RUNNING.. UPDATING"
            app.after(250, internetLabel.config(text=update_text, fg='#ddee4a'))
            if not exists(localappdata_path):
                mkdir(localappdata_path)  # Make directory
            threading.Thread(target=updateIp, daemon=True).start()
    else:
        logging.debug('No internet connection')
        if not exists(ip_version_path):
            controlButtons('disabled')
            update_text = "CONNECTION FAILED... Trying to update"
            app.after(250, internetLabel.config(text=update_text, fg='#ef2626'))
            app.after(1000, check_ip_update)
        else:
            update_text = "NO INTERNET MIGHT BE NOT LATEST IP LIST VERSION"
            app.after(250, internetLabel.config(text=update_text, fg='#ddee4a'))
    updating_state = False
    # if thread_type == 'mainThread':  # If the function is called from main thread call it again after 5 mints
    # app.after(5000 * 60, checkUpdate)


def check_version_update():
    pass


def request_raw_file(url, msg_fail, s):
    try:
        # user-agent is just to trick the website that you are using a browser
        headers = {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:66.0) Gecko/20100101 Firefox/66.0"
        }
        r = s.get(url, headers=headers).content.decode('utf-8')
    except requests.exceptions.RequestException as e:  # This is the correct syntax
        logging.debug('URL REQUEST FAIL RETRYING: ' + url)
        logging.debug(str(e))
        raise SystemExit(e)
    return r


def createTextFile(file_name, contents, progressbar=False):
    # Contents can be a string or a list of strings must end with \n except last element in the list
    with open(localappdata_path + '\\' + file_name + '.txt', "w") as text_file:
        if isinstance(contents, list):
            text_file.writelines(contents)
        else:
            text_file.write(contents)
    if progressbar:
        progressBar.place(x=185, y=520)
        progressBar['value'] += 10


def ipSorter():  # Store ip ranges from Ip_ranges_....txt into Ip_ranges dictionary
    global sorter_initialization
    userConfigSorter()
    if exists(localappdata_path) and exists(ip_version_path):  # If those paths exists it means user updated
        logging.info(localappdata_path + ' FOUND')
        logging.info(ip_version_path + ' FOUND')
        controlButtons('disabled')
        servers_files = listdir(localappdata_path)
        for server in servers_files:
            if server.startswith("Ip_ranges"):
                blockingConfig(path.splitext(server)[0])
                server_path = localappdata_path + '\\' + server
                with open(server_path, "r") as reader:
                    temp_list = []
                    for line in reader.readlines():
                        line = line.strip('\n')
                        if len(line) > 5:
                            temp_list.append(line)
                            temp_list.append(',')
                    Ip_ranges_dic[server] = temp_list
        update_text = "UPDATED"
        app.after(250, internetLabel.config(text=update_text, fg='#26ef4c'))
        sorter_initialization = True
        logging.info("IP LIST SORTED")
    else:  # User running first time
        check_ip_update()
    controlButtons('normal')


def userConfigSorter():
    global customConfig
    customConfig.clear()
    if exists(customConfig_path):  # Handle custom config created by user
        logging.debug(customConfig_path + "FOUND")
        with open(customConfig_path, "r") as filenames:
            for line in filenames.readlines():
                line = line.strip('\n')
                if len(line) > 0:
                    if exists(localappdata_path + "\\" + line):
                        with open(localappdata_path + "\\" + line, "r") as reader:
                            for ip_range in reader.readlines():
                                ip_range = ip_range.strip('\n')
                                if len(ip_range) > 0:
                                    customConfig.append(ip_range)
                                    customConfig.append(',')


def iconMaker():  # Used to check if there is an icon in the same directory or not it will create the icon if not.
    if exists("LOGO_SMALL_APPLICATION.ico"):
        app.iconbitmap("LOGO_SMALL_APPLICATION.ico")
    else:
        icon = Image.open(image_data_SMALL_APPLICATION)
        icon.save("LOGO_SMALL_APPLICATION.ico")
        app.iconbitmap("LOGO_SMALL_APPLICATION.ico")


def controlButtons(command):  # 'disabled' or 'normal' buttons
    PlayMEButton['state'] = command

    ProgrammableButton['state'] = command

    PlayEUButton['state'] = command

    PlayNAWESTButton['state'] = command

    PlayNAEASTButton['state'] = command

    PlayAustraliaButton['state'] = command

    ClearBlocksButton['state'] = command

    DonationButton['state'] = command


def ruleMakerBlock(server_exception, np_ips, block_exception=True, rule_name='@Overwatch Block'):
    # Used to block IP range
    # np_ips is number of ip ranges that included in one function
    # server_exception is the only server to not block can be a list or string
    # If block_exception set to false then the server_exception is blocked ONLY
    controlButtons('disabled')
    x = 0
    temp_ip_ranges = list()
    size_of_ip_range = 0
    if not block_exception:  # Incase we want to block server_exception
        for server in Ip_ranges_dic:
            if server.strip('.txt') in str(server_exception):
                for ip in Ip_ranges_dic[server]:
                    temp_ip_ranges.append(ip)
                blockIpRange(temp_ip_ranges, rule_name)
                logging.info("One rule created")
                checkIfActive()
                controlButtons('normal')
                return
    for server in Ip_ranges_dic:  # Collect all IP ranges except the excluded
        if server.strip('.txt') not in str(server_exception):
            for ip in Ip_ranges_dic[server]:
                temp_ip_ranges.append(ip)
                size_of_ip_range += 1
    size_of_ip_range = int(size_of_ip_range / 2)
    logging.info("Total number of ip ranges: " + str(size_of_ip_range))
    if size_of_ip_range <= np_ips:  # Incase block command fits size limit we create one rule only
        blockIpRange(temp_ip_ranges, rule_name)
        checkIfActive()
        controlButtons('normal')
        logging.info("One rule created")
    else:  # Incase the command gets too long due to number of ip ranges, we slice the ip ranges
        temp_ip_ranges.clear()
        full_ip_ranges = list()
        for indexServer, server in enumerate(Ip_ranges_dic):
            if server.strip('.txt') not in str(server_exception):
                logging.info("ruleMakerBlock |" + str(server) + "is not: " + str(server_exception))
                for indexIp, ip in enumerate(Ip_ranges_dic[server]):
                    full_ip_ranges.append(ip)
        for ip in full_ip_ranges:
            temp_ip_ranges.append(ip)
            if (len(temp_ip_ranges)) == np_ips * 2:
                x += 1
                full_ip_ranges = full_ip_ranges[len(temp_ip_ranges):]
                blockIpRange(temp_ip_ranges, rule_name)
                temp_ip_ranges.clear()
            elif (len(full_ip_ranges) / 2) < np_ips and not len(full_ip_ranges) == 0:
                x += 1
                blockIpRange(full_ip_ranges, rule_name)
                logging.info('Last List')
                logging.info(str(x) + " --- Parsed Rules passed to blockIpRange function")
                checkIfActive()
                controlButtons('normal')
                return


def blockIpRange(ip_list, rule_name):
    ip_string = ''.join(ip_list)
    if ip_string[1:] == ',':
        ip_string = ip_string[1:]
    if ip_string[len(ip_string) - 1] == ',':
        ip_string = ip_string[:-1]
    logging.info("BLOCKING IP RANGES")
    if not ip_string == '':
        if tunnel_option:
            program = ' program=' + '"' + overwatch_path + '"'
        else:
            program = ''
        commands = 'advfirewall firewall add rule name="' \
                   + rule_name + '"' + program + \
                   ' Dir=In Action=Block RemoteIP=' \
                   + ip_string
        shell.ShellExecuteEx(lpVerb='runas', lpFile='netsh.exe', lpParameters=commands)
        commands = 'advfirewall firewall add rule name="' \
                   + rule_name + '"' + program + \
                   ' Dir=Out Action=Block RemoteIP=' \
                   + ip_string
        shell.ShellExecuteEx(lpVerb='runas', lpFile='netsh.exe', lpParameters=commands)
        if len(commands) > 8150:
            logging.debug("Command is too long")


def ruleDelete(rule_name):  # Delete rule by exact name, name must be a string '' or list of strings
    logging.info("DELETING RULES: " + str(rule_name))
    if type(rule_name) == list:
        for rule in rule_name:
            rule = '"' + rule + '"'
            commands = 'advfirewall firewall delete rule name = ' + rule
            shell.ShellExecuteEx(lpVerb='runas', lpFile='netsh.exe', lpParameters=commands)
    else:
        rule_name = '"' + rule_name + '"'
        commands = 'advfirewall firewall delete rule name = ' + rule_name
        shell.ShellExecuteEx(lpVerb='runas', lpFile='netsh.exe', lpParameters=commands)


def checkIfActive():  # To check if server is blocked or not
    servers_active_rule_list = ['ME_OW_SERVER_BLOCKER', 'NAEAST_OW_SERVER_BLOCKER', 'NAWEST_OW_SERVER_BLOCKER',
                                'EU_OW_SERVER_BLOCKER', 'AU_OW_SERVER_BLOCKER', 'Australia_OW_SERVER_BLOCKER',
                                "@CUSTOM_BLOCK"]
    CREATE_NO_WINDOW = 0x08000000
    command_list = ['netsh', 'advfirewall', 'firewall', 'show', 'rule', 'name=all']
    output = run(command_list, capture_output=True, text=True, creationflags=CREATE_NO_WINDOW)
    output = str(output.stdout)
    logging.info("CHECKING IF PREVIOUS RULE IS ACTIVE")
    for rule_name in servers_active_rule_list:
        rules_existence = output.find(rule_name)
        if rules_existence > 0:
            filtered = rule_name.rpartition('_')[0].replace(" ", "").replace("RuleName:@", "").replace("_OW_SERVER",
                                                                                                       "")
            if filtered == 'ME':
                blockingLabel.config(text='ME BLOCKED', bg='#282828', fg='#ef2626', font=futrabook_font)
                return
            if filtered == '@CUSTOM':
                blockingLabel.config(text="CUSTOM BLOCK", bg='#282828', fg='#26ef4c',
                                     font=futrabook_font)
                return
            else:
                if len(filtered) < 8:
                    filtered = filtered[0:2] + ' ' + filtered[2:]
                label_text = 'PLAYING ON ' + filtered
                blockingLabel.config(text=label_text, bg='#282828', fg='#26ef4c',
                                     font=futrabook_font)
                return
    blockingLabel.config(text='ALL UNBLOCKED (DEFAULT SETTINGS)', fg='#ddee4a')


def tunnel():  # Handle tunnelling options for Overwatch.exe
    global overwatch_path, tunnel_option
    checkbotton_state = tunnelCheckBox_state.get()
    logging.info("Check buttons state =  " + str(checkbotton_state))
    if checkbotton_state == 1:
        if exists(overwatch_path):
            logging.info("Game detected")
            createTextFile('Options', 'Tunnel=True', False)
            tunnel_option = True
        else:
            app.overwatch = filedialog.askopenfilename(initialdir='C:\\',
                                                       title='Select Overwatch\_retail_\Overwatch.exe ',
                                                       filetypes=(("Select Overwatch\_retail_\Overwatch.exe",
                                                                   "Overwatch.exe"),))
            existance_overwatch = app.overwatch.find("/_retail_/Overwatch.exe")
            if existance_overwatch > 0:
                overwatch_path = app.overwatch.replace('/', "\\")
                tunnel_option = True
                logging.debug("Overwatch path is:  " + overwatch_path)
                createTextFile('Options', ['Tunnel=True\n', overwatch_path], False)
            else:
                createTextFile('Options', 'Tunnel=False', False)
                tunnelCheckBox_state.set(0)
    else:
        createTextFile('Options', 'Tunnel=False', False)
        tunnel_option = False


def checkOptions():
    global tunnel_option, overwatch_path
    if exists(localappdata_path + '\\Options.txt'):
        with open(localappdata_path + '\\Options.txt', "r") as reader:
            options = reader.readlines()
        if options[0].strip() == 'Tunnel=True':
            tunnel_option = True
        if len(options) > 1:
            temp_path = options[1].strip()
            overwatch_path = temp_path
            tunnel_option = True

    else:
        tunnel_option = False
    return tunnel_option


def blockingConfig(server_name):
    global blockingConfigDic
    temp_block_config_list = []
    if exists(localappdata_path + '\\BlockingConfig.txt'):
        with open(localappdata_path + '\\BlockingConfig.txt', "r") as reader:
            for line in reader.readlines():
                if '@' in line:
                    if 'ipRangeName::' + server_name in line:
                        indexes = [pos for pos, char in enumerate(line) if char == "@"]
                        if len(indexes) > 1:
                            for i in range(len(indexes)):
                                if i != len(indexes) - 1:
                                    temp_block_config_list.append(
                                        line[indexes[i]:indexes[i + 1]].strip('\n').strip('@'))
                                else:
                                    temp_block_config_list.append(line[indexes[-i]:].strip('\n').strip('@'))
                        else:
                            temp_block_config_list.append(line[indexes[0]:].strip('\n').strip('@'))
                        logging.info("Blocking Config: " + 'Play on: ' + str(server_name) + " wants to exclude " +
                                     str(temp_block_config_list) +
                                     " From Blocking")
                        blockingConfigDic[server_name] = temp_block_config_list


def customSettingsWindow():
    global ip_ranges_files, ip_range_checkboxes, top
    savedSettings = []
    if exists(customConfig_path):
        with open(customConfig_path, 'r') as reader:
            for line in reader.readlines():
                if len(line) > 0:
                    savedSettings.append(line.strip())
    top = Toplevel()
    # Set Properties
    top.title('Custom Config')
    top.resizable(False, False)
    top.geometry('300x500')
    top.configure(bg='#404040')
    frameTop = Frame(top, width=300, height=500)
    frameTop.pack()
    frameTop.place(x=-2, y=0)
    backgroundTop = Label(frameTop, image=CUSTOM_SETTINGS_BACKGROUND)
    backgroundTop.pack()
    top.iconbitmap("LOGO_SMALL_APPLICATION.ico")
    # Buttons
    applyButton = Button(top, image=button_img_APPLY_BUTTON, font=futrabook_font, command=apply,
                         bg='#404040', fg='#404040', borderwidth=0, activebackground='#404040')
    applyButton.place(x=177 + 55, y=448, anchor="center")
    resetButton = Button(top, image=button_img_RESET_BUTTON, font=futrabook_font, command=resetCustomSettings,
                         bg='#404040', fg='#404040', borderwidth=0, activebackground='#404040')
    resetButton.place(x=11 + 55, y=448, anchor="center")
    openIpListButton = Button(top, image=button_img_OPEN_IP_LIST_BUTTON, font=futrabook_font, command=openListFolder,
                              bg='#404040', fg='#404040', borderwidth=0, activebackground='#404040')
    openIpListButton.place(x=92 + 55, y=479, anchor="center")
    onlyfiles = [f for f in listdir(localappdata_path) if isfile(join(localappdata_path, f))]
    ip_ranges_files = list()
    integersList = list()
    for file in onlyfiles:
        if file.startswith('Ip_ranges') or file.startswith('usercfg'):
            ip_ranges_files.append(file)
    for i in range(0, len(ip_ranges_files)):
        integerVariable = IntVar()
        integersList.append(integerVariable)
    ip_range_checkboxes = dict(zip(ip_ranges_files, integersList))
    for index, RANGE in enumerate(ip_range_checkboxes):
        ip_range_checkboxes[RANGE] = IntVar()
        chk = Checkbutton(top, text=RANGE[:-4], font=futrabook_font,
                          activebackground='#ddee4a',
                          bg='#404040', fg='#26ef4c', borderwidth=0, variable=ip_range_checkboxes[RANGE], width=200,
                          anchor="w", selectcolor='black',
                          padx=75, pady=0.2)
        if RANGE in savedSettings:
            chk.select()
        chk.pack()


def resetCustomSettings():
    if exists(customConfig_path):
        remove(customConfig_path)
    top.destroy()


def apply():
    global customIpRanges
    customIpRanges.clear()
    for IP_NAME in ip_ranges_files:
        if ip_range_checkboxes[IP_NAME].get() == 1:
            print("Server: ", IP_NAME, " Checkbox state: ", str(ip_range_checkboxes[IP_NAME].get()))
            if IP_NAME not in customIpRanges:
                customIpRanges.append(IP_NAME)
    with open(localappdata_path + '\\customConfig.txt', 'w') as fp:
        for item in customIpRanges:
            # write each item on a new line
            fp.write("%s\n" % item)
    print(customIpRanges)
    userConfigSorter()
    top.destroy()


def openListFolder():
    if exists(localappdata_path):
        startfile(localappdata_path)


def blockALL():  # This function is for testing reasons only DO NOT USE.
    unblockALL()
    blockingLabel.config(text='ALL BLOCKED', fg='#ef2626')


def blockMEServer():  # It removes any rules added by blockserver function
    unblockALL()
    blockingLabel.config(text='WORKING ON IT', fg='#26ef4c')
    commands = 'advfirewall firewall add rule name="@ME_OW_SERVER_BLOCKER" Dir=Out Action=Allow'
    shell.ShellExecuteEx(lpVerb='runas', lpFile='netsh.exe', lpParameters=commands)
    threading.Thread(target=ruleMakerBlock, args=(blockingConfigDic['Ip_ranges_ME'], 445,),
                     daemon=True, kwargs={'block_exception': False}).start()  # Follow main thread


def PlayAustralia_server():
    unblockALL()
    blockingLabel.config(text='WORKING ON IT', fg='#26ef4c')
    commands = 'advfirewall firewall add rule name="@Australia_OW_SERVER_BLOCKER" Dir=Out Action=Allow'
    shell.ShellExecuteEx(lpVerb='runas', lpFile='netsh.exe', lpParameters=commands)
    threading.Thread(target=ruleMakerBlock, args=(blockingConfigDic['Ip_ranges_Australia'], 445,),
                     daemon=True).start()  # Follow main thread


def playNAEast_server():
    unblockALL()
    blockingLabel.config(text='WORKING ON IT', fg='#26ef4c')
    commands = 'advfirewall firewall add rule name="@NAEAST_OW_SERVER_BLOCKER" Dir=Out Action=Allow'
    shell.ShellExecuteEx(lpVerb='runas', lpFile='netsh.exe', lpParameters=commands)
    threading.Thread(target=ruleMakerBlock, args=(blockingConfigDic['Ip_ranges_NA_East'], 445,),
                     daemon=True).start()  # Follow main thread


def playNAWest_server():
    unblockALL()
    blockingLabel.config(text='WORKING ON IT', fg='#26ef4c')
    commands = 'advfirewall firewall add rule name="@NAWEST_OW_SERVER_BLOCKER" Dir=Out Action=Allow'
    shell.ShellExecuteEx(lpVerb='runas', lpFile='netsh.exe', lpParameters=commands)
    threading.Thread(target=ruleMakerBlock, args=(blockingConfigDic['Ip_ranges_NA_West'], 445,),
                     daemon=True).start()  # Follow main thread


def playEU_server():
    unblockALL()
    blockingLabel.config(text='WORKING ON IT', fg='#26ef4c')
    commands = 'advfirewall firewall add rule name="@EU_OW_SERVER_BLOCKER" Dir=Out Action=Allow'
    shell.ShellExecuteEx(lpVerb='runas', lpFile='netsh.exe', lpParameters=commands)
    threading.Thread(target=ruleMakerBlock, args=(blockingConfigDic['Ip_ranges_EU'], 445,),
                     daemon=True).start()  # Follow main thread


def programmableConfig():
    if len(customConfig) >= 1:
        unblockALL()
        commands = 'advfirewall firewall add rule name="@CUSTOM_BLOCK" Dir=Out Action=Allow'
        shell.ShellExecuteEx(lpVerb='runas', lpFile='netsh.exe', lpParameters=commands)
        blockIpRange(customConfig, rule_name='@Overwatch Block')
        blockingLabel.config(text="CUSTOM BLOCK", bg='#282828', fg='#26ef4c', font=futrabook_font)


def unblockALL():
    blockingLabel.config(text='ALL UNBLOCKED (DEFAULT SETTINGS)', fg='#ddee4a')
    list_rule_names = ["@NAEAST_OW_SERVER_BLOCKER", "@EU_OW_SERVER_BLOCKER", "@ME_OW_SERVER_BLOCKER",
                       "@NAWEST1_OW_SERVER_BLOCKER", "@AU_OW_SERVER_BLOCKER", "@NAWEST2_OW_SERVER_BLOCKER",
                       "@Overwatch Block", "@NAWEST_OW_SERVER_BLOCKER", "@Australia_OW_SERVER_BLOCKER", "@CUSTOM_BLOCK"]
    ruleDelete(list_rule_names)


def donationPage():
    webbrowser.open("https://paypal.me/vantverx?country.x=SA&locale.x=en_US")


# Menus
menu = Menu(app)
app.config(menu=menu)
options_menu = Menu(menu)
menu.add_cascade(label="Options", menu=options_menu)
options_menu.add_command(label="Open config folder", command=openListFolder)
options_menu.add_command(label="Check for updates", command=check_ip_update)
options_menu.add_command(label="Exit", command=app.quit)

# Labels
adminLabel = Label(app, text='', bg='#282828', fg='#ef2626', font=futrabook_font)
adminLabel.grid(row=0, column=0)
adminLabel.place(x=250, y=60, anchor="center")

blockingLabel = Label(app, text='', bg='#282828', fg='#ddee4a', font=futrabook_font)
blockingLabel.grid(row=0, column=0)
blockingLabel.place(x=250, y=440, anchor="center")

internetLabel = Label(app, text='', bg='#282828', fg='#26ef4c', font=futrabook_font)
internetLabel.grid(row=0, column=0)
internetLabel.place(x=250, y=420, anchor="center")

progressBar = ttk.Progressbar(app, orient=HORIZONTAL, length=130, mode='determinate')

# Buttons
y_axis = range(70, 450, 48)

PlayMEButton = Button(app, image=button_img_ME, font=futrabook_font, command=blockMEServer,
                      bg='#282828', fg='#282828', borderwidth=0, activebackground='#282828')
PlayMEButton.place(x=135, y=y_axis[0], height=40, width=230)

PlayEUButton = Button(app, image=button_img_EU, font=futrabook_font, command=playEU_server,
                      bg='#282828', fg='#282828', borderwidth=0, activebackground='#282828')
PlayEUButton.place(x=135, y=y_axis[1], height=40, width=230)

PlayNAWESTButton = Button(app, image=button_img_NA_WEST, font=futrabook_font, command=playNAWest_server,
                          bg='#282828', fg='#282828', borderwidth=0, activebackground='#282828')
PlayNAWESTButton.place(x=135, y=y_axis[2], height=40, width=230)

PlayNAEASTButton = Button(app, image=button_img_ME_EAST, font=futrabook_font, command=playNAEast_server,
                          bg='#282828', fg='#282828', borderwidth=0, activebackground='#282828')
PlayNAEASTButton.place(x=135, y=y_axis[3], height=40, width=230)

PlayAustraliaButton = Button(app, image=button_img_Australia, font=futrabook_font, command=PlayAustralia_server,
                             bg='#282828', fg='#282828', borderwidth=0, activebackground='#282828')
PlayAustraliaButton.place(x=135, y=y_axis[4], height=40, width=230)

ProgrammableButton = Button(app, image=button_img_programmable_button, font=futrabook_font, command=programmableConfig,
                            bg='#282828', fg='#282828', borderwidth=0, activebackground='#282828')
ProgrammableButton.place(x=135, y=y_axis[5], height=40, width=230)

ClearBlocksButton = Button(app, image=button_img_Default, font=futrabook_font, command=unblockALL,
                           bg='#282828', fg='#282828', borderwidth=0, activebackground='#282828')
ClearBlocksButton.place(x=135, y=y_axis[6], height=40, width=230)

DonationButton = Button(app, image=button_img_donation, font=futrabook_font, command=donationPage,
                        bg='#282828', fg='#282828', borderwidth=0, activebackground='#282828')
DonationButton.place(x=420, y=480, height=73, width=68)

CustomSettingsSButton = Button(app, image=button_img_CUSTOM_SETTINGS, font=futrabook_font, command=customSettingsWindow,
                               bg='#282828', fg='#282828', borderwidth=0, activebackground='#282828')
CustomSettingsSButton.place(x=370, y=318, height=25, width=25)

# Check box
tunnelCheckBox_state = IntVar()
tunnelCheckBox = Checkbutton(app, text="Only affect Overwatch ", font=futrabook_font, activebackground='#282828',
                             bg='#282828', fg='#26ef4c', borderwidth=0, variable=tunnelCheckBox_state, command=tunnel)
tunnelCheckBox.place(x=175, y=460)

# Start Program
iconMaker()

check_admin()

ipSorter_thread = threading.Thread(target=ipSorter, daemon=True).start()  # Follow main thread

checkIfActive_thread = threading.Thread(target=checkIfActive, daemon=True).start()  # Follow main thread
check_ip_update.thread = threading.Thread(target=check_ip_update, daemon=True).start()  # Follow main thread

if checkOptions():
    tunnelCheckBox.select()

app.mainloop()
