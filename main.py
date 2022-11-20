from tkinter import *
from tkinter.font import Font
from tkinter import ttk, filedialog
import win32api
import win32com.shell.shell as shell
from win32com.client import Dispatch as DispatchCOMObject
from pythoncom import CoInitialize
from itertools import groupby
from subprocess import run
from PIL import ImageTk, Image
from io import BytesIO
import pic2str
import base64
import ctypes
from os.path import exists, isdir, isfile, join
from os import getenv, path, mkdir, listdir, linesep, startfile, remove, chmod
import webbrowser
import socket
import threading
import requests
import logging
import datetime
from urllib3 import Retry
from requests.adapters import HTTPAdapter

# Information
__version__ = '5.1.0'
_AppName_ = 'MINA Overwatch 2 Server Selector'
__author__ = 'Yousef Aljohani'
__copyright__ = 'Copyright (C) 2022, Yousef Aljohani'
__credits__ = ['Yousef Aljohani(foryVERX)', 'chhaugen(Carl)']
__maintainer__ = 'Yousef Aljohani'
__email__ = 'verrrx@gmail.com'

# Create window object
app = Tk()

# Set Properties
app.title(_AppName_)
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
byte_INSTALL_UPDATE = base64.b64decode(pic2str.INSTALL_UPDATE)

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
image_data_INSTALL_UPDATE = BytesIO(byte_INSTALL_UPDATE)

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
button_img_INSTALL_UPDATE = ImageTk.PhotoImage(Image.open(image_data_INSTALL_UPDATE))

# Add font
futrabook_font = Font(family="Futura PT Demi", size=10)

# Global variables
localappdata_path = getenv('APPDATA') + '\\OverwatchServerBlocker'
temp_path = getenv('APPDATA') + '\\Local\\Temp'
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
appVersion_url = 'https://raw.githubusercontent.com/foryVERX/Overwatch-Server-Selector/main/__version__/__latestversion__.txt'

DEFAULT_BLOCK_NAME = "_Overwatch Block"
DEFAULT_GROUPING_NAME = "_MINA Overwatch 2-Server-Selector"

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

# FirewallAPI COM objects and constants

FIREWALL_ACTION_BLOCK = 0
FIREWALL_ACTION_ALLOW = 1

FIREWALL_DIRECTION_IN = 1
FIREWALL_DIRECTION_OUT = 2


# https://stackoverflow.com/a/27966218
# DO NOT PASS TO OTHER THREADS
def dispatchFirewall():
    CoInitialize()
    return DispatchCOMObject("HNetCfg.FwPolicy2")


# DO NOT PASS TO OTHER THREADS
def dispatchFirewallRule():
    CoInitialize()
    return DispatchCOMObject("HNetCfg.FWRule")


def addNewRuleToFirewall(name, direction, action, remoteAddresses=None, applicationName=None, grouping=None,
                         enabled=True):
    firewall = dispatchFirewall()
    rule = dispatchFirewallRule()
    rule.Name = name
    if remoteAddresses is not None:
        rule.RemoteAddresses = remoteAddresses
    rule.Direction = direction  # Outgoing
    rule.Action = action  # Block
    if applicationName is not None:
        rule.ApplicationName = applicationName
    if grouping is not None:
        rule.Grouping = grouping
    rule.enabled = enabled
    firewall.Rules.Add(rule)


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
        update_text = "UPDATING IP LIST..."
        internetLabel.config(text=update_text, fg='#ddee4a')
        msg_fail = "CONNECTION FAILED... Trying to update ip list"
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
        logging.info("IP LIST UPDATED")
        updating_state = False
        checkForUpdate_initialization = False
        loadIpRanges()
    else:
        update_text = "Please check your internet connection to download servers ip"
        internetLabel.config(text=update_text, fg='#ddee4a')


def check_ip_update():  # A function called at the start of the program to check for update
    global isUpdated, updating_state, checkForUpdate_initialization
    if updating_state:
        return
    updating_state = True
    is_connect()
    logging.info("Checking for ip list updates")
    if internetConnection:
        if isdir(localappdata_path) and path.exists(ip_version_path):
            logging.info(localappdata_path + ' FOUND')
            logging.info(ip_version_path + ' FOUND')
            with open(ip_version_path, "r") as reader:  # Read Ip_version.txt from GitHub and analyze
                logging.info('Reading Ip_version.txt')
                for line in reader.readlines():
                    if len(line) > 1:
                        msg_fail = "Ip list update check failed"
                        with requests.Session() as s:
                            adapter = HTTPAdapter(max_retries=Retry(total=4, backoff_factor=1, allowed_methods=None,
                                                                    status_forcelist=[429, 500, 502, 503, 504]))
                            s.mount("http://", adapter)
                            s.mount("https://", adapter)
                            ip_version_request = request_raw_file(ip_version_url, msg_fail, s)
                        ip_version_request = linesep.join([s for s in ip_version_request.splitlines() if s])
                        line = linesep.join([s for s in line.splitlines() if s])
                        logging.debug("LATEST IP LIST VERSION : " + str(ip_version_request))
                        if ip_version_request == line:
                            update_text = "IP LIST IS UPDATED"
                            app.after(250, internetLabel.config(text=update_text, fg='#26ef4c'))
                        else:
                            update_text = "IP LIST IS NOT UPDATED"
                            app.after(250, internetLabel.config(text=update_text, fg='#ef2626'))
                            threading.Thread(target=updateIp, daemon=True).start()
        else:  # Make directory and call updateIp
            update_text = "FIRST TIME RUNNING.. DOWNLOADING IP LIST"
            app.after(250, internetLabel.config(text=update_text, fg='#ddee4a'))
            if not exists(localappdata_path):
                mkdir(localappdata_path)  # Make directory
            threading.Thread(target=updateIp, daemon=True).start()
    else:
        logging.debug('No internet connection')
        if not exists(ip_version_path):
            controlButtons('disabled')
            update_text = "CONNECTION FAILED... Trying to update ip list"
            app.after(250, internetLabel.config(text=update_text, fg='#ef2626'))
            app.after(1000, check_ip_update)
        else:
            update_text = "NO INTERNET MIGHT BE NOT LATEST IP LIST VERSION"
            app.after(250, internetLabel.config(text=update_text, fg='#ddee4a'))
    updating_state = False
    # if thread_type == 'mainThread':  # If the function is called from main thread call it again after 5 mints
    # app.after(5000 * 60, checkUpdate)


def check_app_update():
    """
    Function request the latest app version and compare it with installed one
    Downloads new version if available at path/temp
    Pop a "INSTALL UPDATE" button if new app update is detected
    :return: None
    """
    global latestVersion_path
    with requests.Session() as s:
        adapter = HTTPAdapter(max_retries=Retry(total=4, backoff_factor=1, allowed_methods=None,
                                                status_forcelist=[429, 500, 502, 503, 504]))
        s.mount("http://", adapter)
        s.mount("https://", adapter)
        latestVersion = request_raw_file(appVersion_url, "Getting latest version failed", s).splitlines()
        latestVersion = [i for i in latestVersion if i]  # To make sure no empty lines
        latestVer = latestVersion[0].strip()
        urlDownload = latestVersion[1].strip()
        logging.info('Reading latest app version from GitHub.txt')
        logging.info("Latest " + str(latestVer))
        logging.info("Latest version " + str(urlDownload))
        latestVer = latestVer.strip("version=")
        urlDownload = urlDownload.strip("url=")
        if not __version__ == latestVer:
            logging.info("APP UPDATE IS AVAILABLE: v" + str(latestVer))
            latestVersion_path = temp_path.replace('\\Roaming', '') + '\\' + f'{_AppName_} {latestVer}.exe'
            if not exists(latestVersion_path):
                downloadedBytes = s.get(urlDownload)
                open(latestVersion_path, "wb").write(downloadedBytes.content)
                logging.info("Downloaded installer at " + latestVersion_path)
            InstallUpdateButton.place(x=135, y=500, height=40, width=230)
        else:
            logging.info("APP IS UPDATED")


def installUpdate():
    """
    Function is initiated when install update button is pressed
    It executes the setup downloaded from check_app_update
    :return: None
    """
    if exists(latestVersion_path):
        win32api.ShellExecute(0, 'open', latestVersion_path, None, None, 10)
        app.destroy()


def request_raw_file(url, msg_fail, s):
    """
    :arg: url, msg_fail, s
    url: is the url at which raw text exists
    msg_fail: the message given to log when connection fails
    s: the session initiated

    :return:
    result: decoded bytes from source
    """
    try:
        # user-agent is just to trick the website that you are using a browser
        headers = {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:66.0) Gecko/20100101 Firefox/66.0"
        }
        result = s.get(url, headers=headers).content.decode('utf-8')
    except requests.exceptions.RequestException as e:  # This is the correct syntax
        logging.debug('URL REQUEST FAIL RETRYING: ' + url)
        logging.debug(str(e))
        raise SystemExit(e)
    return result


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


def readIpRangesByFilename(fileName):
    with open(localappdata_path + "\\" + fileName, "r") as reader:
        ipLines = reader.readlines()
        ipLinesStriped = map(lambda line: line.strip('\n'), ipLines)
        ipLinesNonZeroLen = filter(lambda line: len(line) > 0, ipLinesStriped)
        return list(ipLinesNonZeroLen)


def loadIpRanges():  # Store ip ranges from Ip_ranges_....txt into Ip_ranges dictionary
    global Ip_ranges_dic, sorter_initialization
    loadUserConfig()
    loadBlockingConfig()
    if exists(localappdata_path) and exists(ip_version_path):  # If those paths exists it means user updated
        logging.info(f'{localappdata_path} FOUND')
        logging.info(f'{ip_version_path} FOUND')
        controlButtons('disabled')
        appDataFilenames = listdir(localappdata_path)
        ipRangesFilenames = filter(lambda server: server.startswith("Ip_ranges"), appDataFilenames)
        Ip_ranges_dic = {path.splitext(ipRangesFilename)[0]: readIpRangesByFilename(ipRangesFilename) for
                         ipRangesFilename in ipRangesFilenames}

        update_text = "UPDATED"
        app.after(250, internetLabel.config(text=update_text, fg='#26ef4c'))
        sorter_initialization = True
        logging.info("IP LIST LOADED")
    else:  # User running first time
        check_ip_update()
    controlButtons('normal')


def loadUserConfig():
    global customConfig
    customConfig.clear()
    if not exists(customConfig_path):
        return
    logging.debug(f'{customConfig_path} FOUND')
    with open(customConfig_path, "r") as customConfigFile:
        configLines = customConfigFile.readlines()
        configLinesStriped = map(lambda line: line.strip('\n'), configLines)
        configLinesNonZeroLen = filter(lambda line: len(line) > 0, configLinesStriped)
        configLinesFileExists = filter(lambda line: exists(localappdata_path + "\\" + line), configLinesNonZeroLen)
        configIpRangesPerServer = map(lambda line: readIpRangesByFilename(line), configLinesFileExists)

        # [ item for list in listoflists for item in list] https://stackoverflow.com/q/1077015
        customConfig = [ipRange for IpRangesChunk in configIpRangesPerServer for ipRange in IpRangesChunk]
        print()


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


def blockServers(server_exception, block_exception=True, rule_name=DEFAULT_BLOCK_NAME,
                 rule_grouping=DEFAULT_GROUPING_NAME):
    # Used to block IP range
    # server_exception is the only server to not block can be a list or string
    # If block_exception set to false then the server_exception is blocked ONLY
    controlButtons('disabled')
    serversToBlock = Ip_ranges_dic

    if not block_exception:
        serverBlockFilter = lambda serverName: serverName in str(server_exception)
    else:
        serverBlockFilter = lambda serverName: serverName not in str(server_exception)

    serversToBlock = {serverName: ipRanges for serverName, ipRanges in serversToBlock.items() if
                      serverBlockFilter(serverName)}
    # [ item for key, list in dictionaryoflists for item in list] https://stackoverflow.com/q/1077015
    ipRangesToBlock = [ipRange for serverName, ipRanges in serversToBlock.items() for ipRange in ipRanges]

    blockIpRanges(ipRangesToBlock, rule_name, rule_grouping)
    serversBlockedString = ', '.join(serversToBlock.keys())
    logging.info(f'Blocked {serversBlockedString}')
    checkIfActive()
    controlButtons('normal')


def blockIpRanges(ip_list, rule_name, rule_grouping):
    # A Windows Firewall Rule supports blocking unique 10_000 IP range entries. (Tested on Windows 8.1 and Windows 10)

    applicationName = overwatch_path if tunnel_option else None

    if (len(ip_list) <= 10_000):
        ipRangesString = ','.join(ip_list)
        addNewRuleToFirewall(rule_name, FIREWALL_DIRECTION_IN, FIREWALL_ACTION_BLOCK, ipRangesString, applicationName,
                             rule_grouping)
        addNewRuleToFirewall(rule_name, FIREWALL_DIRECTION_OUT, FIREWALL_ACTION_BLOCK, ipRangesString, applicationName,
                             rule_grouping)
        logging.info(f'Made rules "{rule_name} for IN/OUT"')
    else:
        indexedIpRangeList = list(enumerate(ip_list))
        ipRangesGrouped = groupby(indexedIpRangeList, key=lambda item: item[0] // 10_000)  # Make 10_000 chunks
        ipRangesGroupedDict = {k: [x[1] for x in v] for k, v in ipRangesGrouped}
        ipRangesStringChunks = {key: ','.join(data) for (key, data) in ipRangesGroupedDict.items()}

        for chunkNum, ipStringChunk in ipRangesStringChunks.items():
            if chunkNum > 0:
                rule_name = f'{rule_name} {chunkNum}'

            addNewRuleToFirewall(rule_name, FIREWALL_DIRECTION_IN, FIREWALL_ACTION_BLOCK, ipStringChunk,
                                 applicationName, rule_grouping)
            addNewRuleToFirewall(rule_name, FIREWALL_DIRECTION_OUT, FIREWALL_ACTION_BLOCK, ipStringChunk,
                                 applicationName, rule_grouping)
            logging.info(f'Made rules "{rule_name} for IN/OUT"')


def deleteRule(rule_name):  # Delete rule by exact name, name must be a string '' or list of strings
    logging.info("DELETING RULES: " + str(rule_name))
    # Ensure its a list
    if type(rule_name) != list:
        rule_name = [rule_name]
    for name in rule_name:
        firewall = dispatchFirewall()
        firewall.Rules.Remove(name)


def deleteRuleGrouping(rule_grouping):
    logging.info("DELETING RULE GROUPINGS: " + str(rule_grouping))
    # Ensure its a list
    if type(rule_grouping) != list:
        rule_grouping = [rule_grouping]

    firewall = dispatchFirewall()

    ruleNamesToDelete = [rule.Name for rule in firewall.Rules if rule.Grouping in rule_grouping]
    deleteRule(ruleNamesToDelete)


def checkForAndDeleteLegacyRules():
    legacyRuleNames = ["@NAEAST_OW_SERVER_BLOCKER", "@EU_OW_SERVER_BLOCKER", "@ME_OW_SERVER_BLOCKER",
                       "@NAWEST1_OW_SERVER_BLOCKER", "@AU_OW_SERVER_BLOCKER", "@NAWEST2_OW_SERVER_BLOCKER",
                       "@Overwatch Block", "@NAWEST_OW_SERVER_BLOCKER", "@Australia_OW_SERVER_BLOCKER", "@CUSTOM_BLOCK"]
    firewall = dispatchFirewall()
    ruleNamesToDelete = [rule.Name for rule in firewall.Rules if rule.Name in legacyRuleNames]
    if len(ruleNamesToDelete) > 0:
        logging.info(f"DELETING RULES: {str(ruleNamesToDelete)}")
        for ruleName in ruleNamesToDelete:
            commands = f'advfirewall firewall delete rule name = "{ruleName}"'
            shell.ShellExecuteEx(lpVerb='runas', lpFile='netsh.exe', lpParameters=commands)


def checkIfActive():  # To check if server is blocked or not
    servers_active_rule_list = ['_ME_OW_SERVER_BLOCKER', '_NAEAST_OW_SERVER_BLOCKER', '_NAWEST_OW_SERVER_BLOCKER',
                                '_EU_OW_SERVER_BLOCKER', '_AU_OW_SERVER_BLOCKER', '_Australia_OW_SERVER_BLOCKER',
                                "_CUSTOM_BLOCK"]
    firewall = dispatchFirewall()
    rules = [x.Name for x in firewall.Rules]
    logging.info("CHECKING IF PREVIOUS RULE IS ACTIVE")
    for rule_name in servers_active_rule_list:
        if rule_name in rules:
            filtered = rule_name.split('_')[1]
            if filtered == 'ME':
                blockingLabel.config(text='ME BLOCKED', bg='#282828', fg='#ef2626', font=futrabook_font)
                return
            if filtered == 'CUSTOM':
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


def loadBlockingConfig():
    global blockingConfigDic
    if exists(localappdata_path + '\\BlockingConfig.txt'):
        with open(localappdata_path + '\\BlockingConfig.txt', "r") as reader:
            blockingConfigLines = reader.readlines()
            filledLines = filter(lambda line: '@' in line, blockingConfigLines)
            linesRemovePrefix = [line.replace('ipRangeName::', '') for line in filledLines]
            linesStriped = [line.strip('\n') for line in linesRemovePrefix]
            linesEntriesSplit = [line.split('@') for line in linesStriped]
            blockingConfigDic = {entries[0]: entries[1:] for entries in linesEntriesSplit}
        for key, value in blockingConfigDic.items():
            logging.info(f'Blocking Config: Play on: {key} wants to exclude {str(value)} From Blocking')


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
    customConfig.clear()
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
    loadUserConfig()
    top.destroy()


def openListFolder():
    if exists(localappdata_path):
        startfile(localappdata_path)


def blockALL():  # This function is for testing reasons only DO NOT USE.
    unblockALL()
    blockingLabel.config(text='ALL BLOCKED', fg='#ef2626')


def blockMEServer():  # It removes any rules added by block server function
    unblockALL()
    blockingLabel.config(text='WORKING ON IT', fg='#26ef4c')
    addNewRuleToFirewall("_ME_OW_SERVER_BLOCKER", FIREWALL_DIRECTION_OUT, FIREWALL_ACTION_BLOCK, enabled=False,
                         grouping=DEFAULT_GROUPING_NAME)
    threading.Thread(target=blockServers, args=(blockingConfigDic['Ip_ranges_ME'],),
                     daemon=True, kwargs={'block_exception': False}).start()  # Follow main thread


def PlayAustralia_server():
    unblockALL()
    blockingLabel.config(text='WORKING ON IT', fg='#26ef4c')
    addNewRuleToFirewall("_Australia_OW_SERVER_BLOCKER", FIREWALL_DIRECTION_OUT, FIREWALL_ACTION_BLOCK, enabled=False,
                         grouping=DEFAULT_GROUPING_NAME)
    threading.Thread(target=blockServers, args=(blockingConfigDic['Ip_ranges_Australia'],),
                     daemon=True).start()  # Follow main thread


def playNAEast_server():
    unblockALL()
    blockingLabel.config(text='WORKING ON IT', fg='#26ef4c')
    addNewRuleToFirewall("_NAEAST_OW_SERVER_BLOCKER", FIREWALL_DIRECTION_OUT, FIREWALL_ACTION_BLOCK, enabled=False,
                         grouping=DEFAULT_GROUPING_NAME)
    threading.Thread(target=blockServers, args=(blockingConfigDic['Ip_ranges_NA_East'],),
                     daemon=True).start()  # Follow main thread


def playNAWest_server():
    unblockALL()
    blockingLabel.config(text='WORKING ON IT', fg='#26ef4c')
    addNewRuleToFirewall("_NAWEST_OW_SERVER_BLOCKER", FIREWALL_DIRECTION_OUT, FIREWALL_ACTION_BLOCK, enabled=False,
                         grouping=DEFAULT_GROUPING_NAME)
    threading.Thread(target=blockServers, args=(blockingConfigDic['Ip_ranges_NA_West'],),
                     daemon=True).start()  # Follow main thread


def playEU_server():
    unblockALL()
    blockingLabel.config(text='WORKING ON IT', fg='#26ef4c')
    addNewRuleToFirewall("_EU_OW_SERVER_BLOCKER", FIREWALL_DIRECTION_OUT, FIREWALL_ACTION_BLOCK, enabled=False,
                         grouping=DEFAULT_GROUPING_NAME)
    threading.Thread(target=blockServers, args=(blockingConfigDic['Ip_ranges_EU'],),
                     daemon=True).start()  # Follow main thread


def programmableConfig():
    if len(customConfig) >= 1:
        unblockALL()
        addNewRuleToFirewall("_CUSTOM_BLOCK", FIREWALL_DIRECTION_OUT, FIREWALL_ACTION_BLOCK, enabled=False,
                             grouping=DEFAULT_GROUPING_NAME)
        blockIpRanges(customConfig, rule_name=DEFAULT_BLOCK_NAME, rule_grouping=DEFAULT_GROUPING_NAME)
        blockingLabel.config(text="CUSTOM BLOCK", bg='#282828', fg='#26ef4c', font=futrabook_font)


def unblockALL():
    blockingLabel.config(text='ALL UNBLOCKED (DEFAULT SETTINGS)', fg='#ddee4a')
    checkForAndDeleteLegacyRules()
    deleteRuleGrouping(DEFAULT_GROUPING_NAME)


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

versionLabel = Label(app, text=f'V {__version__}', bg='#282828', fg='#26ef4c', font=futrabook_font)
versionLabel.grid(row=0, column=0)
versionLabel.place(x=460, y=590, anchor="center")

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

CustomSettingsButton = Button(app, image=button_img_CUSTOM_SETTINGS, font=futrabook_font, command=customSettingsWindow,
                              bg='#282828', fg='#282828', borderwidth=0, activebackground='#282828')
CustomSettingsButton.place(x=370, y=318, height=25, width=25)

InstallUpdateButton = Button(app, image=button_img_INSTALL_UPDATE, font=futrabook_font, command=installUpdate,
                             bg='#282828', fg='#282828', borderwidth=0, activebackground='#282828')

# Check box
tunnelCheckBox_state = IntVar()
tunnelCheckBox = Checkbutton(app, text="Only affect Overwatch ", font=futrabook_font, activebackground='#282828',
                             bg='#282828', fg='#26ef4c', borderwidth=0, variable=tunnelCheckBox_state, command=tunnel)
tunnelCheckBox.place(x=175, y=460)

# Start Program
iconMaker()
check_admin()
checkForAndDeleteLegacyRules()
if checkOptions(): tunnelCheckBox.select()

check_app_update_thread = threading.Thread(target=check_app_update, daemon=True).start()  # Follow main thread
loadIpRanges_thread = threading.Thread(target=loadIpRanges, daemon=True).start()  # Follow main thread
checkIfActive_thread = threading.Thread(target=checkIfActive, daemon=True).start()  # Follow main thread
check_ip_update.thread = threading.Thread(target=check_ip_update, daemon=True).start()  # Follow main thread

app.mainloop()
