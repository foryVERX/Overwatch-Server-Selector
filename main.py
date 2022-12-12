from tkinter import *
from tkinter.font import Font
from tkinter import ttk, filedialog
import win32api
import win32com.shell.shell as shell
from win32com.client import Dispatch as DispatchCOMObject
from pythoncom import CoInitialize
from itertools import groupby
from PIL import ImageTk, Image
from io import BytesIO
import pic2str
import base64
import ctypes
from os.path import exists, isdir, isfile, join, basename
from os import getenv, path, mkdir, listdir, linesep, startfile, remove, chmod
import webbrowser
import socket
import threading
import requests
import logging
import datetime
from urllib3 import Retry
from requests.adapters import HTTPAdapter
from pythonping import ping

# Information
__version__ = '5.3.0'
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
byte_TEST_PING = base64.b64decode(pic2str.TEST_PING)
byte_Discord_Button = base64.b64decode(pic2str.Discord_Button)

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
image_data_TEST_PING = BytesIO(byte_TEST_PING)
image_data_Discord_Button = BytesIO(byte_Discord_Button)

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
button_img_TEST_PING = ImageTk.PhotoImage(Image.open(image_data_TEST_PING))
button_img_Discord_Button = ImageTk.PhotoImage(Image.open(image_data_Discord_Button))

# Add font
futrabook_font = Font(family="Futura PT Demi", size=10)

# Global variables
localappdata_path = getenv('APPDATA') + '\\OverwatchServerBlocker'
temp_path = getenv('APPDATA') + '\\Local\\Temp'
ip_version_path = localappdata_path + '\\IP_version.txt'
overwatch_path = 'C:\\Program Files (x86)\\Overwatch\\_retail_\\Overwatch.exe'
customConfig_path = localappdata_path + '\\customConfig.txt'
ip_version_url = \
    'https://raw.githubusercontent.com/foryVERX/Overwatch-Server-Selector/main/ip_lists/IP_version.txt'
appVersion_url = \
    'https://raw.githubusercontent.com/foryVERX/Overwatch-Server-Selector/main/__version__/__latestversion__.txt'
urlContainer_url = \
    "https://raw.githubusercontent.com/foryVERX/Overwatch-Server-Selector/main/ip_lists/urlsContainer.txt"

DEFAULT_BLOCK_NAME = "_Overwatch Block"
DEFAULT_GROUPING_NAME = "_MINA Overwatch 2-Server-Selector"

updating_state = False
appUpdate_required = False
sorter_initialization = False
checkForUpdate_initialization = False
tunnel_option = False
isUpdated = ''
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

logging.info(f"Running version v{__version__}")

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


def internetConnection():
    try:
        socket.create_connection(("www.google.com", 80))
        return True
    except OSError:
        return False


def updateIpList():
    global updating_state, sorter_initialization
    controlButtons('disabled')
    pingButton.place_forget()
    InstallUpdateButton.place_forget()
    if not internetConnection():
        update_text = "Please check your internet connection to download servers ip"
        internetLabel.config(text=update_text, fg='#ddee4a')
        return
    updating_state = True
    update_text = "UPDATING IP LIST..."
    internetLabel.config(text=update_text, fg='#ddee4a')
    start = datetime.datetime.now()
    download_Ip_List()
    finish = datetime.datetime.now() - start
    logging.info('UPDATING TOOK: ' + str(finish))
    logging.info("IP LIST UPDATED")
    updating_state = False
    if not appUpdate_required:
        pingButton.place(x=135, y=500, height=40, width=230)
    else:
        InstallUpdateButton.place(x=135, y=500, height=40, width=230)
    loadIpRanges()
    update_text = "IP LIST IS UPTODATE"
    app.after(250, internetLabel.config(text=update_text, fg='#26ef4c'))


def check_ip_update():  # A function called at the start of the program to check for update
    global isUpdated, updating_state
    if updating_state:
        return
    logging.info("CHECKING LATEST IP LIST VERSION FROM GITHUB")
    if not internetConnection():
        logging.debug('No internet connection')
        if not exists(ip_version_path):
            controlButtons('disabled')
            update_text = "CONNECTION FAILED... Trying to update ip list"
            app.after(250, internetLabel.config(text=update_text, fg='#ef2626'))
            app.after(1000, check_ip_update)
            return
        update_text = "NO INTERNET MIGHT BE NOT LATEST IP LIST VERSION"
        app.after(250, internetLabel.config(text=update_text, fg='#ddee4a'))

    if not isdir(localappdata_path):
        logging.info(localappdata_path + ' NOT FOUND')
        update_text = "FIRST TIME RUNNING.. DOWNLOADING IP LIST"
        app.after(250, internetLabel.config(text=update_text, fg='#ddee4a'))
        makeApp_directory()
        threading.Thread(target=updateIpList, daemon=True).start()
        return

    if not exists(ip_version_path):
        logging.info(ip_version_path + ' NOT FOUND')
        threading.Thread(target=updateIpList, daemon=True).start()
        return

    logging.info(localappdata_path + ' FOUND')
    logging.info(ip_version_path + ' FOUND')
    logging.info('READING CURRENT IP LIST VERSION')

    for line in readLine_filtered(ip_version_path):
        msg_fail = "Ip list update check failed"
        session = createHTTP_session()
        LatestIpListVersion = request_raw_file(ip_version_url, msg_fail, session)
        LatestIpListVersion = filterStrings(LatestIpListVersion, removeSpace=True, removeNewlines=True)
        currentIpListVersion = filterStrings(line, removeSpace=True, removeNewlines=True)
        logging.debug("CURRENT IP LIST VERSION : " + str(currentIpListVersion))
        logging.debug("LATEST IP LIST VERSION : " + str(LatestIpListVersion))

        if not LatestIpListVersion == currentIpListVersion:
            update_text = "IP LIST IS NOT UPDATED"
            app.after(250, internetLabel.config(text=update_text, fg='#ef2626'))
            threading.Thread(target=updateIpList, daemon=True).start()
            return
        update_text = "IP LIST IS UPTODATE"
        app.after(250, internetLabel.config(text=update_text, fg='#26ef4c'))
    # if thread_type == 'mainThread':  # If the function is called from main thread call it again after 5 mints
    # app.after(5000 * 60, checkUpdate)


def filterStrings(string, removeSpace=False, removeNewlines=False):
    if removeSpace:
        string = string.replace(" ", "")
    if removeNewlines:
        string = string.splitlines()
        string = str([i for i in string if i])  # remove empty lines from each element
    return string


def makeApp_directory():
    """
    Makes the app directory
    :return: None
    """
    if not exists(localappdata_path):
        mkdir(localappdata_path)  # Make directory


def readLine_filtered(txt_path):
    """
    :param txt_path: the path to the text file
    :return: Lines striped from \n and nonzero length
    """
    with open(txt_path, "r") as lines:  # Read Ip_version.txt from GitHub and analyze
        lines = lines.readlines()
        linesStriped = map(lambda line: line.strip('\n'), lines)
        linesStriped = map(lambda line: line.strip('\t'), linesStriped)
        linesStripedNonZeroLen = list(filter(lambda line: len(line) > 0, linesStriped))
        return linesStripedNonZeroLen


def createHTTP_session():
    print("Creating session")
    with requests.Session() as session:
        adapter = HTTPAdapter(max_retries=Retry(total=4, backoff_factor=1, allowed_methods=None,
                                                status_forcelist=[429, 500, 502, 503, 504]))
        session.mount("http://", adapter)
        session.mount("https://", adapter)
    return session


def check_app_update():
    """
    Function request the latest app version and compare it with installed one
    Downloads new version if available at path/temp
    Pop a "INSTALL UPDATE" button if new app update is detected
    :return: None
    """
    global latestVersion_path, appUpdate_required
    if not internetConnection():
        logging.info("COULD NOT CHECK FOR APP UPDATE DUE TO NO INTERNET CONNECTION")
        return
    session = createHTTP_session()
    logging.info('CHECKING LATEST APP VERSION FROM GITHUB')
    latestVersion = request_raw_file(appVersion_url, "Getting latest version failed", session).splitlines()
    # https://www.geeksforgeeks.org/python-remove-empty-strings-from-list-of-strings/
    latestVersion = [i for i in latestVersion if i]  # To make sure no empty lines
    latestVer = latestVersion[0].strip()
    urlDownload = latestVersion[1].strip()
    logging.info(f"LATEST APP VERSION {latestVer}")
    latestVer = latestVer.strip("version=")
    urlDownload = urlDownload.strip("url=")
    if not __version__ == latestVer:
        logging.info(f"NEW APP UPDATE IS AVAILABLE: v{latestVer}")
        latestVersion_path = temp_path.replace('\\Roaming', '') + '\\' + f'{_AppName_} {latestVer}.exe'
        if not exists(latestVersion_path):
            logging.info(f"DOWNLOADING APP UPDATE v{latestVer}")
            downloadedBytes = session.get(urlDownload)
            open(latestVersion_path, "wb").write(downloadedBytes.content)
            logging.info(f"APP UPDATE IS DOWNLOADED AT THE LOCATION: {latestVersion_path}")
        pingButton.place_forget()
        appUpdate_required = True
        if not updating_state:
            InstallUpdateButton.place(x=135, y=500, height=40, width=230)
        return
    if not updating_state and not appUpdate_required:
        pingButton.place(x=135, y=500, height=40, width=230)
    logging.info("CURRENT VERSION IS UPTODATE")


def installUpdate():
    """
    Function is initiated when install update button is pressed
    It executes the setup downloaded from check_app_update
    :return: None
    """
    if exists(latestVersion_path):
        win32api.ShellExecute(0, 'open', latestVersion_path, None, None, 10)
        app.destroy()


def parseUrls(textFilePath):
    """
    The function used to separate urls that exists in a text file line by line.
    example in textfile.txt:
    url1
    url2
    ... etc
    output is a dictionary where the key is the file name extracted from the url and value is the url

    :param textFilePath:
    :return: list of urls
    """
    with open(textFilePath, "r") as urlContainerTextFileReader:
        urlContainerTextFile = urlContainerTextFileReader.readlines()
        urlContainerLinesStriped = map(lambda line: line.strip('\n'), urlContainerTextFile)
        urlContainerNonZeroLen = list(filter(lambda line: len(line) > 0, urlContainerLinesStriped))
        urlContainerList = urlContainerNonZeroLen
        # Convert List to dictionary using file name as key
        urlContainerDictionary = {path.splitext(path.basename(line))[0]: line for line in urlContainerList}
    return urlContainerDictionary


def downloadUrls(filename, urlTarget):
    """
    Using url to download file

    :param filename: The name of output file
    :param urlTarget: The url to download the content
    :return:
    """
    session = createHTTP_session()
    downloadedBytes = session.get(urlTarget)
    open(filename, "wb").write(downloadedBytes.content)


def url2TextFile(dic, progressBar=False):
    """
    Convert dictionary to a text file where:
     *Key is file name
     *Value is the url that contains the content

    Output is text files where it's content is fetched from url

    :param

    dic: Dictionary where..
     *Key is file name
     *Value is the url that contains the content

    progressBar: if progressbar is true  will be shown at the gui
    :return: None
    """
    lengthProgressBar = 150
    if progressBar:
        progressBar = ttk.Progressbar(app, orient=HORIZONTAL, length=lengthProgressBar, mode='determinate')
        progressBar.place(x=135, y=510, height=20, width=230)
    session = createHTTP_session()
    for file_name, url in dic.items():
        logging.info(f"DOWNLOADING {url}")
        content = request_raw_file(url, "Requesting contents from url failed", session)
        createTextFile(file_name.replace("%20", ' '), content)
        if progressBar:
            progressBar['value'] += lengthProgressBar / len(dic)
    progressBar.place_forget()


def download_Ip_List():
    """
    Depends on these functions:
    *downloadUrls
    Collects the urls from a urlContainer that contains all urls in one text file.

    *parseUrls
     create a dictionary of urls as value with their tail name as a key, www.whatever.com/text.txt
     text.txt is the tail in this case.

    *url2TextFile
     When given the result of parseUrls function, it creates a text file
    :return:
    """
    logging.info("IP LIST UPDATING PROCESS IS INITIATED")
    urlContainerPath = temp_path.replace('\\Roaming', '') + '\\urlContainer.txt'
    logging.info("DOWNLOADING URL CONTAINER")
    downloadUrls(urlContainerPath, urlContainer_url)
    logging.info("URL CONTAINER DOWNLOADED SUCCESSFULLY")
    dictionary = parseUrls(urlContainerPath)
    logging.info("EXTRACTING URLS FROM URL CONTAINER")
    url2TextFile(dictionary, progressBar=True)


def request_raw_file(url, msg_fail, session):
    """
    :arg: url, msg_fail, session
    url: is the url at which raw text exists
    msg_fail: the message given to log when connection fails
    session: the session initiated

    :return:
    result: decoded bytes from source
    """
    try:
        # user-agent is just to trick the website that you are using a browser
        headers = {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:66.0) Gecko/20100101 Firefox/66.0"
        }
        result = session.get(url, headers=headers).content.decode('utf-8')
    except requests.exceptions.RequestException as e:  # This is the correct syntax
        logging.debug('URL REQUEST FAIL RETRYING: ' + url)
        logging.debug(str(e))
        raise SystemExit(e)
    return result


def createTextFile(file_name, contents):
    # Contents can be a string or a list of strings must end with \n except last element in the list
    logging.info(f"CREATING TEXT FILE: {file_name}.txt")
    with open(localappdata_path + '\\' + file_name + '.txt', "w") as text_file:
        if isinstance(contents, list):
            text_file.writelines(contents)
        else:
            text_file.write(contents)
    logging.info(f"FILE: {file_name}.txt CREATED")


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
        sorter_initialization = True
        logging.info("IP LIST LOADED")
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
    logging.info("CHECKING IF PREVIOUS RULES EXIST")
    for rule_name in servers_active_rule_list:
        if rule_name in rules:
            filtered = rule_name.split('_')[1]
            if filtered == 'ME':
                blockingLabel.config(text='ME BLOCKED', bg='#282828', fg='#ef2626', font=futrabook_font)
                logging.info(f"FOUND RULE: {rule_name}")
                return
            if filtered == 'CUSTOM':
                blockingLabel.config(text="CUSTOM BLOCK", bg='#282828', fg='#26ef4c',
                                     font=futrabook_font)
                logging.info(f"FOUND RULE: {rule_name}")
                return
            else:
                if len(filtered) < 8:
                    filtered = filtered[0:2] + ' ' + filtered[2:]
                label_text = 'PLAYING ON ' + filtered
                blockingLabel.config(text=label_text, bg='#282828', fg='#26ef4c',
                                     font=futrabook_font)
                logging.info(f"FOUND RULE: {rule_name}")
                return
    blockingLabel.config(text='ALL UNBLOCKED (DEFAULT SETTINGS)', fg='#ddee4a')
    logging.info("NO RULES FOUND")


def tunnel():  # Handle tunnelling options for Overwatch.exe
    global overwatch_path, tunnel_option
    checkbotton_state = tunnelCheckBox_state.get()
    logging.info("Check buttons state =  " + str(checkbotton_state))
    if checkbotton_state == 1:
        if exists(overwatch_path):
            logging.info("Game detected")
            createTextFile('Options', 'Tunnel=True')
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
                createTextFile('Options', ['Tunnel=True\n', overwatch_path])
            else:
                createTextFile('Options', 'Tunnel=False')
                tunnelCheckBox_state.set(0)
    else:
        createTextFile('Options', 'Tunnel=False')
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
    global ip_ranges_files, ip_range_checkboxes, customWindow
    savedSettings = []
    if exists(customConfig_path):
        logging.debug(f"{customConfig_path} FOUND")
        logging.debug(f"READING customConfig.txt")
        with open(customConfig_path, 'r') as reader:
            for line in reader.readlines():
                if len(line) > 0:
                    savedSettings.append(line.strip())
    logging.debug(f"CREATING CUSTOM WINDOW")
    customWindow = Toplevel()
    # Set Properties
    customWindow.title('Custom Config')
    customWindow.resizable(False, False)
    customWindow.geometry('300x500')
    customWindow.configure(bg='#404040')
    frameTop = Frame(customWindow, width=300, height=500)
    frameTop.pack()
    frameTop.place(x=-2, y=0)
    backgroundTop = Label(frameTop, image=CUSTOM_SETTINGS_BACKGROUND)
    backgroundTop.pack()
    customWindow.iconbitmap("LOGO_SMALL_APPLICATION.ico")
    # Buttons
    applyButton = Button(customWindow, image=button_img_APPLY_BUTTON, font=futrabook_font, command=apply,
                         bg='#404040', fg='#404040', borderwidth=0, activebackground='#404040')
    applyButton.place(x=177 + 55, y=448, anchor="center")
    resetButton = Button(customWindow, image=button_img_RESET_BUTTON, font=futrabook_font, command=resetCustomSettings,
                         bg='#404040', fg='#404040', borderwidth=0, activebackground='#404040')
    resetButton.place(x=11 + 55, y=448, anchor="center")
    openIpListButton = Button(customWindow, image=button_img_OPEN_IP_LIST_BUTTON, font=futrabook_font,
                              command=openListFolder,
                              bg='#404040', fg='#404040', borderwidth=0, activebackground='#404040')
    openIpListButton.place(x=92 + 55, y=479, anchor="center")

    # Labels
    informationLabel = Label(customWindow, text='Selected servers will be blocked', bg='#404040', fg='#ef2626',
                             font=futrabook_font)
    informationLabel.place(x=150, y=415, anchor="center")

    logging.debug(f"LISTING FILES IN APP DIRECTORY")
    onlyfiles = [f for f in listdir(localappdata_path) if isfile(join(localappdata_path, f))]
    ip_ranges_files = list()
    integersList = list()
    logging.debug(f"READING CFG FILES")
    for file in onlyfiles:
        if file.startswith('cfg'):
            ip_ranges_files.append(file)
    for i in range(0, len(ip_ranges_files)):
        integerVariable = IntVar()
        integersList.append(integerVariable)
    ip_range_checkboxes = dict(zip(ip_ranges_files, integersList))
    for index, RANGE in enumerate(ip_range_checkboxes):
        ip_range_checkboxes[RANGE] = IntVar()
        chk = Checkbutton(customWindow, text=RANGE[:-4].strip('cfg -'), font=futrabook_font,
                          activebackground='#ddee4a',
                          bg='#404040', fg='#26ef4c', borderwidth=0, variable=ip_range_checkboxes[RANGE], width=200,
                          anchor="w", selectcolor='black',
                          padx=75, pady=0.2)
        if RANGE in savedSettings:
            chk.select()
        chk.pack()


def createPing_labels(windowLevel, dictionary):
    for index, (serverName, pingValue) in enumerate(dictionary.items()):
        logging.debug(f"CREATING LABELS AT {windowLevel} LABEL: {serverName}  {pingValue}")
        serverLabel = Label(windowLevel, text=f"{serverName}", activebackground='#ddee4a',
                            bg='#404040', fg='#26ef4c', font=futrabook_font)
        serverLabel.grid(row=index + 2, column=0, sticky="EW")
        pingLabel = Label(windowLevel, text=f"{pingValue} ms", activebackground='#ddee4a',
                          bg='#404040', fg='#26ef4c', font=futrabook_font)
        pingLabel.grid(row=index + 2, column=1, sticky="EW")


def pingServers(windowLevel):
    pingList_path = localappdata_path + '\\pinglist.txt'
    severNames = list()
    serverIp = list()
    resultList = list()
    pingValueCompensation = 10  # To compensate for ICMP vs GAME ping difference
    loadingLabel = Label(windowLevel, text='Please wait for ping test...',
                         bg='#404040', fg='#ef2626', font=futrabook_font)
    loadingLabel.grid(row=1, column=0, columnspan=2, sticky="EW")
    logging.debug(f"PING PROCESS STARTED")
    if not exists(pingList_path):
        return
    logging.debug(f"{pingList_path} FOUND")
    pingList = readLine_filtered(pingList_path)

    for server in pingList:
        severNames.append(server.split('|')[0])
        serverIp.append(server.split('|')[1])
    for ip in serverIp:
        logging.debug(f"Measuring ping to {ip}")
        response_list = ping(ip, count=3, timeout=1)
        resultList.append(str(int(response_list.rtt_avg_ms) + pingValueCompensation))
        logging.debug(f"Result of ping to {ip} = {round(response_list.rtt_avg_ms, 2)}")
    pingDictionary = dict(zip(severNames, resultList))
    createPing_labels(windowLevel, pingDictionary)
    loadingLabel.grid_forget()
    logging.debug(f"PING PROCESS IS OVER")


def pingMenu():
    logging.debug(f"CREATING PING MENU")
    pingWindow = Toplevel()
    # Set Properties
    pingWindow.title('Ping')
    pingWindow.resizable(False, False)
    # pingWindow.geometry('300x500')
    pingWindow.configure(bg='#404040')
    frameTop = Frame(pingWindow, width=300, height=500)
    frameTop.pack()
    frameTop.place(x=-2, y=0)
    backgroundTop = Label(frameTop, image=CUSTOM_SETTINGS_BACKGROUND)
    backgroundTop.pack()
    pingWindow.iconbitmap("LOGO_SMALL_APPLICATION.ico")
    # Labels
    informationLabel = Label(pingWindow, text='Ping values are estimates they might be not accurate',
                             bg='#404040', fg='#ef2626', font=futrabook_font)
    informationLabel.grid(row=0, column=0, columnspan=2, sticky="EW")
    # informationLabel.place(x=150, y=415, anchor="center")
    logging.debug(f"LAUNCHING pingServers THREAD WITH args: pingWindow")
    threading.Thread(target=pingServers, args=(pingWindow,), daemon=True).start()  # Follow main thread


def resetCustomSettings():
    logging.debug(f"RESETTING CUSTOM CONFIG")
    customConfig.clear()
    if exists(customConfig_path):
        logging.debug(f"REMOVING PATH {customConfig_path}")
        remove(customConfig_path)
    logging.debug(f"CLOSING --> CUSTOM WINDOW")
    customWindow.destroy()


def apply():
    global customIpRanges
    customIpRanges.clear()
    for IP_NAME in ip_ranges_files:
        if ip_range_checkboxes[IP_NAME].get() == 1:
            logging.debug(f"CUSTOM CONFIG STATES: {IP_NAME} CHECKBOX STATE: {str(ip_range_checkboxes[IP_NAME].get())}")
            if IP_NAME not in customIpRanges:
                customIpRanges.append(IP_NAME)
    with open(localappdata_path + '\\customConfig.txt', 'w') as fp:
        for item in customIpRanges:
            # write each item on a new line
            fp.write("%s\n" % item)
    loadUserConfig()
    customWindow.destroy()


def openListFolder():
    if exists(localappdata_path):
        startfile(localappdata_path)


def reinstallIp_list():
    logging.debug(f"REINSTALLING IP LIST")
    internetLabel.config(text="REINSTALLING IP LIST", fg='#ef2626')
    pingButton.place_forget()
    if exists(ip_version_path):
        remove(ip_version_path)
    updateIpList()
    if not appUpdate_required:
        pingButton.place(x=135, y=500, height=40, width=230)


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


def discordInvite():
    webbrowser.open("https://discord.gg/8CtV7bkJzB")


def checkUpdate():
    """
    Used to check for update for both  IP and APP version.

    :return: None
    """
    check_ip_update()
    check_app_update()


# Menus
menu = Menu(app)
app.config(menu=menu)
options_menu = Menu(menu)
menu.add_cascade(label="Options", menu=options_menu)
options_menu.add_command(label="Open config folder", command=openListFolder)
options_menu.add_command(label="Check for updates", command=lambda: threading.Thread(target=checkUpdate).start())
# lambda to create the thread each time the option is pressed
options_menu.add_command(label="Reinstall IP List", command=lambda: threading.Thread(target=reinstallIp_list).start())
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
DonationButton.place(x=254, y=550, height=28, width=110)

DiscordButton = Button(app, image=button_img_Discord_Button, font=futrabook_font, command=discordInvite,
                       bg='#282828', fg='#282828', borderwidth=0, activebackground='#282828')
DiscordButton.place(x=136, y=550, height=28, width=110)

CustomSettingsButton = Button(app, image=button_img_CUSTOM_SETTINGS, font=futrabook_font, command=customSettingsWindow,
                              bg='#282828', fg='#282828', borderwidth=0, activebackground='#282828')
CustomSettingsButton.place(x=370, y=318, height=25, width=25)

InstallUpdateButton = Button(app, image=button_img_INSTALL_UPDATE, font=futrabook_font, command=installUpdate,
                             bg='#282828', fg='#282828', borderwidth=0, activebackground='#282828')

pingButton = Button(app, image=button_img_TEST_PING, font=futrabook_font, command=pingMenu,
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

loadIpRanges_thread = threading.Thread(target=loadIpRanges, daemon=True).start()  # Follow main thread
checkIfActive_thread = threading.Thread(target=checkIfActive, daemon=True).start()  # Follow main thread
check_for_update_thread = threading.Thread(target=checkUpdate, daemon=True).start()  # Follow main thread

app.mainloop()
