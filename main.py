import time
from tkinter import *
from tkinter.font import Font
from tkinter import ttk, filedialog
import win32com.shell.shell as shell
from subprocess import call, Popen, CREATE_NEW_CONSOLE, PIPE, run, STARTUPINFO, STARTF_USESHOWWINDOW, SW_HIDE
from PIL import ImageTk, Image
from io import BytesIO
import pic2str
import base64
from os.path import exists, isdir
from os import getenv, path, mkdir, listdir, system, linesep
import webbrowser
import socket
import threading
import requests

# Create window object
app = Tk()
# Set Properties
app.title('MINA Overwatch 2 Server Selector')
app.resizable(False, False)
app.geometry('500x500')
app.configure(bg='#282828')
frame = Frame(app, width=500, height=94)
frame.pack()
frame.place(x=-2, y=0)

# For images
# Load byte data
byte_LOGO_SMALL_APPLICATION = base64.b64decode(pic2str.LOGO_SMALL_APPLICATION)
byte_SQUARE_BACKGROUND_MINA_TEST = base64.b64decode(pic2str.SQUARE_BACKGROUND_MINA_TEST)
byte_play_on_eu = base64.b64decode(pic2str.play_on_eu)
byte_play_on_na_east = base64.b64decode(pic2str.play_on_na_east)
byte_play_on_na_west = base64.b64decode(pic2str.play_on_na_west)
byte_BLOCK_MIDDLE_EAST = base64.b64decode(pic2str.BLOCK_MIDDLE_EAST)
byte_play_on_australia = base64.b64decode(pic2str.play_on_australia)
byte_donation = base64.b64decode(pic2str.Donation_Button)
byte_UNBLOCK_ALL_MAIN = base64.b64decode(pic2str.UNBLOCK_ALL_MAIN)

image_data_SMALL_APPLICATION = BytesIO(byte_LOGO_SMALL_APPLICATION)
image_data_SQUARE_BACKGROUND_MINA_TEST = BytesIO(byte_SQUARE_BACKGROUND_MINA_TEST)
image_play_on_eu = BytesIO(byte_play_on_eu)
image_play_on_na_east = BytesIO(byte_play_on_na_east)
image_play_on_na_west = BytesIO(byte_play_on_na_west)
image_BLOCK_MIDDLE_EAST = BytesIO(byte_BLOCK_MIDDLE_EAST)
image_play_on_australia = BytesIO(byte_play_on_australia)
image_donation = BytesIO(byte_donation)
image_UNBLOCK_ALL_MAIN = BytesIO(byte_UNBLOCK_ALL_MAIN)

# Add images
background = ImageTk.PhotoImage(Image.open(image_data_SQUARE_BACKGROUND_MINA_TEST))
logo = Label(frame, image=background)
logo.pack()
button_img_ME = ImageTk.PhotoImage(Image.open(image_BLOCK_MIDDLE_EAST))
button_img_EU = ImageTk.PhotoImage(Image.open(image_play_on_eu))
button_img_NA_WEST = ImageTk.PhotoImage(Image.open(image_play_on_na_west))
button_img_ME_EAST = ImageTk.PhotoImage(Image.open(image_play_on_na_east))
button_img_Australia = ImageTk.PhotoImage(Image.open(image_play_on_australia))
button_img_donation = ImageTk.PhotoImage(Image.open(image_donation))
button_img_Default = ImageTk.PhotoImage(Image.open(image_UNBLOCK_ALL_MAIN))

# Add font
futrabook_font = Font(family="Futura PT Demi", size=10)

# Global variables
localappdata_path = getenv('APPDATA') + '\\OverwatchServerBlocker'
ip_version_path = localappdata_path + '\\IP_version.txt'
overwatch_path = 'C:\\Program Files (x86)\\Overwatch\\_retail_\\Overwatch.exe'
ip_version_url = 'https://raw.githubusercontent.com/foryVERX/Overwatch-Server-Selector-1/main/ip_lists/IP_version.txt'
Ip_ranges_ME_url = 'https://raw.githubusercontent.com/foryVERX/Overwatch-Server-Selector-1/main/ip_lists/Ip_ranges_ME.txt'
Ip_ranges_EU_url = 'https://raw.githubusercontent.com/foryVERX/Overwatch-Server-Selector-1/main/ip_lists/Ip_ranges_EU.txt'
Ip_ranges_NA_East_url = 'https://raw.githubusercontent.com/foryVERX/Overwatch-Server-Selector-1/main/ip_lists/Ip_ranges_NA_East.txt'
Ip_ranges_NA_central_url = 'https://raw.githubusercontent.com/foryVERX/Overwatch-Server-Selector-1/main/ip_lists/Ip_ranges_NA_central.txt'
Ip_ranges_NA_West_url = 'https://raw.githubusercontent.com/foryVERX/Overwatch-Server-Selector-1/main/ip_lists/Ip_ranges_NA_West.txt'
Ip_ranges_AS_Korea_url = 'https://raw.githubusercontent.com/foryVERX/Overwatch-Server-Selector-1/main/ip_lists/Ip_ranges_AS_Korea.txt'
Ip_ranges_AS_1_url = 'https://raw.githubusercontent.com/foryVERX/Overwatch-Server-Selector-1/main/ip_lists/Ip_ranges_AS_1.txt'
Ip_ranges_AS_Singapore_url = 'https://raw.githubusercontent.com/foryVERX/Overwatch-Server-Selector-1/main/ip_lists/Ip_ranges_AS_Singapore.txt'
Ip_ranges_AS_Taiwan_url = 'https://raw.githubusercontent.com/foryVERX/Overwatch-Server-Selector-1/main/ip_lists/Ip_ranges_AS_Taiwan.txt'
Ip_ranges_AS_Japan_url = 'https://raw.githubusercontent.com/foryVERX/Overwatch-Server-Selector-1/main/ip_lists/Ip_ranges_AS_Japan.txt'
Ip_ranges_Australia_url = 'https://raw.githubusercontent.com/foryVERX/Overwatch-Server-Selector-1/main/ip_lists/Ip_ranges_Australia.txt'
Ip_ranges_Brazil_url = 'https://raw.githubusercontent.com/foryVERX/Overwatch-Server-Selector-1/main/ip_lists/Ip_ranges_Brazil.txt'
BlockingConfig_url = 'https://raw.githubusercontent.com/foryVERX/Overwatch-Server-Selector-1/main/ip_lists/BlockingConfig.txt'

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


# Functions
def iconMaker():  # Used to check if there is an icon in the same directory or not it will create the icon if not.
    if exists("LOGO_SMALL_APPLICATION.ico"):
        app.iconbitmap("LOGO_SMALL_APPLICATION.ico")
    else:
        icon = Image.open(image_data_SMALL_APPLICATION)
        icon.save("LOGO_SMALL_APPLICATION.ico")
        app.iconbitmap("LOGO_SMALL_APPLICATION.ico")


def ipSorter():  # Store ip ranges from Ip_ranges_....txt into Ip_ranges dictionary
    global sorter_initialization
    if exists(localappdata_path) and exists(ip_version_path):  # If those paths exists it means user updated
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
                            server = path.splitext(server)[0]
                            temp_list.append(line)
                            temp_list.append(',')
                    Ip_ranges_dic[server] = temp_list
        update_text = "UPDATED"
        app.after(250, internetLabel.config(text=update_text, fg='#26ef4c'))
        sorter_initialization = True
        print("IP LIST SORTED")
    else:  # User running first time
        checkUpdate('ipSorter')
    controlButtons('normal')


def controlButtons(command):  # 'disabled' or 'normal' buttons
    PlayMEButton['state'] = command

    PlayEUButton['state'] = command

    PlayNAWESTButton['state'] = command

    PlayNAEASTButton['state'] = command

    PlayAustraliaButton['state'] = command

    ClearBlocksButton['state'] = command

    DonationButton['state'] = command


def is_connect():
    global internetConnection, internet_initialization
    try:
        socket.create_connection(("www.google.com", 80))
        internetConnection = True
    except OSError:
        internetConnection = False


def request_raw_file(url, msg_fail):
    try:
        r = requests.get(url).content.decode('utf-8')
    except requests.exceptions.RequestException as e:  # This is the correct syntax
        update_text = msg_fail
        app.after(250, internetLabel.config(text=update_text, fg='#ef2626'))
        checkUpdate()
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
        progressBar.place(x=360, y=465)
        progressBar['value'] += 10


def updateIp():
    global updating_state, sorter_initialization, checkForUpdate_initialization
    controlButtons('disabled')
    if internetConnection:
        updating_state = True
        update_text = "UPDATING..."
        internetLabel.config(text=update_text, fg='#ddee4a')
        msg_fail = "CONNECTION FAILED... Trying to update"
        createTextFile('BlockingConfig', request_raw_file(BlockingConfig_url, msg_fail), progressbar=True)
        createTextFile('Ip_ranges_EU', request_raw_file(Ip_ranges_EU_url, msg_fail), progressbar=True)
        createTextFile('Ip_ranges_ME', request_raw_file(Ip_ranges_ME_url, msg_fail), progressbar=True)
        createTextFile('Ip_ranges_AS_Singapore', request_raw_file(Ip_ranges_AS_Singapore_url, msg_fail),
                       progressbar=True)
        createTextFile('Ip_ranges_Australia', request_raw_file(Ip_ranges_Australia_url, msg_fail), progressbar=True)
        createTextFile('Ip_ranges_Brazil', request_raw_file(Ip_ranges_Brazil_url, msg_fail), progressbar=True)
        createTextFile('Ip_ranges_NA_East', request_raw_file(Ip_ranges_NA_East_url, msg_fail), progressbar=True)
        createTextFile('Ip_ranges_NA_West', request_raw_file(Ip_ranges_NA_West_url, msg_fail), progressbar=True)
        createTextFile('Ip_ranges_NA_central', request_raw_file(Ip_ranges_NA_central_url, msg_fail), progressbar=True)
        createTextFile('Ip_ranges_AS_Japan', request_raw_file(Ip_ranges_AS_Japan_url, msg_fail), progressbar=True)
        createTextFile('Ip_ranges_AS_1', request_raw_file(Ip_ranges_AS_1_url, msg_fail), progressbar=True)
        createTextFile('Ip_ranges_AS_Taiwan', request_raw_file(Ip_ranges_AS_Taiwan_url, msg_fail), progressbar=True)
        createTextFile('Ip_ranges_AS_Korea', request_raw_file(Ip_ranges_AS_Korea_url, msg_fail), progressbar=True)
        createTextFile('IP_version', request_raw_file(ip_version_url, msg_fail), progressbar=True)
        progressBar.lower()
        print("Updated")
        updating_state = False
        checkForUpdate_initialization = False
        ipSorter()
    else:
        update_text = "Please check your internet connection to download servers ip"
        internetLabel.config(text=update_text, fg='#ddee4a')


def checkUpdate(thread_type='mainThread'):  # A function called at the start of the program to check for update
    global isUpdated, updating_state, checkForUpdate_initialization
    is_connect()
    print("Checking for updates")
    if internetConnection:
        if isdir(localappdata_path) and path.exists(ip_version_path):
            with open(ip_version_path, "r") as reader:  # Read Ip_version.txt from GitHub and analyze
                for line in reader.readlines():
                    if len(line) > 1:
                        msg_fail = "Update check failed"
                        ip_version_request = request_raw_file(ip_version_url, msg_fail)
                        ip_version_request = linesep.join([s for s in ip_version_request.splitlines() if s])
                        line = linesep.join([s for s in line.splitlines() if s])
                        print("Line length: ", len(line))
                        print("ip length: ", len(ip_version_request))
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
        if not exists(ip_version_path):
            controlButtons('disabled')
            update_text = "CONNECTION FAILED... Trying to update"
            app.after(250, internetLabel.config(text=update_text, fg='#ef2626'))
            app.after(1000, checkUpdate)
        else:
            update_text = "NO INTERNET MIGHT BE NOT LATEST IP LIST VERSION"
            app.after(250, internetLabel.config(text=update_text, fg='#ddee4a'))
    if thread_type == 'mainThread':  # If the function is called from main thread call it again after 5 mints
        app.after(5000 * 60, checkUpdate)


def ruleMakerBlock(server_exception, np_ips, block_exception=True, rule_name='@Overwatch Block'):
    # Used to block IP range
    # np_ips is number of ip ranges that included in one function
    # server_exception is the only server to not block can be a list or string
    # If block_exception set to false then the server_exception is blocked ONLY
    controlButtons('disabled')
    x = 0
    temp_ip_ranges = []
    size_of_ip_range = 0
    if not block_exception:
        for server in Ip_ranges_dic:
            if (server in server_exception):
                for ip in Ip_ranges_dic[server]:
                    temp_ip_ranges.append(ip)
                blockIpRange(temp_ip_ranges, rule_name)
                print("One rule created")
                checkIfActive()
                controlButtons('normal')
                return
    for server in Ip_ranges_dic:
        if not (server in server_exception):
            for ip in Ip_ranges_dic[server]:
                temp_ip_ranges.append(ip)
                size_of_ip_range += 1
            temp_ip_ranges.append(',')
    size_of_ip_range = int(size_of_ip_range / 2)
    print(size_of_ip_range)
    if size_of_ip_range <= np_ips:
        blockIpRange(temp_ip_ranges, rule_name)
        print("One rule created")
    else:
        temp_ip_ranges.clear()
        for indexServer, server in enumerate(Ip_ranges_dic):
            if not (server in server_exception):
                print("Server: ", server, "is not: ", server_exception)
                for indexIp, ip in enumerate(Ip_ranges_dic[server]):
                    temp_ip_ranges.append(ip)
                    if int(len(temp_ip_ranges) / 2) == np_ips:  # /2 because each range separated by ','
                        x += 1
                        blockIpRange(temp_ip_ranges, rule_name)
                        temp_ip_ranges.clear()
                    if indexServer == len(Ip_ranges_dic) - 1 and indexIp == len(Ip_ranges_dic[server]) - 1:
                        x += 1
                        print("Last index")
                        blockIpRange(temp_ip_ranges, rule_name)
        print(str(x) + " Rules created")
        checkIfActive()
        controlButtons('normal')


def blockIpRange(ip_list, rule_name):
    ip_string = ''.join(ip_list)
    if ip_string[1:] == ',':
        ip_string = ip_string[1:]
    if ip_string[len(ip_string) - 1] == ',':
        ip_string = ip_string[:-1]
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


def ruleDelete(rule_name):  # Delete rule by exact name, name must be a string '' or list of strings
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
                                'EU_OW_SERVER_BLOCKER', 'AU_OW_SERVER_BLOCKER', 'Australia_OW_SERVER_BLOCKER']
    command_list = ['netsh', 'advfirewall', 'firewall', 'show', 'rule', 'name=all']

    output = run(command_list, capture_output=True, text=True)
    output = str(output.stdout)
    for rule_name in servers_active_rule_list:
        rules_existence = output.find(rule_name)
        if rules_existence > 0:
            filtered = rule_name.rpartition('_')[0].replace(" ", "").replace("RuleName:@", "").replace("_OW_SERVER",
                                                                                                       "")
            if filtered == 'ME':
                blockingLabel.config(text='ME BLOCKED', bg='#282828', fg='#ef2626', font=futrabook_font)
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
    print("Check buttons state =  ", str(checkbotton_state))
    if checkbotton_state == 1:
        if exists(overwatch_path):
            print("Game detected")
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
                print("Overwatch path is:  " + overwatch_path)
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
                        print("String blocking server: ", server_name, "want to exclude", temp_block_config_list)
                        blockingConfigDic[server_name] = temp_block_config_list


def blockALL():  # This function is for testing reasons only DO NOT USE.
    unblockALL()
    blockingLabel.config(text='ALL BLOCKED', fg='#ef2626')


def blockMEServer():  # It removes any rules added by blockserver function
    unblockALL()
    blockingLabel.config(text='WORKING ON IT', fg='#26ef4c')
    commands = 'advfirewall firewall add rule name="@ME_OW_SERVER_BLOCKER" Dir=Out Action=Allow'
    shell.ShellExecuteEx(lpVerb='runas', lpFile='netsh.exe', lpParameters=commands)
    threading.Thread(target=ruleMakerBlock, args=(blockingConfigDic['Ip_ranges_ME'], 467,),
                     daemon=True, kwargs={'block_exception': False}).start()  # Follow main thread


def PlayAustralia_server():
    unblockALL()
    blockingLabel.config(text='WORKING ON IT', fg='#26ef4c')
    commands = 'advfirewall firewall add rule name="@Australia_OW_SERVER_BLOCKER" Dir=Out Action=Allow'
    shell.ShellExecuteEx(lpVerb='runas', lpFile='netsh.exe', lpParameters=commands)
    threading.Thread(target=ruleMakerBlock, args=(blockingConfigDic['Ip_ranges_Australia'], 467,),
                     daemon=True).start()  # Follow main thread


def playNAEast_server():
    unblockALL()
    blockingLabel.config(text='WORKING ON IT', fg='#26ef4c')
    commands = 'advfirewall firewall add rule name="@NAEAST_OW_SERVER_BLOCKER" Dir=Out Action=Allow'
    shell.ShellExecuteEx(lpVerb='runas', lpFile='netsh.exe', lpParameters=commands)
    threading.Thread(target=ruleMakerBlock, args=(blockingConfigDic['Ip_ranges_NA_East'], 467,),
                     daemon=True).start()  # Follow main thread


def playNAWest_server():
    unblockALL()
    blockingLabel.config(text='WORKING ON IT', fg='#26ef4c')
    commands = 'advfirewall firewall add rule name="@NAWEST_OW_SERVER_BLOCKER" Dir=Out Action=Allow'
    shell.ShellExecuteEx(lpVerb='runas', lpFile='netsh.exe', lpParameters=commands)
    threading.Thread(target=ruleMakerBlock, args=(blockingConfigDic['Ip_ranges_NA_West'], 467,),
                     daemon=True).start()  # Follow main thread


def playEU_server():
    unblockALL()
    blockingLabel.config(text='WORKING ON IT', fg='#26ef4c')
    commands = 'advfirewall firewall add rule name="@EU_OW_SERVER_BLOCKER" Dir=Out Action=Allow'
    shell.ShellExecuteEx(lpVerb='runas', lpFile='netsh.exe', lpParameters=commands)
    threading.Thread(target=ruleMakerBlock, args=(blockingConfigDic['Ip_ranges_EU'], 467,),
                     daemon=True).start()  # Follow main thread


def unblockALL():
    blockingLabel.config(text='ALL UNBLOCKED (DEFAULT SETTINGS)', fg='#ddee4a')
    list_rule_names = ["@NAEAST_OW_SERVER_BLOCKER", "@EU_OW_SERVER_BLOCKER", "@ME_OW_SERVER_BLOCKER",
                       "@NAWEST1_OW_SERVER_BLOCKER", "@AU_OW_SERVER_BLOCKER", "@NAWEST2_OW_SERVER_BLOCKER",
                       "@Overwatch Block", "@NAWEST_OW_SERVER_BLOCKER", "@Australia_OW_SERVER_BLOCKER"]
    ruleDelete(list_rule_names)


def donationPage():
    webbrowser.open("https://paypal.me/vantverx?country.x=SA&locale.x=en_US")


# Labels
blockingLabel = Label(app, text='', bg='#282828', fg='#ddee4a', font=futrabook_font)
blockingLabel.grid(row=0, column=0)
blockingLabel.place(x=250, y=440, anchor="center")

internetLabel = Label(app, text='', bg='#282828', fg='#26ef4c', font=futrabook_font)
internetLabel.grid(row=0, column=0)
internetLabel.place(x=250, y=420, anchor="center")

progressBar = ttk.Progressbar(app, orient=HORIZONTAL, length=130, mode='determinate')

# Buttons
PlayMEButton = Button(app, image=button_img_ME, font=futrabook_font, command=blockMEServer,
                      bg='#282828', fg='#282828', borderwidth=0, activebackground='#282828')
PlayMEButton.place(x=135, y=70, height=40, width=230)

PlayEUButton = Button(app, image=button_img_EU, font=futrabook_font, command=playEU_server,
                      bg='#282828', fg='#282828', borderwidth=0, activebackground='#282828')
PlayEUButton.place(x=135, y=120, height=40, width=230)

PlayNAWESTButton = Button(app, image=button_img_NA_WEST, font=futrabook_font, command=playNAWest_server,
                          bg='#282828', fg='#282828', borderwidth=0, activebackground='#282828')
PlayNAWESTButton.place(x=135, y=180, height=40, width=230)

PlayNAEASTButton = Button(app, image=button_img_ME_EAST, font=futrabook_font, command=playNAEast_server,
                          bg='#282828', fg='#282828', borderwidth=0, activebackground='#282828')
PlayNAEASTButton.place(x=135, y=240, height=40, width=230)

PlayAustraliaButton = Button(app, image=button_img_Australia, font=futrabook_font, command=PlayAustralia_server,
                             bg='#282828', fg='#282828', borderwidth=0, activebackground='#282828')
PlayAustraliaButton.place(x=135, y=300, height=40, width=230)

ClearBlocksButton = Button(app, image=button_img_Default, font=futrabook_font, command=unblockALL,
                           bg='#282828', fg='#282828', borderwidth=0, activebackground='#282828')
ClearBlocksButton.place(x=135, y=360, height=40, width=230)

DonationButton = Button(app, image=button_img_donation, font=futrabook_font, command=donationPage,
                        bg='#282828', fg='#282828', borderwidth=0, activebackground='#282828')
DonationButton.place(x=420, y=360, height=73, width=68)

# Check box
tunnelCheckBox_state = IntVar()
tunnelCheckBox = Checkbutton(app, text="Only affect Overwatch ", font=futrabook_font, activebackground='#282828',
                             bg='#282828', fg='#26ef4c', borderwidth=0, variable=tunnelCheckBox_state, command=tunnel)

# Start Program
iconMaker()
ipSorter_thread = threading.Thread(target=ipSorter, daemon=True)  # Follow main thread
ipSorter_thread.start()

checkIfActive_thread = threading.Thread(target=checkIfActive, daemon=True).start()  # Follow main thread
checkUpdate.thread = threading.Thread(target=checkUpdate, daemon=True).start()  # Follow main thread

tunnelCheckBox.place(x=0, y=475)
if checkOptions():
    tunnelCheckBox.select()

app.mainloop()
