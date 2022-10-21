from tkinter import *
from tkinter.font import Font
import win32com.shell.shell as shell
from subprocess import Popen, CREATE_NEW_CONSOLE, PIPE, STARTUPINFO, STARTF_USESHOWWINDOW, SW_HIDE
from PIL import ImageTk, Image
from io import BytesIO
import pic2str
import base64
from os.path import exists
from ping3 import ping
import webbrowser

# Create window object
app = Tk()
# Set Properties
app.title('MINA Overwatch 2 Server Selector Beta Version 3.0')
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

# Ip ranges
Ip_ranges_ME = "157.175.0.0-157.175.255.255,15.185.0.0-15.185.255.255,15.184.0.0-15.184.255.255"
Ip_ranges_EU1 = "5.42.184.0-5.42.191.255,5.42.168.0-5.42.175.255"
Ip_ranges_EU2 = "35.198.64.0-35.198.191.255,34.107.0.0-34.107.127.255,35.195.0.0-35.195.255.255,35.246.0.0-35.246.255.255,35.228.0.0-35.228.255.255,34.89.128.0-34.89.255.255,35.242.128.0-35.242.255.255,34.159.0.0-34.159.255.255,34.141.0.0-34.141.127.255,34.88.0.0-34.88.255.255"
Ip_ranges_NA_East = "35.236.192.0-35.236.255.255,35.199.0.0-35.199.63.255,34.86.0.0-34.86.255.255,35.245.0.0-35.245.255.255,35.186.160.0-35.186.191.255,34.145.128.0-34.145.255.255,34.150.128.0-34.150.255.255,34.85.128.0-34.85.255.255"
Ip_ranges_NA_central = "24.105.40.0-24.105.47.255"
Ip_ranges_NA_West1 = "24.105.8.0-24.105.15.255"
Ip_ranges_NA_West2 = "35.247.0.0-35.247.127.255,35.236.0.0-35.236.127.255,35.235.70.0-35.235.130.255,34.102.0.0-34.102.127.255,34.94.0.0-34.94.255.255"
Ip_ranges_NA_West3 = "34.19.0.0-34.19.127.255,34.82.0.0-34.83.255.255,34.105.0.0-34.105.127.255,34.118.192.0-34.118.199.255,34.127.0.0-34.127.127.255,34.145.0.0-34.145.127.255,34.157.112.0-34.157.119.255,34.157.240.0-34.157.247.255,34.168.0.0-34.169.255.255,35.185.192.0-35.185.255.255,35.197.0.0-35.197.127.255,35.199.144.0-35.199.159.255,35.199.160.0-35.199.191.255,35.203.128.0-35.203.191.255,35.212.128.0-35.212.255.255,35.220.48.0-35.220.55.255,35.227.128.0-35.227.191.255,35.230.0.0-35.230.127.255,35.233.128.0-35.233.255.255,35.242.48.0-35.242.55.255,35.243.32.0-35.243.39.255,35.247.0.0-35.247.127.255"
Ip_ranges_AS_Korea = "34.64.0.0-34.64.255.255,117.52.0.0-117.52.255.255"
Ip_ranges_AS_1 = "104.198.0.0-104.198.255.255,34.84.0.0-34.84.255.255,34.85.0.0-34.85.255.255,35.200.0.0-35.200.255.255,35.221.0.0-35.221.255.255,34.146.0.0-34.146.255.255,117.52.0.0-117.52.255.255,121.254.0.0-121.254.255.255,5.42.0.0-5.42.255.255,34.87.0.0-34.87.255.255,34.126.0.0-34.126.255.255,35.187.0.0-35.187.255.255,37.244.42.0-37.244.42.255,34.142.0.0-34.143.255.255"
Ip_ranges_AS_Singapore1 = "34.124.0.0-34.124.255.255,34.124.42.0-34.124.43.255,34.142.128.0-34.142.255.255,35.185.176.0-35.185.191.255,35.186.144.0-35.186.159.255,35.247.128.0-35.247.191.255,34.87.0.0-34.87.191.255,34.143.128.0-34.143.255.255,34.124.128.0-34.124.255.255,34.126.64.0-34.126.191.255,35.240.128.0-35.240.255.255,35.198.192.0-35.198.255.255"
Ip_ranges_AS_Singapore2 = "34.21.128.0-34.21.255.255,34.87.0.0-34.87.191.255,34.104.58.0-34.104.59.255,34.124.41.0-34.124.42.255,34.124.128.0-34.124.255.255,34.126.64.0-34.126.191.255,34.157.82.0-34.157.83.255,34.157.88.0-34.157.89.255,34.157.210.0-34.157.211.255,35.187.224.0-35.187.255.255,35.197.128.0-35.197.159.255,35.198.192.0-35.198.255.255,35.213.128.0-35.213.191.255,35.220.24.0-35.220.25.255,35.234.192.0-35.234.207.255,35.240.128.0-35.240.255.255,35.242.24.0-35.242.25.255,35.247.128.0-35.247.191.255"
Ip_ranges_AS_Taiwan = "5.42.160.0-5.42.160.255"
Ip_ranges_AS_Japan = "34.85.0.0-34.85.127.255,34.84.0.0-34.84.255.255,35.190.224.0-35.190.239.255,35.194.96.0-35.194.255.255,35.221.64.0-35.221.255.255,34.146.0.0-34.146.255.255"
Ip_ranges_oc = "37.244.42.0-37.244.42.255"
Ip_ranges_SA = "34.151.0.0-34.151.255.255,34.95.128.0-34.95.255.255,35.198.0.0-35.198.63.255,35.247.192.0-35.247.255.255,35.199.64.0-35.199.127.255"


# NA CENTRAL CAN'T QUEUE
#

# Functions
def iconMaker():  # Used to check if there is an icon in the same directory or not it will create the icon if not.
    if exists("LOGO_SMALL_APPLICATION.ico"):
        app.iconbitmap("LOGO_SMALL_APPLICATION.ico")
    else:
        icon = Image.open(image_data_SMALL_APPLICATION)
        icon.save("LOGO_SMALL_APPLICATION.ico")
        app.iconbitmap("LOGO_SMALL_APPLICATION.ico")


def pingServers():  # Return ping to all regions
    server_list = ["24.105.30.129", "24.105.62.129", "185.60.114.159", "au-syd-speedtest01.urlnetworks.net"]
    ping_list = []
    for ip in server_list:
        result = ping(ip, unit="ms")
        ping_list.append(int(result))
    na_west_ping = str(ping_list[0])
    na_central_ping = str(ping_list[1])
    eu_ping = str(ping_list[2])
    australia_ping = str(ping_list[3])
    pingtext = "NA West: " + na_west_ping + " ms" + " | NA Central: " + na_central_ping + " ms" + " | EU: " \
               + eu_ping + " ms" + " | AU_Syd: " + australia_ping + " ms"
    pingLabel.config(text=pingtext, fg='#26ef4c')
    return na_west_ping, na_central_ping, eu_ping


def ruleMakerBlock(*argv, rule_name="@Overwatch Block"):  # Used to block IP range
    ip_range = ""
    for arg in argv:
        if len(argv) > 1:
            ip_range += arg + ","
        else:
            ip_range = arg
    commands = 'advfirewall firewall add rule name="' + rule_name + '" Dir=In Action=Block RemoteIP=' + ip_range
    shell.ShellExecuteEx(lpVerb='runas', lpFile='netsh.exe', lpParameters=commands)
    commands = 'advfirewall firewall add rule name="' + rule_name + '" Dir=Out Action=Block RemoteIP=' + ip_range
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
    # pingServers()
    servers_active_rule_list = ['"@ME_OW_SERVER_BLOCKER"', '"@NAEAST_OW_SERVER_BLOCKER"', '"@NAWEST_OW_SERVER_BLOCKER"',
                                '"@EU_OW_SERVER_BLOCKER"', '"@AU_OW_SERVER_BLOCKER"']
    for rule in servers_active_rule_list:
        command = 'netsh advfirewall firewall show rule name=' + rule
        proc = Popen(command, creationflags=CREATE_NEW_CONSOLE, stdout=PIPE)
        output = proc.communicate()[0]
        rules_existence = str(output)
        rules_existence = rules_existence.find('Rule Name:')  # To avoid processing useless bytes
        if rules_existence == 6:  # Position of the Rule Name string
            output = str(output.strip().decode("utf-8"))
            temp_rule = rule.replace('"', "")
            if temp_rule in output:
                filtered = output.rpartition('_')[0].replace(" ", "").replace("RuleName:@", "").replace("_OW_SERVER",
                                                                                                        "")
                if filtered == 'ME':
                    blockingLabel.config(text='ME BLOCKED', bg='#282828', fg='#ef2626', font=futrabook_font)
                    break
                else:
                    filtered = filtered[0:2] + ' ' + filtered[2:]
                    label_text = 'PLAYING ON ' + filtered
                    blockingLabel.config(text=label_text, bg='#282828', fg='#26ef4c',
                                         font=futrabook_font)
                    break
        blockingLabel.config(text='ALL UNBLOCKED (DEFAULT SETTINGS)', fg='#ddee4a')


def blockALL():  # This function is for testing reasons only DO NOT USE.
    unblockALL()
    blockingLabel.config(text='ALL BLOCKED', fg='#ef2626')

    # Block ALL


def blockMEServer():  # It removes any rules added by blockserver function
    unblockALL()
    blockingLabel.config(text='ME BLOCKED', fg='#ef2626')
    # commands = 'advfirewall firewall add rule name="@ME_OW_SERVER_BLOCKER" Dir=Out Action=Allow RemoteIP=' + Ip_ranges_EU
    # shell.ShellExecuteEx(lpVerb='runas', lpFile='netsh.exe', lpParameters=commands)

    # Block ME
    ruleMakerBlock(Ip_ranges_ME)


def PlayAustralia_server():
    unblockALL()
    blockingLabel.config(text='PLAYING ON Australia', fg='#26ef4c')
    commands = 'advfirewall firewall add rule name="@AU_OW_SERVER_BLOCKER" Dir=Out Action=Allow RemoteIP=' + Ip_ranges_oc
    shell.ShellExecuteEx(lpVerb='runas', lpFile='netsh.exe', lpParameters=commands)

    ruleMakerBlock(Ip_ranges_ME, Ip_ranges_EU2, Ip_ranges_AS_1, Ip_ranges_EU1, Ip_ranges_NA_East, Ip_ranges_NA_West1,
                   Ip_ranges_NA_West2,
                   Ip_ranges_NA_West3
                   , Ip_ranges_AS_Japan, Ip_ranges_AS_Korea, Ip_ranges_AS_Taiwan,
                   Ip_ranges_SA, Ip_ranges_AS_Singapore1, Ip_ranges_NA_central, Ip_ranges_AS_Singapore2)


def playNAEast_server():
    unblockALL()
    blockingLabel.config(text='PLAYING ON NA EAST', fg='#26ef4c')
    commands = 'advfirewall firewall add rule name="@NAEAST_OW_SERVER_BLOCKER" Dir=Out Action=Allow RemoteIP=' + Ip_ranges_NA_East
    shell.ShellExecuteEx(lpVerb='runas', lpFile='netsh.exe', lpParameters=commands)

    # Block ME, EU, NA West, AS
    ruleMakerBlock(Ip_ranges_ME, Ip_ranges_EU2, Ip_ranges_AS_1, Ip_ranges_EU1, Ip_ranges_NA_West1, Ip_ranges_NA_West2,
                   Ip_ranges_NA_West3
                   , Ip_ranges_AS_Japan, Ip_ranges_AS_Korea, Ip_ranges_AS_Taiwan,
                   Ip_ranges_SA, Ip_ranges_AS_Singapore1, Ip_ranges_AS_Singapore2, Ip_ranges_NA_central, Ip_ranges_oc)


def playNAWest_server():
    unblockALL()
    blockingLabel.config(text='PLAYING ON NA WEST', fg='#26ef4c')
    commands = 'advfirewall firewall add rule name="@NAWEST1_OW_SERVER_BLOCKER" Dir=Out Action=Allow RemoteIP=' + Ip_ranges_NA_West1
    shell.ShellExecuteEx(lpVerb='runas', lpFile='netsh.exe', lpParameters=commands)
    commands = 'advfirewall firewall add rule name="@NAWEST2_OW_SERVER_BLOCKER" Dir=Out Action=Allow RemoteIP=' + Ip_ranges_NA_West2
    shell.ShellExecuteEx(lpVerb='runas', lpFile='netsh.exe', lpParameters=commands)

    ruleMakerBlock(Ip_ranges_ME, Ip_ranges_EU2, Ip_ranges_AS_1, Ip_ranges_EU1, Ip_ranges_AS_Japan, Ip_ranges_AS_Korea,
                   Ip_ranges_AS_Taiwan,
                   Ip_ranges_SA, Ip_ranges_AS_Singapore1, Ip_ranges_AS_Singapore2, Ip_ranges_NA_central,
                   Ip_ranges_NA_East, Ip_ranges_oc)


def playEU_server():
    unblockALL()
    blockingLabel.config(text='PLAYING ON EU', fg='#26ef4c')
    commands = 'advfirewall firewall add rule name="@EU_OW_SERVER_BLOCKER" Dir=Out Action=Allow RemoteIP=' + Ip_ranges_EU1
    shell.ShellExecuteEx(lpVerb='runas', lpFile='netsh.exe', lpParameters=commands)

    ruleMakerBlock(Ip_ranges_ME, Ip_ranges_NA_West1, Ip_ranges_AS_1, Ip_ranges_NA_West2
                   , Ip_ranges_AS_Japan, Ip_ranges_AS_Korea, Ip_ranges_AS_Taiwan,
                   Ip_ranges_SA, Ip_ranges_AS_Singapore1, Ip_ranges_AS_Singapore2, Ip_ranges_NA_central, Ip_ranges_oc,
                   Ip_ranges_NA_East)


def unblockALL():
    blockingLabel.config(text='ALL UNBLOCKED (DEFAULT SETTINGS)', fg='#ddee4a')
    list_rule_names = ["@NAEAST_OW_SERVER_BLOCKER", "@EU_OW_SERVER_BLOCKER", "@ME_OW_SERVER_BLOCKER",
                       "@NAWEST1_OW_SERVER_BLOCKER", "@AU_OW_SERVER_BLOCKER", "@NAWEST2_OW_SERVER_BLOCKER"
        , "@Overwatch Block", "@NAWEST_OW_SERVER_BLOCKER"]
    ruleDelete(list_rule_names)


def donationPage():
    webbrowser.open("https://paypal.me/vantverx?country.x=SA&locale.x=en_US")


# Labels
blockingLabel = Label(app, text='', bg='#282828', fg='#ddee4a', font=futrabook_font)
blockingLabel.grid(row=0, column=0)
blockingLabel.place(x=250, y=440, anchor="center")

# pingLabel = Label(app, text='', bg='#282828', fg='#ddee4a', font=futrabook_font)
# pingLabel.grid(row=0, column=0)
# pingLabel.place(x=250, y=430, anchor="center")

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

# Start Program
iconMaker()
checkIfActive()
app.mainloop()
