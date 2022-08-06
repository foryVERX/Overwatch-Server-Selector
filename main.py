from tkinter import *
from tkinter.font import Font
import win32com.shell.shell as shell
from subprocess import Popen, CREATE_NEW_CONSOLE, PIPE, STARTUPINFO, STARTF_USESHOWWINDOW, SW_HIDE
from PIL import ImageTk, Image
from io import BytesIO
import pic2str
import base64
from os.path import exists

# Create window object
app = Tk()
# Set Properties
app.title('MINA Overwatch Server Selector')
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
byte_UNBLOCK_ALL_MAIN = base64.b64decode(pic2str.UNBLOCK_ALL_MAIN)

image_data_SMALL_APPLICATION = BytesIO(byte_LOGO_SMALL_APPLICATION)
image_data_SQUARE_BACKGROUND_MINA_TEST = BytesIO(byte_SQUARE_BACKGROUND_MINA_TEST)
image_play_on_eu = BytesIO(byte_play_on_eu)
image_play_on_na_east = BytesIO(byte_play_on_na_east)
image_play_on_na_west = BytesIO(byte_play_on_na_west)
image_BLOCK_MIDDLE_EAST = BytesIO(byte_BLOCK_MIDDLE_EAST)
image_UNBLOCK_ALL_MAIN = BytesIO(byte_UNBLOCK_ALL_MAIN)

# Add images
background = ImageTk.PhotoImage(Image.open(image_data_SQUARE_BACKGROUND_MINA_TEST))
logo = Label(frame, image=background)
logo.pack()
button_img_ME = ImageTk.PhotoImage(Image.open(image_BLOCK_MIDDLE_EAST))
button_img_EU = ImageTk.PhotoImage(Image.open(image_play_on_eu))
button_img_NA_WEST = ImageTk.PhotoImage(Image.open(image_play_on_na_west))
button_img_ME_EAST = ImageTk.PhotoImage(Image.open(image_play_on_na_east))
button_img_Default = ImageTk.PhotoImage(Image.open(image_UNBLOCK_ALL_MAIN))

# Add font
futrabook_font = Font(family="Futura PT Demi", size=10)

# Ip ranges
Ip_ranges_ME = "157.175.0.0-157.175.255.255,15.185.0.0-15.185.255.255,15.184.0.0-15.184.255.255"
Ip_ranges_EU = "5.42.184.0-5.42.191.255,5.42.168.0-5.42.175.255"
Ip_ranges_NA_East = "24.105.40.0-24.105.47.255"
Ip_ranges_NA_West = "24.105.8.0-24.105.15.255"
Ip_ranges_AS = "104.198.0.0-104.198.255.255,34.84.0.0-34.84.255.255,34.85.0.0-34.85.255.255,35.200.0.0-35.200.255.255,35.221.0.0-35.221.255.255,34.146.0.0-34.146.255.255,117.52.0.0-117.52.255.255,121.254.0.0-121.254.255.255,5.42.0.0-5.42.255.255,34.87.0.0-34.87.255.255,34.126.0.0-34.126.255.255,35.187.0.0-35.187.255.255,37.244.42.0-37.244.42.255,34.142.0.0-34.143.255.255"


# Functions
def iconMaker():  # Used to check if there is an icon in the same directory or not it will create the icon if not.
    if exists("LOGO_SMALL_APPLICATION.ico"):
        app.iconbitmap("LOGO_SMALL_APPLICATION.ico")
    else:
        icon = Image.open(image_data_SMALL_APPLICATION)
        icon.save("LOGO_SMALL_APPLICATION.ico")
        app.iconbitmap("LOGO_SMALL_APPLICATION.ico")


def ruleMakerBlock(*argv):  # Used to block IP range
    ip_range = ""
    for arg in argv:
        if len(argv) > 1:
            ip_range += arg + ","
        else:
            ip_range = arg
    commands = 'advfirewall firewall add rule name="@Overwatch Block" Dir=In Action=Block RemoteIP=' + ip_range
    shell.ShellExecuteEx(lpVerb='runas', lpFile='netsh.exe', lpParameters=commands)
    commands = 'advfirewall firewall add rule name="@Overwatch Block" Dir=Out Action=Block RemoteIP=' + ip_range
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
    servers_active_rule_list = ['"@ME_OW_SERVER_BLOCKER"', '"@NAEAST_OW_SERVER_BLOCKER"', '"@NAWEST_OW_SERVER_BLOCKER"',
                                '"@EU_OW_SERVER_BLOCKER"']
    for rule in servers_active_rule_list:
        command = 'netsh advfirewall firewall show rule name=' + rule
        proc = Popen(command, creationflags=CREATE_NEW_CONSOLE, stdout=PIPE)
        output = proc.communicate()[0]
        output = str(output.strip().decode("utf-8"))
        temp_rule = rule.replace('"', "")
        if temp_rule in output:
            filtered = output.rpartition('_')[0].replace(" ", "").replace("RuleName:@", "").replace("_OW_SERVER", "")
            if filtered == 'ME':
                blockingLabel.config(text='ME BLOCKED', bg='#282828', fg='#ef2626', font=futrabook_font)
                break
            else:
                filtered = filtered[0:2] + ' ' + filtered[2:]
                label_text = 'YOUR PLAYING ON ' + filtered
                blockingLabel.config(text=label_text, bg='#282828', fg='#26ef4c',
                                     font=futrabook_font)
                break
        blockingLabel.config(text='ALL UNBLOCKED (DEFAULT SETTINGS)', fg='#ddee4a')


def blockMEServer():  # It removes any rules added by blockserver function
    unblockALL()
    blockingLabel.config(text='ME BLOCKED', fg='#ef2626')
    commands = 'advfirewall firewall add rule name="@ME_OW_SERVER_BLOCKER" Dir=Out Action=Allow RemoteIP=' + Ip_ranges_EU
    shell.ShellExecuteEx(lpVerb='runas', lpFile='netsh.exe', lpParameters=commands)

    # Block ME, NA West, AS, NA East
    ruleMakerBlock(Ip_ranges_ME)


def playNAEast_server():
    unblockALL()
    blockingLabel.config(text='YOUR PLAYING ON NA EAST', fg='#26ef4c')
    commands = 'advfirewall firewall add rule name="@NAEAST_OW_SERVER_BLOCKER" Dir=Out Action=Allow RemoteIP=' + Ip_ranges_NA_East
    shell.ShellExecuteEx(lpVerb='runas', lpFile='netsh.exe', lpParameters=commands)

    # Block ME, EU, NA West, AS
    ruleMakerBlock(Ip_ranges_ME, Ip_ranges_AS, Ip_ranges_NA_West, Ip_ranges_EU)


def playNAWest_server():
    unblockALL()
    blockingLabel.config(text='YOUR PLAYING ON NA WEST', fg='#26ef4c')
    commands = 'advfirewall firewall add rule name="@NAWEST_OW_SERVER_BLOCKER" Dir=Out Action=Allow RemoteIP=' + Ip_ranges_NA_West
    shell.ShellExecuteEx(lpVerb='runas', lpFile='netsh.exe', lpParameters=commands)

    # Block ME, EU, AS, NA East
    ruleMakerBlock(Ip_ranges_ME, Ip_ranges_AS, Ip_ranges_NA_East, Ip_ranges_EU)


def playEU_server():
    unblockALL()
    blockingLabel.config(text='YOUR PLAYING ON EU', fg='#26ef4c')
    commands = 'advfirewall firewall add rule name="@EU_OW_SERVER_BLOCKER" Dir=Out Action=Allow RemoteIP=' + Ip_ranges_EU
    shell.ShellExecuteEx(lpVerb='runas', lpFile='netsh.exe', lpParameters=commands)

    # Block ME, NA West, AS, NA East
    ruleMakerBlock(Ip_ranges_ME, Ip_ranges_NA_West, Ip_ranges_NA_East, Ip_ranges_AS)


def unblockALL():
    blockingLabel.config(text='ALL UNBLOCKED (DEFAULT SETTINGS)', fg='#ddee4a')
    list_rule_names = ["@NAEAST_OW_SERVER_BLOCKER", "@EU_OW_SERVER_BLOCKER", "@ME_OW_SERVER_BLOCKER",
                       "@NAWEST_OW_SERVER_BLOCKER", "@Overwatch Block"]
    ruleDelete(list_rule_names)


# Labels
blockingLabel = Label(app, text='', bg='#282828', fg='#ddee4a', font=futrabook_font)
blockingLabel.grid(row=0, column=0)
blockingLabel.place(x=250, y=410, anchor="center")

# Buttons
PlayMEButton = Button(app, image=button_img_ME, font=futrabook_font, command=blockMEServer,
                      bg='#282828', fg='#282828', borderwidth=0, activebackground='#282828')
PlayMEButton.place(x=135, y=90, height=40, width=230)

PlayEUButton = Button(app, image=button_img_EU, font=futrabook_font, command=playEU_server,
                      bg='#282828', fg='#282828', borderwidth=0, activebackground='#282828')
PlayEUButton.place(x=135, y=150, height=40, width=230)

PlayNAWESTButton = Button(app, image=button_img_NA_WEST, font=futrabook_font, command=playNAWest_server,
                          bg='#282828', fg='#282828', borderwidth=0, activebackground='#282828')
PlayNAWESTButton.place(x=135, y=210, height=40, width=230)

PlayNAEASTButton = Button(app, image=button_img_ME_EAST, font=futrabook_font, command=playNAEast_server,
                          bg='#282828', fg='#282828', borderwidth=0, activebackground='#282828')
PlayNAEASTButton.place(x=135, y=270, height=40, width=230)

ClearBlocksButton = Button(app, image=button_img_Default, font=futrabook_font, command=unblockALL,
                           bg='#282828', fg='#282828', borderwidth=0, activebackground='#282828')
ClearBlocksButton.place(x=135, y=330, height=40, width=230)

# Start Program
iconMaker()
checkIfActive()
app.mainloop()
