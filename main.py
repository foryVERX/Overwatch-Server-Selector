from tkinter import *
import win32com.shell.shell as shell
import win32con
import time
from subprocess import Popen, CREATE_NEW_CONSOLE, PIPE


# /K option tells cmd to run the command and keep the command window from closing. You may use /C instead to close the command window after the





# Create window object
app = Tk()
# Set Properties
app.title('Overwatch Middle East Server Blocker')
app.geometry('500x150')
# Ip ranges
Ip_ranges = "157.175.0.0-157.175.255.255,15.185.0.0-15.185.255.255,15.184.0.0-15.184.255.255"


# To check if server is blocked or not
def checkIfActive():
    command = 'netsh advfirewall firewall show rule name="@Overwatch Middle East Server Block"'
    proc = Popen(command, creationflags=CREATE_NEW_CONSOLE, stdout=PIPE)
    output = str(proc.communicate()[0])
    if "No rules" in output:
        unblockserver()
    else:
        blockserver()


# Functions
def blockserver():
    blockingLabel.config(text='Server Blocked ☑', fg='green')
    commands = 'advfirewall firewall delete rule name = "@Overwatch Middle East Server Block"'
    shell.ShellExecuteEx(lpVerb='runas', lpFile='netsh.exe', lpParameters=commands)
    time.sleep(0.1)
    commands = 'advfirewall firewall add rule name="@Overwatch Middle East Server Block" Dir=Out Action=Block RemoteIP=' + Ip_ranges
    shell.ShellExecuteEx(lpVerb='runas', lpFile='netsh.exe', lpParameters=commands)
    time.sleep(0.1)
    commands = 'advfirewall firewall add rule name="@Overwatch Middle East Server Block" Dir=In Action=Block RemoteIP=' + Ip_ranges
    shell.ShellExecuteEx(lpVerb='runas', lpFile='netsh.exe', lpParameters=commands)


def unblockserver():
    blockingLabel.config(text='Server unblocked ☒', fg='red')
    commands = 'advfirewall firewall delete rule name = "@Overwatch Middle East Server Block"'
    shell.ShellExecuteEx(lpVerb='runas', lpFile='netsh.exe', lpParameters=commands)


# Text
# Program Information text
info_text = StringVar()
infor_rawtext = 'This application is used to Manage Middle East Servers of Overwatch \n made by Discord: VERX#2227'
info_label = Label(app, text=infor_rawtext, font=('bold', 10), pady=20)
info_label.grid(row=0, column=0)
info_label.place(x=250, y=40, anchor="center")

blockingLabel = Label(app, text='')
blockingLabel.grid(row=0, column=0)
blockingLabel.place(x=250, y=100, anchor="center")

# Buttons
blockButton = Button(app, text='Block Middle East Server', command=blockserver)
unBlockButton = Button(app, text='Unblock Middle East Server', command=unblockserver)
blockButton.place(x=100, y=60)
unBlockButton.place(x=250, y=60)

# Start Program
checkIfActive()
app.mainloop()
