from tkinter import font
import tkinter as tk

import pandas as pd
from openpyxl import load_workbook

import ctypes
import os
import string
import win32api
import subprocess
from subprocess import Popen, PIPE, STDOUT

#Obtains all open windows to check for potential file transfers
#Output - List of strings of open window names
transfer_keywords = ('preparing', 'copying', 'remaining')
def get_open_windows():
    EnumWindows = ctypes.windll.user32.EnumWindows
    EnumWindowsProc = ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.POINTER(ctypes.c_int), ctypes.POINTER(ctypes.c_int))
    GetWindowText = ctypes.windll.user32.GetWindowTextW
    GetWindowTextLength = ctypes.windll.user32.GetWindowTextLengthW
    IsWindowVisible = ctypes.windll.user32.IsWindowVisible
     
    titles = []
    def foreach_window(hwnd, lParam):
        if IsWindowVisible(hwnd):
            length = GetWindowTextLength(hwnd)
            buff = ctypes.create_unicode_buffer(length + 1)
            GetWindowText(hwnd, buff, length + 1)
            titles.append(buff.value)
        return True
    EnumWindows(EnumWindowsProc(foreach_window), 0)
    return titles


#Record all the drives that are connected, put them on arrays, and find a drive letter that is not taken.
def check_mapped_drives():    
    readout = subprocess.Popen('net use', shell=True, stdin=PIPE, stdout=PIPE, stderr=STDOUT)
    output = readout.stdout.read()
    output = output.decode('utf-8').split('\n')

    mapped_drives = {}
    mapped_labname = {}
    
    # Adding in mapped network and names
    for line in output:
        if ':' in line:
            pos_colon = str.find(line, ':')
            letter = line[pos_colon-1]
            
            line_split = line.split('\\')
            drive = line_split[2]
            labname = line_split[3].split()[0]
            
            mapped_labname[letter] = labname
            mapped_drives[letter] = drive

    # Adding in local
    local_drives = win32api.GetLogicalDriveStrings()
    local_drives = local_drives.split(':\\\000')[:-1]
    for let in local_drives:
    	if let not in mapped_drives:
    		mapped_labname[let] = 'Local'
    		mapped_drives[let] = let

    mapped_inv = {v: k for k, v, in mapped_drives.items()}
    return mapped_drives, mapped_inv, mapped_labname

mapped_drives, mapped_inv, mapped_labname = check_mapped_drives()

# Obtaining unused letter to map starting from Z
alph = list(string.ascii_uppercase)
to_map = ''
while not to_map:
    letter_try = alph.pop()
    if letter_try not in mapped_drives:
        to_map = letter_try


#Read the excel file in M drive and create the GUI
os.chdir('C:/Users/EFeng/Desktop/')
path = 'Addresses.xlsx'

book = load_workbook(path)
writer = pd.ExcelWriter(path, engine='openpyxl')
writer.book = book
writer.sheets = dict((ws.title, ws) for ws in book.worksheets)

df = pd.read_excel(path, sheet_name='Sheet1')
lab_names = df['Lab NameA'].tolist() + df['Lab NameL'].tolist()
lab_address = df['Lab AddressA'].tolist() + df['Lab AddressL'].tolist()

lab_names = [x for x in lab_names if str(x) != 'nan']
lab_address = [x for x in lab_address if str(x) != 'nan']

labs = dict(zip(lab_names, lab_address))
labs_alph = {}

for lab in lab_names:
    letter = lab[0]
    if letter not in labs_alph:
        labs_alph[letter] = [lab]
    else:
        labs_alph[letter].append(lab)

limit = 11
recents_df = pd.read_excel(path, sheet_name='Sheet2')
recents = recents_df['Recent'].tolist()
counts = recents_df['Counter'].tolist()

#Setting up the GUI
row_msg = 0
row_button = 4
row_address = 8
row_user = 10
row_pass = 11
row_connect = 13
row_space = [7, 9, 12, 14]

root = tk.Tk()
root.title("Connect to your labshare")

for i in row_space:
    space = tk.Message(root, text="")
    space.grid(row=i)

arial12=font.Font(family='Arial', size=12)
arial10=font.Font(family='Arial', size=10)

label_message = """Please choose your lab OR directly type the lab share address below. 
                 \n\nIf you do not see your lab in the list, please contact staff."""
label=tk.Message(root,width=500,font=arial12,padx=20,pady=10,text=label_message)
label.grid(row=row_msg,columnspan=4)

example = tk.Label(root, padx=5, text="Address (e.g. \\\example\labname): ", font=arial10)
example.grid(row=row_address, column=0, columnspan=2, sticky="e")
address_input = tk.Entry(root, bd=2, width=30)
address_input.grid(row=row_address, column=2, columnspan=2, sticky="w")

def select_lab(lab_name):
    address_input.delete(first=0, last=100)
    address_input.insert(0, labs[lab_name])
    address_input.lab_name = lab_name

r = row_button
c = 0
for i in range(limit):
    lab_e = recents[i]
    button = tk.Button(root, text=lab_e, command= lambda lab_e=lab_e: select_lab(lab_e))
    button.config(width=18, font=('Arial', 10, 'bold'))
    button.grid(row=r, column=c, padx=5, pady=5)
    c += 1
    if c > 3:
        c = 0
        r += 1

other = tk.StringVar(root)
other.set('OTHER')
other.trace('w', lambda *args: select_lab(other.get()))   
     
other_button = tk.Menubutton(root, textvariable=other, relief='raised', indicatoron=True,
                             borderwidth=1)
other_button.config(width=16, font=('Arial', 10, 'bold'), fg='red')
other_button.grid(row=r, column=c)

other_mainmenu = tk.Menu(other_button, tearoff=False)
other_button.configure(menu=other_mainmenu)

for letter in labs_alph:
    menu = tk.Menu(other_mainmenu, tearoff=False)
    other_mainmenu.add_cascade(label=letter, menu=menu)    
    for lab in labs_alph[letter]:
        menu.add_radiobutton(value=lab, label=lab, variable=other)
           
tk.Label(root, text="Username: ",font=arial10).grid(row=row_user, column=1, sticky="e")
tk.Label(root, text="Password: ",font=arial10).grid(row=row_pass, column=1, sticky="e")
username=tk.Entry(root, bd=2, width=30)
password=tk.Entry(root, bd=2, width=30)
password.config(show="*")
username.grid(row=row_user, column=2, columnspan=2, sticky="w")
password.grid(row=row_pass, column=2, columnspan=2, sticky="w")

def closeprogram():
    root.destroy()
    
def connect():
    print('Connecting...')
    mapped_drives, mapped_inv, mapped_labname = check_mapped_drives()
    
    error_msg = tk.Message(root, width=300, fg='red', font=arial12, text='')
    error_msg.grid(row=row_connect, column=2, columnspan=2, sticky='w')
    
    address = address_input.get()
    user = 'example\\' + username.get()
    pswd = password.get()
    
    prefix_pos = address[2:].find('\\') + 2
    prefix = address[2:prefix_pos]
    
    double = False
    for drive_name in mapped_drives.values():
        if prefix == drive_name:
            double = True
            curr_mapped_drive = drive_name

    if double:
        transferring = False
        drive_letter = mapped_inv[curr_mapped_drive]
        open_windows = get_open_windows()
        for title in open_windows:
            for keyword in transfer_keywords:
                if keyword in title.lower():
                    transferring = True
                    
        if transferring:
            error_text = 'Someones drive already mapped and transferring.' + \
                         ' Please try another computer'
            error_msg.configure(text=error_text)
            return
        else:
            os.system(r"net use "+drive_letter+": /delete /y")

    subprocess.call(r"net use "+to_map+": "+address+" "+pswd+" /USER:"+user+" /persistent:no")
    if os.path.isdir(''+to_map+':\\'):
        dummy_user_folder = os.listdir(to_map+':\\')[0]
        user_folder_path = to_map+':\\'+dummy_user_folder
        
        subprocess.Popen(r'explorer /select, "C:\Users\Desktop\Data"')
        subprocess.Popen(r'explorer /select, '+user_folder_path)  
        update_recents(address_input.lab_name)
        root.destroy() 
    else:
        error_text = "Username or Password is wrong, please try again"   
        error_msg.configure(text=error_text)
    return        
        
def onclick(event):
    connect()
    
def update_recents(lab_name):
    pos = recents.index(lab_name)
    counts[pos] += 1

    recents.insert(0, recents.pop(pos))
    counts.insert(0, counts.pop(pos))
    
    recents_df = pd.DataFrame(recents)
    counter_df = pd.DataFrame(counts)
    recents_df.to_excel(writer, sheet_name='Sheet2', na_rep='', float_format=None, columns=None,
                        header=False, index=False, index_label=None, startrow=1)
    counter_df.to_excel(writer, sheet_name='Sheet2', na_rep='', float_format=None, columns=None,
                    header=False, index=False, index_label=None, startrow=1, startcol=1)
    writer.save()
    writer.close()
            
root.bind('<Return>', onclick)

labchosen=tk.Button(root, bg="salmon", text="Connect to my lab",font=arial12, command=connect)
labchosen.grid(row=row_connect, columnspan=2)


root.mainloop()
