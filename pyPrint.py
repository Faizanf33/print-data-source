### Imports go here
import os
import time
import logging

from tkinter import *
from tkinter import ttk, filedialog
import tkinter.messagebox as tkmb
import tkinter.simpledialog as tksd

## Imports from files
from Data.send_mail import Mail
from Data.set_data import Delegate
from Data.tool_tip import CreateToolTip
# from AnimatedGIF import start_gif

## To run the program explicitly
try:
    import ctypes
    myappid = 'mycompany.myproduct.subproduct.version' # arbitrary string
    ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)
except:
    pass

## Setting up debugging
# Level set to logging.Debug
try:
    Path = os.path.join('tmp', 'info.log')
    logging.basicConfig(format="[NYMUN Forms]:[%(asctime)s]:%(levelname)s:%(message)s", datefmt='%m/%d/%Y %I:%M:%S %p', level=logging.DEBUG, filename=Path)

except (IOError or FileNotFoundError):
    os.mkdir('tmp')
    Path = os.path.join('tmp', 'info.log')
    logging.basicConfig(format="[NYMUN Forms]:[%(asctime)s]:%(levelname)s:%(message)s", datefmt='%m/%d/%Y %I:%M:%S %p', level=logging.DEBUG, filename=Path)

users_info = {}
admin = ""
status = False

def start_server():
    window.wm_withdraw()
    id = tksd.askstring("Credentials", "EMAIL-ID :", parent=window)
    if id:
        import re
        if not re.match(r"^[A-Za-z0-9\.\+_-]+@[A-Za-z0-9\._-]+\.[a-zA-Z]*$", id):
            terminal.insert(END, ">> Invalid EMAIL_ID Entered!")
            tkmb.showwarning("INVALID ID", "EMAIL_ID IS INVALID!\nTRY AGAIN...")
            window.deiconify()
            return

        else:
            passwrd = tksd.askstring("Credentials", "PASSWORD :", show="\u2022", parent=window)
            if passwrd:
                import threading
                window.update()
                terminal.insert(END, ">> Credentials received!")
                global status
                def start_server_thread():
                    global admin
                    admin = Mail(id, passwrd)

                # def start_gif_thread():
                #     new_path = os.getcwd()
                #     os.chdir(initial_path)
                #     start_gif()

                server_thread = threading.Thread(target=start_server_thread)
                server_thread.start()

                # gif_thread = threading.Thread(target=start_gif_thread)
                # gif_thread.start()
                # os.chdir(new_path)

                window.config(cursor="wait")
                terminal.insert(END, ">> Connecting, please wait...")
                window.update()

                if server_thread.is_alive():
                    window.deiconify()
                    window.update()
                    time.sleep(8)

                else:
                    time.sleep(1)

                try:
                    status = admin.server_status
                    if status == True:
                        window.config(cursor="")
                        window.update()
                        terminal.insert(END, ">> Logged in successfully")
                        terminal.insert(END, ">> Using Id {} at server {}".format(admin.user_id, admin.server))
                        b1.configure(text="RESTART SERVER", command=start_server)
                        b1_ttp = CreateToolTip(b1, "Restart mail server")

                    elif admin.start_service() == True:
                        window.config(cursor="")
                        window.update()
                        del server_thread
                        mainEntry.focus()
                        logging.warning("EMAIL_ID/PASSWORD INCORRECT!")
                        terminal.insert(END, ">> EMAIL_ID/PASSWORD INCORRECT")
                        return False

                    else:
                        raise TimeoutError

                except Exception:
                    window.config(cursor="")
                    window.update()
                    del server_thread
                    mainEntry.focus()
                    terminal.insert(END, ">> ____CONNECTION-TIMEDOUT_____")
                    return False
            else:
                window.update()
                window.deiconify()
                terminal.insert(END, ">> Credentials were not provided!")

    else:
        window.update()
        window.deiconify()
        terminal.insert(END, ">> Credentials were not provided!")

    mainEntry.focus()
    return False

def openfile():
    file = filedialog.askopenfilename(filetypes = (("Delimited Text File", "*.csv"),("All Files","*.*")))
    if file.endswith("NYMUN Multan.csv"):
        return file
    else:
        return ""

def set_data(event=None):
    file = openfile()
    window.config(cursor="wait")
    window.update()
    try:
        new = Delegate(file)
        window.config(cursor="")
        global users_info
        users_info = new.info
        terminal.insert(END, ">> File found => '{}'".format(file[-16:]))
        terminal.insert(END, ">> {0} Files written in directory -> {1}".format(new.files_created, new.dir))
        terminal.insert(END, ">> On Your Desktop, Look For '{}' Directory".format(new.dir))
        terminal.insert(END, ">> Process completed with {} errors".format(new.errors))
        b1.configure(text="START SERVER", command=start_server)
        b1_ttp = CreateToolTip(b1, "Start mail server")
        reply =  tkmb.askquestion("EMAIL", "START MAIL SERVER?")
        if reply == 'yes':
            terminal.insert(END, "\n>> Mail request is accepted!")
            start_server()
            if status == False:
                return

        else:
            window.update()
            terminal.insert(END, ">> Request for mail was rejected!")

        mainEntry.focus()
        return

    except FileNotFoundError:
        window.config(cursor="")
        logging.error("File not found => '{}'".format(file))
        logging.info("No further proceedings!")
        terminal.insert(END, ">> File -> 'NYMUN Multan.csv' Not Found!")
        terminal.insert(END, ">> No further proceedings!")
        mainEntry.focus()
        return

def find_serial(event=None):
    if users_info == {}:
        tkmb.showwarning("DATA UN-ARRANGED", "Please arrange your data first!")
        return

    prov_info = tksd.askstring("FIND", "EMAIL_ID/PHONE NO.", parent=frame)
    if prov_info:
        terminal.insert(END, ">> Searching record for : {}".format(prov_info))
        if prov_info.isdigit():
            prov_info = prov_info[1:]

        for serial in users_info.keys():
            if prov_info in users_info[serial]:
                tkmb.showinfo("PASSED", "RECORD FOUND!\nSERIAL NO. : {}".format(serial))
                terminal.insert(END, ">> Record matched with serial no {}".format(serial))
                mainEntry.delete(0, END)
                mainEntry.insert(0, serial)
                mainEntry.select_range(0, END)
                mainEntry.focus()
                return

            else:
                pass

        terminal.insert(END, ">> No match found!")
        tkmb.showinfo("FAILED", "NO RECORD FOUND!")
        return

    else:
        terminal.insert(END, ">> Search request cancelled!")
        return

def send_mail(serial_no):
    window.config(cursor="wait")
    window.update()
    if serial_no.upper() in users_info.keys():
        address = users_info[serial_no][7]
        name = users_info[serial_no][3]
        no_of_delegate = users_info[serial_no][2]
        if no_of_delegate == "Individual":
            amount = 3500

        else:
            amount = 2500 + (int(no_of_delegate) * 2500)

        default = "Dear {0}! \n\nWe have received a payment of Rs. {1:,}/- for NYMUN Registration. Thank you for being a part of NYMUN Multan 2018. \nFor more information and updates, visit us at 'https://www.facebook.com/NYMUNpakistan/'. \n\nIn case of any query, feel free to contact the undersigned. \n\nMr. Faraz Sheikh ( +92 306 8365633 )\n\n\n\nRegards\nRegistraton Team NYMUN"
        msg = default.format(name, amount)
        terminal.insert(END, ">> Sending mail at {}".format(address))
        try:
            admin.sendmail(address, msg)
            window.config(cursor="")
            terminal.insert(END, ">> Mail sent successfully.")
        except Exception:
            window.config(cursor="")
            global status
            status = False
            terminal.insert(END, ">> Unable to send mail, try restarting server!")
    else:
        terminal.insert(END, ">> Unable to send mail! serial_no {} not found".format(serial_no))

    return


def print_file(event=None):
    window.config(cursor="wait")
    window.update()
    path = os.getcwd()
    serial_no = mainEntry.get()
    serial_no = serial_no.upper()

    if path == initial_path:
        window.config(cursor="")
        window.update()
        mainEntry.delete(0, END)
        b1.focus()
        tkmb.showwarning("DATA UN-ARRANGED", "Please arrange your data first!")
        return

    if serial_no == "":
        window.config(cursor="")
        window.update()
        tkmb.showwarning("NO SERIAL NO.", "PLEASE ENTER A SERIAL NO. FIRST!")
        return

    elif not serial_no.startswith("S#"):
        window.config(cursor="")
        window.update()
        tkmb.showwarning("INVALID SERIAL NO.", "PLEASE ENTER A VALID SERIAL NO.\ne.g. s#0002 or S#0010 ...")
        return


    else:
        try:
            import win32api, win32print
        except:
            window.config(cursor="")
            window.update()
            pass

        my_file = "{}.pdf".format(serial_no)
        directories = ["Individual", 3, 4, 5, 6]
        mainEntry.delete(0, END)
        mainEntry.focus()
        terminal.insert(END, ">> Searching for file with Serial No. {}".format(serial_no))
        for root, dirs, files in os.walk(path):
            if my_file in files:
                path_to_doc = os.path.join(root, my_file)
                terminal.insert(END, ">> File found at {}".format(path_to_doc))
                try:
                    win32api.ShellExecute (0, "print", path_to_doc, None, '/d:"{}"'.format(win32print.GetDefaultPrinter()), 0)
                    window.config(cursor="")
                    window.update()
                    logging.info("File {} sent for printing".format(my_file))
                    terminal.insert(END, ">> File {} sent for printing".format(my_file))
                    os.chdir(path)
                    with open("PrintRecords.txt", "a") as inFile:
                        text = "File '{0}' , Print Time : '{1}' , Printed by : '{2}'\n".format(my_file, time.strftime('%c'), admin_name)
                        inFile.write(text)

                    if status == True:
                        reply =  tkmb.askquestion("EMAIL", "SEND CONFIRMATION MAIL?")
                        if reply == "yes":
                            send_mail(serial_no)

                        else:
                            pass

                except Exception as exc:
                    window.config(cursor="")
                    window.update()
                    logging.warning("While printing file {} error was raised => {}".format(my_file, exc))
                    terminal.insert(END, ">> Unable to print file {}".format(my_file))
                    terminal.insert(END, ">> Please make sure your system has 'ADOBE READER' installed!")

                return

            else:
                window.config(cursor="")
                window.update()
                pass
            os.chdir("..")

        window.config(cursor="")
        window.update()
        terminal.insert(END, ">> No such file with Serial No. {}".format(serial_no))
        terminal.insert(END, ">> Recheck serial no. and try again..")
        os.chdir(path)
        return

def exit(event=None):
    path = os.getcwd()

    if path == initial_path:
        logging.info('exiting window')
        window.destroy()
        window.quit()

    else:
        window.withdraw()
        if status == True:
            msg = tkmb.askquestion("CONFIRM","ARE YOU SURE YOU WANT TO EXIT?\n\tServer still running...", icon='warning')
        else:
            msg = tkmb.askquestion("CONFIRM","ARE YOU SURE YOU WANT TO EXIT?\n\tAll progress will be lost!", icon='warning')

        if msg == "yes":
            logging.info('exiting window')
            window.destroy()
            window.quit()

        else:
            logging.info('window still running')
            window.deiconify()
            return

def help_me():
    window.iconify()
    tkmb.showinfo("Help", "1. Open file from your computer. \n2. Filename must exactly be as 'NYMUN Multan.csv'. \n3. Start server in order to send mail(s). \n4. Type correct & complete email-id/password. \n-> Try to use NYMUN Official Id. \n5. DO NOT start server if your Internet Connection is slow. \n6. Search using email-id/phone-no. \n-> (Hint) You may also search using name(s). \n7. Type serial number as provided(without e.g.). \n8. Print records may be found in 'NYMUN FORMS' dir(placed at desktop). \n\nIn case of any issue/query email at 'faizanahmad33.fa@gmail.com' \n-> Do attach 'info.log' file(placed in 'tmp' dir) in mail.")
    window.deiconify()

###################################################################################################
#                                           Creating windows                                      #
###################################################################################################
window = Tk()
window.title("NYMUN FORMS")

### Setting up windows geometry
w, h = 360, 360
## to open window in the centre of screen
ws = window.winfo_screenwidth()
hs = window.winfo_screenheight()
x_axis = (ws/2) - (w/2)
y_axis = (hs/2) - (h/2)

window.geometry('%dx%d+%d+%d' % (w, h, x_axis, y_axis))
window.resizable(0,0)

### remember initial_path
initial_path = os.getcwd()

# if windows default cross button is pressed
window.protocol('WM_DELETE_WINDOW',exit)

# Add transparency
window.attributes("-alpha", 0.9)

### Creating Frames
frame = Frame(window, bg="LightCyan2", bd=4, colormap="new", height=h)
frame.pack(fill=BOTH, side=TOP, padx=0, pady=0, ipadx=0, ipady=0)

### Creating Frame2
frame2 = Frame(window, bg="LightCyan2", bd=4, colormap="new", height=h)
frame2.pack(fill=BOTH, side=TOP, padx=0, pady=0, ipadx=0, ipady=0)

### Creating logo button
logo_button = Button(frame, width = 44, height = 44, command=help_me, relief=RIDGE, bg="LightCyan2", bd=3)
logo_button.place(x=0, y=-5)
img = PhotoImage(file="img\\nymun.png")
logo_button.config(image=img)
logo_button.bind("<Return>", help_me)
b1_ttp = CreateToolTip(logo_button, "Click for help")

### Creating Button
ttk.Style().configure("TButton", padding=4, borderwidth=8, relief=RAISED, background="cyan2", font=('times', 12, 'bold'))
b1 = ttk.Button(frame, text="CHOOSE FILE", command=set_data)
b1.pack(side=TOP, padx=6, ipadx=0, pady=6, ipady=0)
b1.bind("<Return>", set_data)
b1_ttp = CreateToolTip(b1, "Choose .csv file")

### Creating day
day = Label(frame, text=time.strftime('%A'), font=('times', 12), bg="LightCyan2", fg='grey1')
day.place(x=270, y=-1)

### Creating date
date = Label(frame, text=time.strftime('%d-%b-%y'), font=('times', 12), bg="LightCyan2", fg='grey1')
date.place(x=270, y=20)

## Creating other Buttons
searchButton = Button(frame2, command=find_serial, relief=GROOVE, text="SEARCH", width=8, height=0, font=('times', 12), bg="azure", fg="black", bd=3, activebackground="pale turquoise", activeforeground="black")
searchButton.place(x=5, y=56)
searchButton.bind("<Return>", find_serial)
prB_ttp = CreateToolTip(searchButton, "Search serial no.")


printButton = Button(frame2, command=print_file, relief=GROOVE, text="PRINT", width=8, height=0, font=('times', 12, 'bold'), bg="azure", fg="black", bd=3, activebackground="pale turquoise", activeforeground="black")
printButton.place(x=165, y=56)
printButton.bind("<Return>", print_file)
prB_ttp = CreateToolTip(printButton, "Print file")

exitButton = Button(frame2, command=exit, relief=GROOVE, text="EXIT", width=8, height=1, font=('times', 12, 'bold'), bg="brown3", fg="white", bd=3, activebackground="brown4", activeforeground="white")
exitButton.pack(side=BOTTOM, anchor=E, padx=6, ipadx=0, pady=10, ipady=0)
exitButton.bind("<Return>", exit)
exit_ttp = CreateToolTip(exitButton, "Close")

### Creating message box
scroll = Scrollbar(frame, bd=4, relief=RIDGE)
scroll.pack(side=RIGHT, anchor=E, fill=Y)
terminal = Listbox(frame, yscrollcommand = scroll.set, bg="grey35", fg="white", font=('times', 12), height=8, width=41, bd=5, relief=SUNKEN)
terminal.pack(fil=BOTH, side=LEFT, anchor=W, pady=0, ipadx=0, padx=0, ipady=0)
scroll.config(command=terminal.yview)

### Creating Labels
mainLabel = Label(frame2, text="Enter Serial No. :\t\t", font="times 12 bold", bg="LightCyan3", relief=RIDGE, bd=2)
mainLabel.pack(side=LEFT, anchor=SW, pady=8, ipadx=0, ipady=0, fil=BOTH)

### Creating entry
mainEntry = Entry(frame2, font="times 12", bd=4, bg="white", fg="black", relief=RIDGE)
mainEntry.insert(0, "e.g. s#0000")
mainEntry.focus()
mainEntry.select_range(0, END)
mainEntry.pack(side=RIGHT, anchor=E, pady=8, ipadx=0, ipady=0)
mE_ttp = CreateToolTip(mainEntry, "Type here, e.g. S#0001")
mainEntry.bind("<Return>", print_file)

### Copyright label
cplabel = Label(window, text="NYMUN FORMS {} 2018".format(chr(0xa9)), font="chiller 12 bold", bg="LightCyan4", fg="grey2", relief=SUNKEN, bd=2)
cplabel.pack(side=BOTTOM, anchor=SW, fill=X, padx=0, pady=0, ipadx=0, ipady=0)

if __name__ == "__main__":
    time.sleep(1)
    window.withdraw()
    window.iconbitmap("img\\nymun.ico")
    window.update()
    ### Admin name
    admin_name = tksd.askstring("Greets", "NAME PLEASE...")
    if not admin_name:
        admin_name = "ADMIN"

    window.deiconify()
    window.update()
    window.mainloop()
