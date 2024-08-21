from tkinter import *
import tkinter.messagebox
from PIL import Image, ImageTk
import mysql.connector
import re
from sign_up_page import Newuser
from welcome_copo import welcome_dashboard

# pillow library
root = Tk()
root.geometry('800x400+240+100')
root.resizable(False,False)
root.title("CO PO Outcome")

# to set background image
myimage = Image.open('images/mybackgroundimage.jpg')
myimage_resize = myimage.resize((800, 400), Image.LANCZOS)
get_image = ImageTk.PhotoImage(myimage_resize)
my_bg_label = Label(root, image=get_image, bd=0)
my_bg_label.place(x=0, y=0)

# to set an app background
myimage_icon = Image.open('images/appicon.png')
myimage_icon_resize = myimage_icon.resize((70, 70), Image.LANCZOS)
get_image_icon = ImageTk.PhotoImage(myimage_icon_resize)
root.iconphoto(False, get_image_icon)

# to add text boxes
app_title = Label(root, text='CO PO Outcome', fg='white', bg='black', font='copper 38')
app_title.place(x=200, y=15)

email = Label(root, text="Email: ", fg='white', bg='black', font='copper 28')
email.place(x=30, y=100)
email_entry = Entry(root, font='copper 20')
email_entry.place(x=300, y=100, width=330)

password = Label(root, text="Password", fg='white', bg='black', font='copper 28')
password.place(x=30, y=200)
password_entry = Entry(root, font='copper 20')
password_entry.place(x=300, y=200)

def login_check():
    em = email_entry.get()
    if em.strip() != '':
        myexpression_email = '^\S+@\S+\.\S+$'
        if re.search(myexpression_email, em):
            passw = password_entry.get()
            if passw.strip() != '':
                try:
                    mydatabase = mysql.connector.connect(host='localhost', user='root', password='', database='copo')
                    mycommand = mydatabase.cursor()
                    q = "select * from proflogin where email='{}' and password='{}'".format(em, passw)
                    mycommand.execute(q)
                    mydata = mycommand.fetchall()
                    n= len(mydata)
                    if n==0:
                        tkinter.messagebox.showwarning("Unauthorised User", 'Invalid Login Details')
                    else:
                        for value in mydata:
                            Name = value[1]
                            Email = value[2]
                            root.withdraw()
                            welcome_dashboard(Name, Email, root)

                except Exception as ex:
                    tkinter.messagebox.showwarning("Error", ex)
            else:
                tkinter.messagebox.showwarning("Error", "Please Enter Your Password")
        else:
            tkinter.messagebox.showwarning("Error", "Please Enter Valid Email-Id")
    else:
        tkinter.messagebox.showwarning("Error", "Please Enter Your Email-Id")


login = Button(root, text='Login', fg='white', bg='black', font='copper 20', command=login_check)
login.place(x=300, y=300)
new_user = Button(root, text='New User', fg='white', bg='black', font='copper 20', command=Newuser)
new_user.place(x=400, y=300)


root.mainloop()




