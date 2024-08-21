import re
import tkinter.messagebox
from tkinter import *
from PIL import Image, ImageTk
import  mysql.connector
def Newuser():
    newuser_root=Toplevel()
    newuser_root.geometry('500x500+320+100')
    newuser_root.resizable(False,False) #(take koi resize na kr ske window ko)
    newuser_root.title("New Account Regestration") #(to change the title of the window box)
    #to set background image
    myimage_user=Image.open('images/new_user.jpeg')
    myimage_resize_user=myimage_user.resize((500,500),Image.LANCZOS)   #(yeh image ko resize kr dega)
    get_image_user=ImageTk.PhotoImage(myimage_resize_user)
    my_bg_label_user=Label(newuser_root,image=get_image_user,bd=0)
    my_bg_label_user.place(x=0,y=0)

    #to set app icon
    myimage_icon_user=Image.open('images/appicon.png')
    myimage_resize_icon_user=myimage_icon_user.resize((70,70),Image.LANCZOS)
    get_image_icon_user=ImageTk.PhotoImage(myimage_resize_icon_user)
    newuser_root.iconphoto(False,get_image_icon_user)

    #to give a title
    newuser_title=Label(newuser_root,text="Registration Form",fg="white",bg="black",font="copper 30 bold")
    newuser_title.place(x=100,y=15)

    email_user=Label(newuser_root,text="Enter Email Id",fg="white",bg="black",font="copper 20")
    email_user.place(x=30,y=100)
    password_user=Label(newuser_root,text="Enter Password",fg="white",bg="black",font="copper 20")
    password_user.place(x=30,y=200)
#to add atext box
    email_entry_user=Entry(newuser_root,font="copper 20")
    email_entry_user.place(x=260,y=100, width=200)

    password_entry_user=Entry(newuser_root,show="*",font="copper 20")
    password_entry_user.place(x=260,y=200, width=200)

    name_user=Label(newuser_root,text="Enter Name",fg="white",bg="black",font="copper 20")
    name_user.place(x=30,y=290)

    name_entry_user=Entry(newuser_root,font="copper 20")
    name_entry_user.place(x=260,y=290, width=200)

    mobile_user=Label(newuser_root,text="Enter Mobile",fg="white",bg="black",font="copper 20")
    mobile_user.place(x=30,y=370)

    mobile_entry_user=Entry(newuser_root,font="copper 20")
    mobile_entry_user.place(x=260,y=370, width=200)

    def save_user():
        em = email_entry_user.get()
        pw = password_entry_user.get()
        nm = name_entry_user.get()
        mb = mobile_entry_user.get()
        if em.strip() != '':
            myexpression_email = '^\S+@\S+\.\S+$'
            if re.search(myexpression_email, em):
                if pw.strip() != '':
                    if nm.strip() != '':
                        if mb.strip() != '':
                            try:
                                mydatabase = mysql.connector.connect(host='localhost', user='root', password = "", database ="copo")
                                q = "insert into proflogin (name,email,password,mobile_no)values('{}','{}','{}','{}')".format(nm,em,pw,mb)

                                mycommand = mydatabase.cursor()
                                mycommand.execute(q)
                                mydatabase.commit()
                                tkinter.messagebox.showwarning("Success", "SignUp Completed")
                                newuser_root.withdraw()
                            except Exception as ex:
                                tkinter.messagebox.showwarning("Error", ex)
                        else:
                            tkinter.messagebox.showwarning("Error", "Please Enter Mobile Number")
                    else:
                        tkinter.messagebox.showwarning("Error", "Please Enter Name")
                else:
                    tkinter.messagebox.showwarning("Error", "Please Enter Password")
            else:
                tkinter.messagebox.showwarning("Error", "Please Enter valid Email_Id")
        else:
                tkinter.messagebox.showwarning("Error", "Please Enter Email_Id")

        # to create sign up button
    signup_btn = Button(newuser_root, text='Create Account', fg='white', bg='black', font='copper 20', command=save_user)
    signup_btn.place(x=150, y=430)

    newuser_root.mainloop()
