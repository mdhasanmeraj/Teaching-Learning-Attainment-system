from tkinter import *
import tkinter.messagebox
from PIL import Image, ImageTk
from sign_up_page import Newuser
import mysql.connector
import pandas as pd
from sqlalchemy import create_engine
import pymysql
import xlsxwriter
import os
from attainment_matrix import matrix
from assignment_matrix import matrix1
from creating_sco import sco


def welcome_dashboard(username, useremail, root_window):
    wel_root = Toplevel()
    wel_root.title("welcome ,"+username+", "+useremail)
    wel_root.geometry('700x500+320+100')
    wel_root.resizable(False, False)

    # to set background image
    myimage_wel = Image.open('images/dashboard.jpg')
    myimage_resize_wel = myimage_wel.resize((700, 500), Image.LANCZOS)
    get_image_wel = ImageTk.PhotoImage(myimage_resize_wel)
    my_bg_label_wel = Label(wel_root, image=get_image_wel, bd=0)
    my_bg_label_wel.place(x=0, y=0)

    # to set app icon
    myicon_wel = Image.open('images/appicon.png')
    myicon_resize_wel = myicon_wel.resize((70, 70), Image.LANCZOS)
    get_icon_wel = ImageTk.PhotoImage(myicon_resize_wel)
    wel_root.iconphoto(False, get_icon_wel)

    app_title_wel = Label(wel_root, text='Welcome To CO-PO Calculator', fg='white', bg='black', font='copper 30')
    app_title_wel.place(x=80, y=15)




    def close():
        wel_root.quit()

    close_butn = Button(wel_root, text='Log Out', fg='red', bg='black', font='copper 25', command=close)
    close_butn.place(x=500, y=100)

    def open_word_file():
        # Path to your Word document; ensure this is accessible and correct
        word_file_path = r'D:\college_project\CoPo outcomes.docx'
        try:
            os.startfile(word_file_path)
        except Exception as e:
            print(f"Failed to open file: {e}")

    # Create a button that opens the Word document
    open_file_button = Button(wel_root, text="CoPo Outcomes Mapping List", fg='yellow', bg='black', font='copper 15', command=open_word_file)
    open_file_button.pack()
    open_file_button.place(x=400, y=200)


    def mst1():
        tkinter.messagebox.showwarning("Error", "First Click On The Clear Button to clear and reset the database and then enter your data.")
        window_frame = Frame(wel_root, bg='black')
        window_frame.place(x=50, y=50, width=800, height=500)

        def close_frame():
            window_frame.destroy()

        roll_no = Label(window_frame, text="Roll No: ", fg='white', bg='black', font='copper 15')
        roll_no.place(x=50, y=40)
        roll_no_entry = Entry(window_frame, font='copper 15')
        roll_no_entry.place(x=200, y=40)

        std_name = Label(window_frame, text='Student Name: ', fg='white', bg='black', font='copper 15')
        std_name.place(x=50, y=80)
        std_name_entry = Entry(window_frame, font='copper 15')
        std_name_entry.place(x=200, y=80)

        q1 = Label(window_frame, text='Q1: ', fg='white', bg='black', font='copper 15')
        q1.place(x=50, y=120)
        q1_entry = Entry(window_frame, font='copper 15')
        q1_entry.place(x=200, y=120)

        q2 = Label(window_frame, text='Q2: ', fg='white', bg='black', font='copper 15')
        q2.place(x=50, y=160)
        q2_entry = Entry(window_frame, font='copper 15')
        q2_entry.place(x=200, y=160)

        q3 = Label(window_frame, text='Q3: ', fg='white', bg='black', font='copper 15')
        q3.place(x=50, y=200)
        q3_entry = Entry(window_frame, font='copper 15')
        q3_entry.place(x=200, y=200)

        q4 = Label(window_frame, text='Q4: ', fg='white', bg='black', font='copper 15')
        q4.place(x=50, y=240)
        q4_entry = Entry(window_frame, font='copper 15')
        q4_entry.place(x=200, y=240)

        q5 = Label(window_frame, text='Q5: ', fg='white', bg='black', font='copper 15')
        q5.place(x=50, y=280)
        q5_entry = Entry(window_frame, font='copper 15')
        q5_entry.place(x=200, y=280)

        q6 = Label(window_frame, text='Q6: ', fg='white', bg='black', font='copper 15')
        q6.place(x=50, y=320)
        q6_entry = Entry(window_frame, font='copper 15')
        q6_entry.place(x=200, y=320)

        matrix1 = Button(window_frame, text='matrix', fg='white', bg='black', font='copper 20', command=matrix)
        matrix1.place(x=500, y=300)



        def save():
            rol = roll_no_entry.get()
            nm = std_name_entry.get()
            a = q1_entry.get()
            b = q2_entry.get()
            c = q3_entry.get()
            d = q4_entry.get()
            e = q5_entry.get()
            f = q6_entry.get()
            if rol.strip() != '':
                if nm.strip() != '':
                    try:
                        mydatabase = mysql.connector.connect(host='localhost', user='root', password='', database='copo')
                        mycommand = mydatabase.cursor()
                        q = "insert into mst1(rollno,name,q1,q2,q3,q4,q5,q6)values('{}','{}','{}','{}','{}','{}','{}','{}')".format(rol,nm,a,b,c,d,e,f)
                        mycommand.execute(q)
                        mydatabase.commit()
                        tkinter.messagebox.showwarning("Error", "Values added Successfully")
                    except Exception as ex:
                        tkinter.messagebox.showwarning("Error", ex)
                else:
                    tkinter.messagebox.showwarning("Error", "Please Enter Student Name")
            else:
                tkinter.messagebox.showwarning("Error", "Please Enter Roll No.")


        Add_btn = Button(window_frame, text='Add', fg='white', bg='black', font='copper 15', command=save)
        Add_btn.place(x=300, y=360)

        def destroy_save():
            mydatabase = mysql.connector.connect(host='localhost', user='root', password='', database='copo')
            mycommand = mydatabase.cursor()
            q = "delete from mst1"
            mycommand.execute(q)
            mydatabase.commit()
            # tkinter.messagebox.showwarning("Error", "Values cleared Successfully.")

            mydatabase = mysql.connector.connect(host='localhost', user='root', password='', database='copo')
            mycommand = mydatabase.cursor()
            m = "alter table mst1 AUTO_INCREMENT = 1"
            mycommand.execute(m)
            mydatabase.commit()
            tkinter.messagebox.showwarning("Error", "Values cleared Successfully.")

        def cal():
            # Establishing the database connection
            connection_string = "mysql+pymysql://root:@localhost/copo"
            engine = create_engine(connection_string)

            def fetch_data(query):
                try:
                    return pd.read_sql(query, engine)
                except Exception as e:
                    print(f"An error occurred while fetching data: {e}")
                    return None

            # Fetching data from tables
            df_matrix = fetch_data("SELECT * FROM matrix")
            df_table = fetch_data("SELECT * FROM mst1")

            columns = ['q1', 'q2', 'q3', 'q4', 'q5', 'q6']
            new_column_names = ['AQ1', 'AQ2', 'AQ3', 'AQ4', 'AQ5', 'AQ6']

            results = []
            stats = []

            for old_column, new_column in zip(columns, new_column_names):
                query = f"""
                    SELECT 
                        (SELECT COUNT(*) FROM mst1 WHERE {old_column} > (SELECT AVG({old_column}) FROM mst1)) AS N1,
                        (SELECT COUNT(*) FROM mst1 WHERE {old_column} < (SELECT AVG({old_column}) FROM mst1)) AS N2,
                        (SELECT COUNT(*) FROM mst1) AS N
                """
                df_temp = fetch_data(query)
                if df_temp is not None:
                    N1, N2, N = df_temp.at[0, 'N1'], df_temp.at[0, 'N2'], df_temp.at[0, 'N']
                    score = (100 * N1 + 50 * N2) / N if N > 0 else None
                    results.append({'Column': new_column, 'Score': score})
                    stats.append({'Column': old_column, 'N1': N1, 'N2': N2, 'N': N})

            df_results = pd.DataFrame(results)
            df_stats = pd.DataFrame(stats)

            # Additional Calculations
            df_results.set_index('Column', inplace=True)
            A1CO1 = (df_results.loc['AQ1', 'Score'] + df_results.loc['AQ2', 'Score']) / 2
            A1CO2 = (df_results.loc['AQ1', 'Score'] + df_results.loc['AQ2', 'Score'] + df_results.loc['AQ3', 'Score'] + df_results.loc['AQ4', 'Score']) / 4
            A1CO3 = (df_results.loc['AQ1', 'Score'] + df_results.loc['AQ3', 'Score'] + df_results.loc['AQ4', 'Score']) / 3

            additional_calcs = pd.DataFrame({
                'A1CO1': [A1CO1],
                'A1CO2': [A1CO2],
                'A1CO3': [A1CO3]
            })

            # Writing to Excel
            excel_file_path = 'mst-1 calculations.xlsx'
            with pd.ExcelWriter(excel_file_path, engine='openpyxl') as writer:
                df_table.to_excel(writer, sheet_name='mst-1', index=False)
                df_matrix.to_excel(writer, sheet_name='Matrix Data', index=False)
                df_results.reset_index().to_excel(writer, sheet_name='Scores', index=False)
                df_stats.to_excel(writer, sheet_name='Statistics', index=False)
                additional_calcs.to_excel(writer, sheet_name='Additional Calculations', index=False)

            print(df_results)
            print(additional_calcs)
            print(f"Matrix data and results, including additional calculations, have been saved to {excel_file_path}.")



        cal_btn = Button(window_frame, text='Calculate', fg='white', bg='black', font='copper 15', command=cal)
        cal_btn.place(x=280, y=400)

        clear_btn = Button(window_frame, text='Clear', fg='white', bg='black', font='copper 15', command=destroy_save)
        clear_btn.place(x=380, y=360)




        close_button=Button(window_frame, text="X", fg="white", bg="red", font="copper 20", cursor="hand2", command=close_frame)
        close_button.place(x=600, y=8, width=50, height=50)

    def mst2():
        tkinter.messagebox.showwarning("Error", "First Click On The Clear Button to clear and reset the database and then enter your data.")
        window_frame = Frame(wel_root, bg='black')
        window_frame.place(x=50, y=50, width=800, height=500)

        def close_frame():
            window_frame.destroy()

        roll_no = Label(window_frame, text="Roll No: ", fg='white', bg='black', font='copper 15')
        roll_no.place(x=50, y=40)
        roll_no_entry = Entry(window_frame, font='copper 15')
        roll_no_entry.place(x=200, y=40)

        std_name = Label(window_frame, text='Student Name: ', fg='white', bg='black', font='copper 15')
        std_name.place(x=50, y=80)
        std_name_entry = Entry(window_frame, font='copper 15')
        std_name_entry.place(x=200, y=80)

        q1 = Label(window_frame, text='Q1: ', fg='white', bg='black', font='copper 15')
        q1.place(x=50, y=120)
        q1_entry = Entry(window_frame, font='copper 15')
        q1_entry.place(x=200, y=120)

        q2 = Label(window_frame, text='Q2: ', fg='white', bg='black', font='copper 15')
        q2.place(x=50, y=160)
        q2_entry = Entry(window_frame, font='copper 15')
        q2_entry.place(x=200, y=160)

        q3 = Label(window_frame, text='Q3: ', fg='white', bg='black', font='copper 15')
        q3.place(x=50, y=200)
        q3_entry = Entry(window_frame, font='copper 15')
        q3_entry.place(x=200, y=200)

        q4 = Label(window_frame, text='Q4: ', fg='white', bg='black', font='copper 15')
        q4.place(x=50, y=240)
        q4_entry = Entry(window_frame, font='copper 15')
        q4_entry.place(x=200, y=240)

        q5 = Label(window_frame, text='Q5: ', fg='white', bg='black', font='copper 15')
        q5.place(x=50, y=280)
        q5_entry = Entry(window_frame, font='copper 15')
        q5_entry.place(x=200, y=280)

        q6 = Label(window_frame, text='Q6: ', fg='white', bg='black', font='copper 15')
        q6.place(x=50, y=320)
        q6_entry = Entry(window_frame, font='copper 15')
        q6_entry.place(x=200, y=320)

        matrix1 = Button(window_frame, text='matrix', fg='white', bg='black', font='copper 20', command=matrix)
        matrix1.place(x=500, y=300)



        def save():
            rol = roll_no_entry.get()
            nm = std_name_entry.get()
            a = q1_entry.get()
            b = q2_entry.get()
            c = q3_entry.get()
            d = q4_entry.get()
            e = q5_entry.get()
            f = q6_entry.get()
            if rol.strip() != '':
                if nm.strip() != '':
                    try:
                        mydatabase = mysql.connector.connect(host='localhost', user='root', password='', database='copo')
                        mycommand = mydatabase.cursor()
                        q = "insert into mst2(rollno,name,q1,q2,q3,q4,q5,q6)values('{}','{}','{}','{}','{}','{}','{}','{}')".format(rol,nm,a,b,c,d,e,f)
                        mycommand.execute(q)
                        mydatabase.commit()
                        tkinter.messagebox.showwarning("Error", "Values added Successfully")
                    except Exception as ex:
                        tkinter.messagebox.showwarning("Error", ex)
                else:
                    tkinter.messagebox.showwarning("Error", "Please Enter Student Name")
            else:
                tkinter.messagebox.showwarning("Error", "Please Enter Roll No.")


        Add_btn = Button(window_frame, text='Add', fg='white', bg='black', font='copper 15', command=save)
        Add_btn.place(x=300, y=360)

        def destroy_save():
            mydatabase = mysql.connector.connect(host='localhost', user='root', password='', database='copo')
            mycommand = mydatabase.cursor()
            q = "delete from mst2"
            mycommand.execute(q)
            mydatabase.commit()
            # tkinter.messagebox.showwarning("Error", "Values cleared Successfully.")

            mydatabase = mysql.connector.connect(host='localhost', user='root', password='', database='copo')
            mycommand = mydatabase.cursor()
            m = "alter table mst2 AUTO_INCREMENT = 1"
            mycommand.execute(m)
            mydatabase.commit()
            tkinter.messagebox.showwarning("Error", "Values cleared Successfully.")

        def cal():
            # Establishing the database connection
            connection_string = "mysql+pymysql://root:@localhost/copo"
            engine = create_engine(connection_string)

            def fetch_data(query):
                try:
                    return pd.read_sql(query, engine)
                except Exception as e:
                    print(f"An error occurred while fetching data: {e}")
                    return None

            # Fetching data from tables
            df_matrix = fetch_data("SELECT * FROM matrix")
            df_table = fetch_data("SELECT * FROM mst2")

            columns = ['q1', 'q2', 'q3', 'q4', 'q5', 'q6']
            new_column_names = ['AQ1', 'AQ2', 'AQ3', 'AQ4', 'AQ5', 'AQ6']

            results = []
            stats = []

            for old_column, new_column in zip(columns, new_column_names):
                query = f"""
                    SELECT 
                        (SELECT COUNT(*) FROM mst2 WHERE {old_column} > (SELECT AVG({old_column}) FROM mst2)) AS N1,
                        (SELECT COUNT(*) FROM mst2 WHERE {old_column} < (SELECT AVG({old_column}) FROM mst2)) AS N2,
                        (SELECT COUNT(*) FROM mst2) AS N
                """
                df_temp = fetch_data(query)
                if df_temp is not None:
                    N1, N2, N = df_temp.at[0, 'N1'], df_temp.at[0, 'N2'], df_temp.at[0, 'N']
                    score = (100 * N1 + 50 * N2) / N if N > 0 else None
                    results.append({'Column': new_column, 'Score': score})
                    stats.append({'Column': old_column, 'N1': N1, 'N2': N2, 'N': N})

            df_results = pd.DataFrame(results)
            df_stats = pd.DataFrame(stats)

            # Additional Calculations
            df_results.set_index('Column', inplace=True)
            A2CO3 = (df_results.loc['AQ1', 'Score'] + df_results.loc['AQ2', 'Score'] + df_results.loc['AQ3', 'Score']) / 3
            A2CO4 = df_results.loc['AQ1', 'Score']
            A2CO5 = df_results.loc['AQ4', 'Score']
            A2CO6 = (df_results.loc['AQ1', 'Score'] + df_results.loc['AQ3', 'Score'] + df_results.loc['AQ4', 'Score']) / 3

            additional_calcs = pd.DataFrame({
                'A2CO3': [A2CO3],
                'A2CO4': [A2CO4],
                'A2CO5': [A2CO5],
                'A2CO6': [A2CO6]
            })

            # Writing to Excel
            excel_file_path = 'mst-2 calculations.xlsx'
            with pd.ExcelWriter(excel_file_path, engine='openpyxl') as writer:
                df_table.to_excel(writer, sheet_name='mst-1', index=False)
                df_matrix.to_excel(writer, sheet_name='Matrix Data', index=False)
                df_results.reset_index().to_excel(writer, sheet_name='Scores', index=False)
                df_stats.to_excel(writer, sheet_name='Statistics', index=False)
                additional_calcs.to_excel(writer, sheet_name='Additional Calculations', index=False)

            print(df_results)
            print(additional_calcs)
            print(f"Matrix data and results, including additional calculations, have been saved to {excel_file_path}.")



            # w = """insert into calculate(Q1,Q2,Q3,Q4,Q5,Q6) select AVG(q1) AS Q1, AVG(q2) AS Q2, AVG(q3) AS Q3, AVG(q4) AS Q4, AVG(q5) AS Q5, AVG(q6) AS Q6 from mst1"""
            # mycommand.execute(w)
            # mydatabase1.commit()
            # tkinter.messagebox.showwarning("Error", "Value Calculated Successfully")



        cal_btn = Button(window_frame, text='Calculate', fg='white', bg='black', font='copper 15', command=cal)
        cal_btn.place(x=280, y=400)

        clear_btn = Button(window_frame, text='Clear', fg='white', bg='black', font='copper 15', command=destroy_save)
        clear_btn.place(x=380, y=360)




        close_button=Button(window_frame, text="X", fg="white", bg="red", font="copper 20", cursor="hand2", command=close_frame)
        close_button.place(x=600, y=8, width=50, height=50)

    def mst3():
        tkinter.messagebox.showwarning("Error", "First Click On The Clear Button to clear and reset the database and then enter your data.")
        window_frame = Frame(wel_root, bg='black')
        window_frame.place(x=50, y=50, width=800, height=500)

        def close_frame():
            window_frame.destroy()

        roll_no = Label(window_frame, text="Roll No: ", fg='white', bg='black', font='copper 15')
        roll_no.place(x=50, y=40)
        roll_no_entry = Entry(window_frame, font='copper 15')
        roll_no_entry.place(x=200, y=40)

        std_name = Label(window_frame, text='Student Name: ', fg='white', bg='black', font='copper 15')
        std_name.place(x=50, y=80)
        std_name_entry = Entry(window_frame, font='copper 15')
        std_name_entry.place(x=200, y=80)

        q1 = Label(window_frame, text='Q1: ', fg='white', bg='black', font='copper 15')
        q1.place(x=50, y=120)
        q1_entry = Entry(window_frame, font='copper 15')
        q1_entry.place(x=200, y=120)

        q2 = Label(window_frame, text='Q2: ', fg='white', bg='black', font='copper 15')
        q2.place(x=50, y=160)
        q2_entry = Entry(window_frame, font='copper 15')
        q2_entry.place(x=200, y=160)

        q3 = Label(window_frame, text='Q3: ', fg='white', bg='black', font='copper 15')
        q3.place(x=50, y=200)
        q3_entry = Entry(window_frame, font='copper 15')
        q3_entry.place(x=200, y=200)

        q4 = Label(window_frame, text='Q4: ', fg='white', bg='black', font='copper 15')
        q4.place(x=50, y=240)
        q4_entry = Entry(window_frame, font='copper 15')
        q4_entry.place(x=200, y=240)

        q5 = Label(window_frame, text='Q5: ', fg='white', bg='black', font='copper 15')
        q5.place(x=50, y=280)
        q5_entry = Entry(window_frame, font='copper 15')
        q5_entry.place(x=200, y=280)

        q6 = Label(window_frame, text='Q6: ', fg='white', bg='black', font='copper 15')
        q6.place(x=50, y=320)
        q6_entry = Entry(window_frame, font='copper 15')
        q6_entry.place(x=200, y=320)

        matrix1 = Button(window_frame, text='matrix', fg='white', bg='black', font='copper 20', command=matrix)
        matrix1.place(x=500, y=300)



        def save():
            rol = roll_no_entry.get()
            nm = std_name_entry.get()
            a = q1_entry.get()
            b = q2_entry.get()
            c = q3_entry.get()
            d = q4_entry.get()
            e = q5_entry.get()
            f = q6_entry.get()
            if rol.strip() != '':
                if nm.strip() != '':
                    try:
                        mydatabase = mysql.connector.connect(host='localhost', user='root', password='', database='copo')
                        mycommand = mydatabase.cursor()
                        q = "insert into mst3(rollno,name,q1,q2,q3,q4,q5,q6)values('{}','{}','{}','{}','{}','{}','{}','{}')".format(rol,nm,a,b,c,d,e,f)
                        mycommand.execute(q)
                        mydatabase.commit()
                        tkinter.messagebox.showwarning("Error", "Values added Successfully")
                    except Exception as ex:
                        tkinter.messagebox.showwarning("Error", ex)
                else:
                    tkinter.messagebox.showwarning("Error", "Please Enter Student Name")
            else:
                tkinter.messagebox.showwarning("Error", "Please Enter Roll No.")


        Add_btn = Button(window_frame, text='Add', fg='white', bg='black', font='copper 15', command=save)
        Add_btn.place(x=300, y=360)

        def destroy_save():
            mydatabase = mysql.connector.connect(host='localhost', user='root', password='', database='copo')
            mycommand = mydatabase.cursor()
            q = "delete from mst3"
            mycommand.execute(q)
            mydatabase.commit()
            # tkinter.messagebox.showwarning("Error", "Values cleared Successfully.")

            mydatabase = mysql.connector.connect(host='localhost', user='root', password='', database='copo')
            mycommand = mydatabase.cursor()
            m = "alter table mst3 AUTO_INCREMENT = 1"
            mycommand.execute(m)
            mydatabase.commit()
            tkinter.messagebox.showwarning("Error", "Values cleared Successfully.")

        def cal():
            # Establishing the database connection
            connection_string = "mysql+pymysql://root:@localhost/copo"
            engine = create_engine(connection_string)

            def fetch_data(query):
                try:
                    return pd.read_sql(query, engine)
                except Exception as e:
                    print(f"An error occurred while fetching data: {e}")
                    return None

            # Fetching data from tables
            df_matrix = fetch_data("SELECT * FROM matrix")
            df_table = fetch_data("SELECT * FROM mst3")

            columns = ['q1', 'q2', 'q3', 'q4', 'q5', 'q6']
            new_column_names = ['AQ1', 'AQ2', 'AQ3', 'AQ4', 'AQ5', 'AQ6']

            results = []
            stats = []

            for old_column, new_column in zip(columns, new_column_names):
                query = f"""
                    SELECT 
                        (SELECT COUNT(*) FROM mst3 WHERE {old_column} > (SELECT AVG({old_column}) FROM mst3)) AS N1,
                        (SELECT COUNT(*) FROM mst3 WHERE {old_column} < (SELECT AVG({old_column}) FROM mst3)) AS N2,
                        (SELECT COUNT(*) FROM mst3) AS N
                """
                df_temp = fetch_data(query)
                if df_temp is not None:
                    N1, N2, N = df_temp.at[0, 'N1'], df_temp.at[0, 'N2'], df_temp.at[0, 'N']
                    score = (100 * N1 + 50 * N2) / N if N > 0 else None
                    results.append({'Column': new_column, 'Score': score})
                    stats.append({'Column': old_column, 'N1': N1, 'N2': N2, 'N': N})

            df_results = pd.DataFrame(results)
            df_stats = pd.DataFrame(stats)

            # Additional Calculations
            df_results.set_index('Column', inplace=True)
            A3CO2 = (df_results.loc['AQ1', 'Score'] + df_results.loc['AQ4', 'Score']) / 2
            A3CO3 = (df_results.loc['AQ1', 'Score'] + df_results.loc['AQ4', 'Score']) / 2
            A3CO4 = df_results.loc['AQ1', 'Score']
            A3CO5 = df_results.loc['AQ3', 'Score']
            A3CO6 = (df_results.loc['AQ2', 'Score'] + df_results.loc['AQ3', 'Score']) / 2

            additional_calcs = pd.DataFrame({
                'A3CO2': [A3CO2],
                'A3CO3': [A3CO3],
                'A3CO4': [A3CO4],
                'A3CO5': [A3CO5],
                'A3CO6': [A3CO6]
            })

            # Writing to Excel
            excel_file_path = 'mst-3 calculations.xlsx'
            with pd.ExcelWriter(excel_file_path, engine='openpyxl') as writer:
                df_table.to_excel(writer, sheet_name='mst-3', index=False)
                df_matrix.to_excel(writer, sheet_name='Matrix Data', index=False)
                df_results.reset_index().to_excel(writer, sheet_name='Scores', index=False)
                df_stats.to_excel(writer, sheet_name='Statistics', index=False)
                additional_calcs.to_excel(writer, sheet_name='Additional Calculations', index=False)

            print(df_results)
            print(additional_calcs)
            print(f"Matrix data and results, including additional calculations, have been saved to {excel_file_path}.")



            # w = """insert into calculate(Q1,Q2,Q3,Q4,Q5,Q6) select AVG(q1) AS Q1, AVG(q2) AS Q2, AVG(q3) AS Q3, AVG(q4) AS Q4, AVG(q5) AS Q5, AVG(q6) AS Q6 from mst1"""
            # mycommand.execute(w)
            # mydatabase1.commit()
            # tkinter.messagebox.showwarning("Error", "Value Calculated Successfully")



        cal_btn = Button(window_frame, text='Calculate', fg='white', bg='black', font='copper 15', command=cal)
        cal_btn.place(x=280, y=400)

        clear_btn = Button(window_frame, text='Clear', fg='white', bg='black', font='copper 15', command=destroy_save)
        clear_btn.place(x=380, y=360)




        close_button=Button(window_frame, text="X", fg="white", bg="red", font="copper 20", cursor="hand2", command=close_frame)
        close_button.place(x=600, y=8, width=50, height=50)


    def assignment1():
        tkinter.messagebox.showwarning("Error", "First Click On The Clear Button to clear and reset the database and then enter your data.")
        window_frame = Frame(wel_root, bg='black')
        window_frame.place(x=50, y=50, width=800, height=500)

        def close_frame():
            window_frame.destroy()

        roll_no = Label(window_frame, text="Roll No: ", fg='white', bg='black', font='copper 15')
        roll_no.place(x=50, y=40)
        roll_no_entry = Entry(window_frame, font='copper 15')
        roll_no_entry.place(x=200, y=40)

        std_name = Label(window_frame, text='Student Name: ', fg='white', bg='black', font='copper 15')
        std_name.place(x=50, y=80)
        std_name_entry = Entry(window_frame, font='copper 15')
        std_name_entry.place(x=200, y=80)

        A1 = Label(window_frame, text='A1: ', fg='white', bg='black', font='copper 15')
        A1.place(x=50, y=120)
        A1_entry = Entry(window_frame, font='copper 15')
        A1_entry.place(x=200, y=120)

        A2 = Label(window_frame, text='A2: ', fg='white', bg='black', font='copper 15')
        A2.place(x=50, y=160)
        A2_entry = Entry(window_frame, font='copper 15')
        A2_entry.place(x=200, y=160)

        matrix2 = Button(window_frame, text='matrix', fg='white', bg='black', font='copper 20', command=matrix1)
        matrix2.place(x=500, y=300)



        def save():
            rol = roll_no_entry.get()
            nm = std_name_entry.get()
            a = A1_entry.get()
            b = A2_entry.get()
            if rol.strip() != '':
                if nm.strip() != '':
                    try:
                        mydatabase = mysql.connector.connect(host='localhost', user='root', password='', database='copo')
                        mycommand = mydatabase.cursor()
                        q = "insert into assignment(rollno,name,A1,A2)values('{}','{}','{}','{}')".format(rol,nm,a,b)
                        mycommand.execute(q)
                        mydatabase.commit()
                        tkinter.messagebox.showwarning("Error", "Values added Successfully")
                    except Exception as ex:
                        tkinter.messagebox.showwarning("Error", ex)
                else:
                    tkinter.messagebox.showwarning("Error", "Please Enter Student Name")
            else:
                tkinter.messagebox.showwarning("Error", "Please Enter Roll No.")


        Add_btn = Button(window_frame, text='Add', fg='white', bg='black', font='copper 15', command=save)
        Add_btn.place(x=300, y=360)

        def destroy_save():
            mydatabase = mysql.connector.connect(host='localhost', user='root', password='', database='copo')
            mycommand = mydatabase.cursor()
            q = "delete from assignment"
            mycommand.execute(q)
            mydatabase.commit()
            # tkinter.messagebox.showwarning("Error", "Values cleared Successfully.")

            mydatabase = mysql.connector.connect(host='localhost', user='root', password='', database='copo')
            mycommand = mydatabase.cursor()
            m = "alter table assignment AUTO_INCREMENT = 1"
            mycommand.execute(m)
            mydatabase.commit()
            tkinter.messagebox.showwarning("Error", "Values cleared Successfully.")

        def cal():
            # Establish database connection using SQLAlchemy
            connection_string = "mysql+pymysql://root:@localhost/copo"
            engine = create_engine(connection_string)

            # Function to fetch data from the database
            def fetch_data(query):
                return pd.read_sql(query, engine)

            # Fetch grades data for both assignments
            df_grades = fetch_data("SELECT * FROM assignment")  # Fetch the entire table

            # Function to calculate statistics and apply the formula
            def calculate_stats_and_formula(df, column):
                counts = df[column].value_counts().reindex(['A', 'B', 'C'], fill_value=0)
                A = counts.get('A', 0)
                B = counts.get('B', 0)
                C = counts.get('C', 0)
                N = df[column].count()

                # Apply the formula
                if N > 0:
                    score = ((A * 100) + (B * 50) + (C * 25)) / N
                else:
                    score = None  # Handle case where N is 0 to avoid division by zero

                return {
                    'Total Students (N)': N,
                    'Grade A Count': A,
                    'Grade B Count': B,
                    'Grade C Count': C,
                    'Calculated Score': score
                }

            # Calculating statistics and scores for each assignment
            stats_A1 = calculate_stats_and_formula(df_grades, 'A1')
            stats_A2 = calculate_stats_and_formula(df_grades, 'A2')

            # Create DataFrames for the results
            df_results_A1 = pd.DataFrame([stats_A1], index=['Assignment 1'])
            df_results_A2 = pd.DataFrame([stats_A2], index=['Assignment 2'])

            # Create the combined report DataFrame
            data = {
                'AOCO1': [stats_A1['Calculated Score']],
                'AOCO2': [(stats_A1['Calculated Score'] + stats_A2['Calculated Score']) / 2],
                'AOCO3': [(stats_A1['Calculated Score'] + stats_A2['Calculated Score']) / 2],
                'AOCO4': [stats_A2['Calculated Score']],
                'AOCO5': [stats_A2['Calculated Score']],
                'AOCO6': [stats_A2['Calculated Score']]
                }
            df_report = pd.DataFrame(data)

            # Specify the Excel file path
            excel_file_path = 'grades_summary.xlsx'

            # Write both dataframes to different sheets within the same Excel file
            with pd.ExcelWriter(excel_file_path, engine='openpyxl') as writer:
                df_grades.to_excel(writer, sheet_name='Student Grades', index=False)  # Full student grades table
                df_results_A1.to_excel(writer, sheet_name='Grade Statistics A1', index=True)  # Statistics for A1
                df_results_A2.to_excel(writer, sheet_name='Grade Statistics A2', index=True)  # Statistics for A2
                df_report.to_excel(writer, sheet_name='Combined Report', index=False) # Course outcome calculation

            print(df_results_A1)
            print(df_results_A2)
            print(f"Grade statistics for both assignments and complete student grades have been saved to {excel_file_path}.")





        cal_btn = Button(window_frame, text='Calculate', fg='white', bg='black', font='copper 15', command=cal)
        cal_btn.place(x=280, y=400)

        clear_btn = Button(window_frame, text='Clear', fg='white', bg='black', font='copper 15', command=destroy_save)
        clear_btn.place(x=380, y=360)

        close_button=Button(window_frame, text="X", fg="white", bg="red", font="copper 20", cursor="hand2", command=close_frame)
        close_button.place(x=600, y=8, width=50, height=50)

    def fetch_calculated_values(file_paths, sheet_name='Additional Calculations'):
        combined_values = {}
        try:
            for file_path, keys in file_paths:
                data = pd.read_excel(file_path, sheet_name=sheet_name)
                for key in keys:
                    combined_values[key] = data[key].values[0] if key in data.columns else None
        except Exception as e:
            tkinter.messagebox.showerror("Error", f"An error occurred while fetching the data from {file_path}: {e}")
            return None
        return combined_values

    def perform_calculations_and_save(values):
        sco_values = {
            'SCO1': values['A1CO1'],
            'SCO2': (values['A1CO2'] + values['A3CO3']) / 3 if values['A1CO2'] is not None and values['A3CO3'] is not None else None,
            'SCO3': (values['A1CO3'] + values['A2CO3'] + values['A3CO3']) / 3 if values['A1CO3'] is not None and values['A2CO3'] is not None and values['A3CO3'] is not None else None,
            'SCO4': (values['A2CO4'] + values['A3CO4']) / 2 if values['A2CO4'] is not None and values['A3CO4'] is not None else None,
            'SCO5': (values['A2CO5'] + values['A3CO5']) / 2 if values['A2CO5'] is not None and values['A3CO5'] is not None else None,
            'SCO6': (values['A2CO6'] + values['A3CO6']) / 2 if values['A2CO6'] is not None and values['A3CO6'] is not None else None
        }

        results_df = pd.DataFrame([sco_values])
        output_excel_path = 'SCO_Results.xlsx'
        with pd.ExcelWriter(output_excel_path, engine='openpyxl') as writer:
            results_df.to_excel(writer, index=False, sheet_name='Calculated Scores')
        tkinter.messagebox.showinfo("Success", f"Calculated scores have been saved to {output_excel_path}.")

    def sco():
        file_paths = [
            ('mst-1 calculations.xlsx', ['A1CO1', 'A1CO2', 'A1CO3']),
            ('mst-2 calculations.xlsx', ['A2CO3', 'A2CO4', 'A2CO5', 'A2CO6']),
            ('mst-3 calculations.xlsx', ['A3CO2', 'A3CO3', 'A3CO4', 'A3CO5', 'A3CO6'])
        ]
        values = fetch_calculated_values(file_paths)
        if values:
            perform_calculations_and_save(values)

    # Button to trigger calculations
    calculate_button = Button(wel_root, text="Calculate SCO", fg='white', bg='black', font='copper 20', command=sco)
    calculate_button.place(x=450, y=300)





    mst_1 = Button(wel_root, text='MST-I', fg='white', bg='black', font='copper 20', command=mst1)
    mst_1.place(x=30, y=100)

    mst_2 = Button(wel_root, text='MST-II', fg='white', bg='black', font='copper 20', command=mst2)
    mst_2.place(x=30, y=200)

    mst_3 = Button(wel_root, text='MST-III', fg='white', bg='black', font='copper 20', command=mst3)
    mst_3.place(x=30, y=300)

    assignment = Button(wel_root, text='Assignments', fg='white', bg='black', font='copper 20', command=assignment1)
    assignment.place(x=30, y=400)

    wel_root.mainloop()
