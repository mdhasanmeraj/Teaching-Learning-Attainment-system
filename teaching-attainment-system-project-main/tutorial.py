




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
            # Establishing the connection using SQLAlchemy
            connection_string = "mysql+pymysql://root:@localhost/copo"
            engine = create_engine(connection_string)

            # Function to fetch data from the database
            def fetch_data(query):
                return pd.read_sql(query, engine)

            # Fetching data from both tables
            df_matrix = fetch_data("SELECT * FROM matrix")
            df_table = fetch_data("SELECT * FROM mst3")

            # Original column names for database queries
            columns = ['q1', 'q2', 'q3', 'q4', 'q5', 'q6']
            # Mapping original columns to new names for the scores presentation
            new_column_names = ['AQ1', 'AQ2', 'AQ3', 'AQ4', 'AQ5', 'AQ6']

            results = []
            stats = []

            for old_column, new_column in zip(columns, new_column_names):
                df_temp = fetch_data(f"""
                    SELECT 
                        (SELECT COUNT(*) FROM mst3 WHERE {old_column} > (SELECT AVG({old_column}) FROM mst3)) AS N1,
                        (SELECT COUNT(*) FROM mst3 WHERE {old_column} < (SELECT AVG({old_column}) FROM mst3)) AS N2,
                        (SELECT COUNT(*) FROM mst3) AS N
                """)
                N1, N2, N = df_temp.at[0, 'N1'], df_temp.at[0, 'N2'], df_temp.at[0, 'N']
                score = (100 * N1 + 50 * N2) / N if N > 0 else None  # Calculating the score safely
                results.append({'Column': new_column, 'Score': score})
                stats.append({'Column': old_column, 'N1': N1, 'N2': N2, 'N': N})

            df_results = pd.DataFrame(results)
            df_stats = pd.DataFrame(stats)

            # Writing to Excel
            excel_file_path = 'combined_output.xlsx'
            with pd.ExcelWriter(excel_file_path, engine='openpyxl') as writer:
                df_table.to_excel(writer, sheet_name='mst-3', index=False)
                df_matrix.to_excel(writer, sheet_name='Matrix Data', index=False)
                df_results.to_excel(writer, sheet_name='Scores', index=False)  # Results using new column names
                df_stats.to_excel(writer, sheet_name='Statistics', index=False)  # Statistics using original column names

            print(df_results)
            print(f"Matrix data and results have been saved to {excel_file_path}.")


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
