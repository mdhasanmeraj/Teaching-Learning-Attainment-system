import tkinter as tk
import mysql.connector
from mysql.connector import Error

def matrix1():
    def destroy_save():
            mydatabase = mysql.connector.connect(host='localhost', user='root', password='', database='copo')
            mycommand = mydatabase.cursor()
            q = "delete from assignment_matrix"
            mycommand.execute(q)
            mydatabase.commit()
            # tk.messagebox.showwarning("Error", "Values cleared Successfully.")

            mydatabase = mysql.connector.connect(host='localhost', user='root', password='', database='copo')
            mycommand = mydatabase.cursor()
            m = "alter table assignment_matrix AUTO_INCREMENT = 1"
            mycommand.execute(m)
            mydatabase.commit()
            tk.messagebox.showwarning("Error", "Values cleared Successfully.")

    def create_matrix(uroot, rows, cols, column_names):
        # Create column labels
        for i, name in enumerate(column_names):
            label = tk.Label(uroot, text=name)
            label.grid(row=0, column=i, padx=5, pady=5, sticky="w")

        # Create dictionary to hold entry widgets, each key will be (row, col)
        entries = {}
        for i in range(1, rows + 1):
            row_entries = []
            for j in range(cols):
                entry = tk.Entry(uroot)
                entry.grid(row=i, column=j, padx=5, pady=5, sticky="w")
                row_entries.append(entry)
            entries[i] = row_entries
        return entries

    def collect_and_store_data(entries):
        data_to_store = []
        # Collect data from entries
        for row_entries in entries.values():
            row_data = [entry.get() for entry in row_entries]
            data_to_store.append(tuple(row_data))

        # Database connection parameters
        try:
            connection = mysql.connector.connect(host='localhost',
                                                 database='copo',
                                                 user='root',
                                                 password='')
            cursor = connection.cursor()
            # SQL query to insert data
            insert_query = """
            INSERT INTO assignment_matrix(assignment, co1, co2, co3, co4, co5, co6)
            VALUES (%s, %s, %s, %s, %s, %s, %s)
            """
            cursor.executemany(insert_query, data_to_store)
            connection.commit()
            print("Data inserted successfully")
        except Error as error:
            print("Failed to insert data into MySQL table", error)
        finally:
            if connection.is_connected():
                cursor.close()
                connection.close()

    # Main window
    uroot = tk.Tk()
    uroot.title("Matrix Entry with Column Labels")

    # Define matrix dimensions and column names
    rows = 2
    cols = 7
    column_names = ["Assignment/CO", "CO1", "CO2", "CO3", "CO4", "CO5", "CO6"]

    entries = create_matrix(uroot, rows, cols, column_names)

    # Button to collect and store data
    submit_button = tk.Button(uroot, text="Submit", command=lambda: collect_and_store_data(entries))
    submit_button.grid(row=rows+1, columnspan=cols, pady=10)

    clear_button = tk.Button(uroot, text='Clear', command=destroy_save)
    clear_button.grid(row=rows+2, columnspan=cols+1, pady=10)

    uroot.mainloop()
