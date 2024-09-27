import tkinter as tk
from pyodbc import connect
from pandas import read_excel
from tkinter import ttk, filedialog

class BeastAutoInput:
    def __init__(self, root):
        self.root = root
        self.root.title("Exceedance Analyser")
        self.root.geometry("650x250")

        self.logo_image = tk.PhotoImage(file="C:/Users/Faiq Raedaya/Documents/Python/ShepherdAnalyser/MES.png")
        self.logo_label = tk.Label(root, image=self.logo_image)
        self.logo_label.grid(row=0, column=0, padx=10, pady=10, sticky="nw")
        
        self.title = tk.Label(root, text="BEAST Auto Input Tool", font=('Arial', 20))
        self.title.grid(row=0, column=1, columnspan=2, pady=10, sticky="w")

        self.create_label("1. Select Excel file:", 1, 1, 10, "w")
        self.import_excel_button = self.create_button("Import", self.import_excel, 1, 2, 1, "w")

        self.create_label("2. Select MDB (BEAST) file:", 2, 1, 10, "w")
        self.import_mdb_button = self.create_button("Import", self.import_mdb, 2, 2, 1, "w")

        self.info_label = tk.Label(root, text="Please import an .xlsx and an .mdb file")
        self.info_label.grid(row=3, column=1, columnspan=3, pady=10)

    def create_label(self, text, row, column, pady, sticky):
        '''
        Create a tkinter label
        '''
        label = tk.Label(self.root, text=text)
        label.grid(row=row, column=column, pady=pady, sticky=sticky)

    def create_button(self, text, command, row, column, columnspan, sticky):
        '''
        Create a tkinter button
        '''
        button = tk.Button(self.root, text=text, command=command, state="active")
        button.grid(row=row, column=column, columnspan=columnspan, pady=10, sticky=sticky)
        return button

    def display_info(self, text):
        '''
        Updates tkinter info label text.
        '''
        self.info_label.config(text=text)
    
    def import_excel(self):
        '''
        Imports the .xlsx file using tkinter filedialog.
        '''
        excel_file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
        if excel_file_path:
            self.excel_data = read_excel(excel_file_path)
            self.display_info("Excel file imported successfully.")
        else:
            self.display_info("Import cancelled.")

    def import_mdb(self):
        '''
        Imports the .mdb file using tkinter filedialog.
        '''
        mdb_file_path = filedialog.askopenfilename(filetypes=[("MDB files", "*.mdb")])
        if mdb_file_path:
            self.display_info("MDB file imported successfully.")
            if hasattr(self, 'excel_data'):
                self.update_mdb(mdb_file_path)
        else:
            self.display_info("Import cancelled.")

    def update_mdb(self, mdb_file_path):
        '''
        Updates the .mdb file using SQL commands
        '''
        self.display_info("Beginning data transfer...")
        # Create a connection using pyodbc
        conn_str = r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=' + mdb_file_path
        connection = connect(conn_str)
        cursor = connection.cursor()

        # Iterate through all rows of the table
        for row in self.excel_data.itertuples(index=False):
            # Define what each row represents
            building_number = row[0]
            overpressure = row[1]
            impulse = row[2]

            # Create SQL query to update the table after setting new values 
            query = f"UPDATE Buildings SET Pressure={overpressure}, Impulse={impulse} WHERE BuildingName LIKE '%{building_number}%'"
            
            # Execute the query
            cursor.execute(query)
            
            # Commit the change
            connection.commit()

        # Terminate connection
        cursor.close()
        connection.close()
        self.display_info("Data updated successfully.")

if __name__ == "__main__":
    root = tk.Tk()
    app = BeastAutoInput(root)
    root.mainloop()