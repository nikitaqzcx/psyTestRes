import tkinter as tk
from datetime import datetime
import customtkinter as ctk
from tkinter import messagebox, ttk
import tkfilebrowser
import sqlite3
from tkcalendar import Calendar
import locale
import openpyxl
import xlrd
from lasarus import LasarusResults


class App:
    def __init__(self, root):
        self.root = root
        self.root.title("PsyTestResults")
        self.root.geometry("550x800")  # Set the width and height of the window (width x height)
        self.root.minsize(550, 800)

        locale.setlocale(locale.LC_TIME, "uk_UA.UTF-8")
        # CustomTkinter styling
        ctk.set_appearance_mode("System")  # Use system appearance mode (light/dark)
        ctk.set_default_color_theme("blue")  # Choose a color theme

        # SQLite setup
        self.conn = sqlite3.connect('app_data.db')
        self.cursor = self.conn.cursor()
        self.create_table()

        # UI Elements
        self.create_widgets()
        self.load_data()

    def create_table(self):
        self.cursor.execute('''
            CREATE TABLE IF NOT EXISTS TestStructures (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                Name TEXT,
                NameColumn TEXT,
                DateColumn TEXT,
                DataColumn TEXT
            )
        ''')
        self.conn.commit()

        # Insert default data if the table is empty
        self.cursor.execute("SELECT COUNT(*) FROM TestStructures")
        if self.cursor.fetchone()[0] == 0:
            self.insert_default_data()

    def insert_default_data(self):
        # Insert default values into TestStructures
        self.cursor.executemany('''
            INSERT INTO TestStructures (Name, NameColumn, DateColumn, DataColumn) VALUES (?, ?, ?, ?)
        ''', [
            ('Адаптивність 200', 'C', 'A', 'HA,HC,HE,HG,HI,HK,HM,HO'),
            ('Соціоніка', 'B', 'A', 'BS'),
            ('Акцентуація Особистості', 'B', 'A', 'GD,GF,GH,GJ,GL,GN,GP,GR,GT,GV,GX,GZ')
        ])
        self.conn.commit()

    def create_widgets(self):
        # Top-left button for open file
        self.open_file_button = ctk.CTkButton(self.root, text="Відкрити", command=self.open_file)
        self.open_file_button.grid(row=0, column=0, padx=10, pady=10, sticky="nw")

        # Label to display the selected file name
        self.selected_file_label = ctk.CTkLabel(self.root, text="Файл не вибрано")
        self.selected_file_label.grid(row=1, column=0, padx=10, sticky="w", columnspan=2)

        # Listbox for items (show only Name)
        self.listbox = ttk.Treeview(self.root, columns=("Name"), show="headings", height=5)
        self.listbox.heading("Name", text="Назва тесту")
        self.listbox.grid(row=2, column=0, padx=10, pady=10, columnspan=2, sticky="nsew")

        # Add, Edit, Remove Buttons
        self.add_button = ctk.CTkButton(self.root, text="Добавити", command=self.add_item)
        self.add_button.grid(row=3, column=0, columnspan=2, padx=10, pady=10, sticky='ew')
        self.edit_button = ctk.CTkButton(self.root, text="Редагувати", command=self.edit_item)
        self.edit_button.grid(row=4, column=0, pady=10,columnspan=2, padx=10, sticky='ew')
        self.remove_button = ctk.CTkButton(self.root, text="Видалити", command=self.remove_item, fg_color='#e84f5e', hover_color='#85323b')
        self.remove_button.grid(row=5, column=0,pady=10, padx=10,columnspan=2, sticky='ew')

        # Radio buttons for time limitation
        self.time_limit_var = tk.StringVar(value="all")
        self.radiobutton_all = ctk.CTkRadioButton(self.root, text="Усі данні", variable=self.time_limit_var, value="all")
        self.radiobutton_all.grid(row=6, column=0, padx=10, pady=10, sticky="w")
        self.radiobutton_time = ctk.CTkRadioButton(self.root, text="Дані за період", variable=self.time_limit_var, value="time")
        self.radiobutton_time.grid(row=6, column=1, padx=10, pady=10, sticky="w")

        # Date selection
        self.start_date = Calendar(self.root, locale="uk_UA")
        self.start_date.grid(row=7, column=0, padx=10, pady=10)
        self.end_date = Calendar(self.root, locale="uk_UA")
        self.end_date.grid(row=7, column=1, padx=10, pady=10)

        # New Treeview list below the date pickers
        self.sheet_listbox = ttk.Treeview(self.root, columns=("Data"), show="headings", height=5)
        self.sheet_listbox.heading("Data", text="Листи з данними")
        self.sheet_listbox.grid(row=8, column=0, padx=10, pady=10, columnspan=2, sticky="nsew")
        
        # Save button
        self.save_button = ctk.CTkButton(self.root, text="Зберегти", command=self.save_data)
        self.save_button.grid(row=9, column=0, padx=10, pady=10, columnspan=3)

    def load_data(self):
        # Clear current list and load data from database
        for row in self.listbox.get_children():
            self.listbox.delete(row)
        for row in self.cursor.execute("SELECT id, Name, NameColumn, DateColumn, DataColumn FROM TestStructures"):
            self.listbox.insert("", "end", values=(row[1],), tags=(row[0],))  # Only show Name, store id as tag

    def open_file(self):
        # Let the user select an Excel file (both .xls and .xlsx)
        file_path = tkfilebrowser.askopenfilename(filetypes=[("Excel files", "*.xls;*.xlsx")])
        if file_path:

            # Store the file path
            self.selected_file_path = file_path
            # Update the label with the selected file's name
            file_name = file_path.split("/")[-1]  # Get the file name without the full path
            self.selected_file_label.configure(text=f"Вибрано: {file_name}")

            # Clear the sheet_listbox before adding any sheet names
            for row in self.sheet_listbox.get_children():
                self.sheet_listbox.delete(row)

            # Open the Excel file and read sheet names
            self.open_excel_file(file_path)

    def open_excel_file(self, file_path):
        # Try opening with openpyxl (for .xlsx files)
        try:
            wb = openpyxl.load_workbook(file_path, keep_vba=False)
            sheet_names = wb.sheetnames
            self.populate_listbox_with_sheet_names(sheet_names)
        except Exception as e:
            # If openpyxl fails, try xlrd for older .xls files
            try:
                wb = xlrd.open_workbook(file_path)
                sheet_names = wb.sheet_names()
                self.populate_listbox_with_sheet_names(sheet_names)
            except Exception as e:
                messagebox.showerror("Помилка", f"Не вдалося відкрити файл: {e}")

    def populate_listbox_with_sheet_names(self, sheet_names):
        # Populate sheet_listbox with sheet names
        for sheet in sheet_names:
            self.sheet_listbox.insert("", "end", values=(sheet,))

    def add_item(self):
        # Open dialog to add a new item
        self.open_item_dialog()

    def edit_item(self):
        selected_item = self.listbox.selection()
        if selected_item:
            # Get the id of the selected item (stored in the tag)
            item_id = self.listbox.item(selected_item)["tags"][0]
            # Fetch the full data for the selected item
            self.cursor.execute("SELECT * FROM TestStructures WHERE id=?", (item_id,))
            item_data = self.cursor.fetchone()
            self.open_item_dialog(item_data, selected_item)
        else:
            messagebox.showwarning("Редагування", "Виберіть тест для редагування")

    def open_item_dialog(self, item_data=None, selected_item=None):
        # Create a dialog to enter item details
        dialog = ctk.CTkToplevel(self.root)
        dialog.title("Add / Edit Item")
        dialog.geometry("300x250")

        # Fields
        name_label = ctk.CTkLabel(dialog, text="Назва:")
        name_label.grid(row=0, column=0, padx=10, pady=10)
        name_entry = ctk.CTkEntry(dialog)
        name_entry.grid(row=0, column=1, padx=10, pady=10)

        name_col_label = ctk.CTkLabel(dialog, text="Колонка з ім'ям:")
        name_col_label.grid(row=1, column=0, padx=10, pady=10)
        name_col_entry = ctk.CTkEntry(dialog)
        name_col_entry.grid(row=1, column=1, padx=10, pady=10)

        date_col_label = ctk.CTkLabel(dialog, text="Колонка з датою:")
        date_col_label.grid(row=2, column=0, padx=10, pady=10)
        date_col_entry = ctk.CTkEntry(dialog)
        date_col_entry.grid(row=2, column=1, padx=10, pady=10)

        data_col_label = ctk.CTkLabel(dialog, text="Колонка з данними:")
        data_col_label.grid(row=3, column=0, padx=10, pady=10)
        data_col_entry = ctk.CTkEntry(dialog)
        data_col_entry.grid(row=3, column=1, padx=10, pady=10)

        # Pre-fill fields if editing
        if item_data:
            name_entry.insert(0, item_data[1])
            name_col_entry.insert(0, item_data[2])
            date_col_entry.insert(0, item_data[3])
            data_col_entry.insert(0, item_data[4])

        def save_item():
            name = name_entry.get()
            name_col = name_col_entry.get()
            date_col = date_col_entry.get()
            data_col = data_col_entry.get()

            if item_data:  # Edit
                query = "UPDATE TestStructures SET Name=?, NameColumn=?, DateColumn=?, DataColumn=? WHERE id=?"
                self.cursor.execute(query, (name, name_col, date_col, data_col, item_data[0]))
            else:  # Add new item
                query = "INSERT INTO TestStructures (Name, NameColumn, DateColumn, DataColumn) VALUES (?, ?, ?, ?)"
                self.cursor.execute(query, (name, name_col, date_col, data_col))

            self.conn.commit()
            self.load_data()  # Reload data into the listbox
            dialog.destroy()

        save_button = ctk.CTkButton(dialog, text="Save", command=save_item)
        save_button.grid(row=4, column=0, columnspan=2, pady=10)

    def remove_item(self):
        selected_item = self.listbox.selection()
        if selected_item:
            # Show confirmation dialog
            confirm = messagebox.askyesno("Підтвердити видалення", "Ви впевнені, що хочете видалити цей елемент?")
            if confirm:  # If the user clicks 'Yes'
                item_id = self.listbox.item(selected_item)["tags"][0]
                self.cursor.execute("DELETE FROM TestStructures WHERE id=?", (item_id,))
                self.conn.commit()
                self.load_data()
                messagebox.showinfo("Елемент видалено", "Елемент успішно видалено.")
        else:
            messagebox.showwarning("Видалити елемент", "Будь ласка, виберіть елемент для видалення.")

    def save_data(self):
        selected_item = self.listbox.selection()
        if selected_item:
            item_id = self.listbox.item(selected_item)["tags"][0]
            self.cursor.execute("SELECT * FROM TestStructures WHERE id=?", (item_id,))
            item_data = self.cursor.fetchone()

            # Gather selected sheet from sheet_listbox
            selected_sheet = self.sheet_listbox.selection()
            if selected_sheet:
                sheet_name = self.sheet_listbox.item(selected_sheet)["values"][0]
            else:
                messagebox.showwarning("Виберіть лист", "Виберіть лист зі списку.")
                return

            # Gather date range if time limitation is selected
            if self.time_limit_var.get() == "time":
                date_from = datetime.strptime(self.start_date.get_date(), '%d.%m.%y')
                date_to = datetime.strptime(self.end_date.get_date(), '%d.%m.%y')
            else:
                date_from = None
                date_to = None

            # Pre-fill the filename with the selected item name
            default_filename = item_data[1]  # Using the item's name for the default filename

            # Prompt user for file save location with pre-filled name
            save_path = tkfilebrowser.asksaveasfilename(defaultextension=".docx", 
                                                    initialfile=default_filename + ".docx", 
                                                    filetypes=[("Word documents", "*.docx")])
            if not save_path:  # If the user cancels the file dialog
                return

            # Create structure for LasarusResults
            structure = {
                'name_column': item_data[2],  # NameColumn
                'date_column': item_data[3],  # DateColumn
                'data_columns': item_data[4]  # DataColumn
            }

            # Assuming 'LasarusResults' is your class that saves the results
            lasarus_results = LasarusResults(self.selected_file_path, date_from, date_to)
            lasarus_results.save_results(sheet_name, save_path, structure)

            messagebox.showinfo("Зберегти дані", "Дані успішно збережено.")
        else:
            messagebox.showwarning("Зберегти дані", "Будь ласка, виберіть елемент для збереження.")


if __name__ == "__main__":
    root = ctk.CTk()
    app = App(root)
    root.mainloop()
