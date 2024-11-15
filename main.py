import sys
from PySide6.QtCore import QDate, QLocale, Qt
from PySide6.QtWidgets import QButtonGroup, QListView, QFileDialog, QApplication, QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLabel, QRadioButton, QTreeView, QDialog, QFormLayout, QLineEdit, QDateEdit, QMessageBox
from PySide6.QtGui import QStandardItemModel, QStandardItem
import sqlite3
import openpyxl
import xlrd
from lasarus import LasarusResults

class App(QWidget):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("PsyTestResults")
        self.setGeometry(100, 100, 450, 600)

        # SQLite setup
        self.conn = sqlite3.connect('app_data.db')
        self.cursor = self.conn.cursor()
        self.create_table()

        # UI Elements
        self.create_widgets()

        # Data loading
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
        # Layouts
        main_layout = QVBoxLayout(self)
        
        # Label to display the selected file name
        self.selected_file_label = QLabel("Файл не вибрано")
        main_layout.addWidget(self.selected_file_label)

        # Top-left button for open file
        self.open_file_button = QPushButton("Відкрити")
        self.open_file_button.clicked.connect(self.open_file)
        main_layout.addWidget(self.open_file_button)


        # List for items (show only Name)
        tests_label = QLabel("Тести")
        main_layout.addWidget(tests_label)
        self.listview = QListView(self)
        self.model = QStandardItemModel(self.listview)
        self.listview.setModel(self.model)
        main_layout.addWidget(self.listview)

        # Add, Edit, Remove Buttons
        button_layout = QHBoxLayout()

        self.add_button = QPushButton("Добавити")
        self.add_button.clicked.connect(self.add_item)
        button_layout.addWidget(self.add_button)

        self.edit_button = QPushButton("Редагувати")
        self.edit_button.clicked.connect(self.edit_item)
        button_layout.addWidget(self.edit_button)

        self.remove_button = QPushButton("Видалити")
        self.remove_button.clicked.connect(self.remove_item)
        button_layout.addWidget(self.remove_button)

        main_layout.addLayout(button_layout)

        # Radio buttons for time limitation
        self.date_button_group = QButtonGroup()
        
        self.radiobutton_all = QRadioButton("Усі данні")
        self.radiobutton_time = QRadioButton("Дані за період")
        self.radiobutton_all.setChecked(True)

        self.date_button_group.addButton(self.radiobutton_all, 1)
        self.date_button_group.addButton(self.radiobutton_time, 2)
        main_layout.addWidget(self.radiobutton_all)
        main_layout.addWidget(self.radiobutton_time)

        # Date selection
        self.date_picker_layout = QHBoxLayout()

        self.start_date = QDateEdit(self)
        self.end_date = QDateEdit(self)
        
        self.start_date.setCalendarPopup(True)
        self.end_date.setCalendarPopup(True)
        ukrainian_locale = QLocale(QLocale.Ukrainian)
        self.start_date.setLocale(ukrainian_locale)
        self.end_date.setLocale(ukrainian_locale)

        current_date = QDate.currentDate()

        # Set the first day of the current month        
        default_start_date = QDate(current_date.year(), current_date.month(), 1)
        default_end_date = current_date  

        self.start_date.setDate(default_start_date)
        self.end_date.setDate(default_end_date)

        self.date_picker_layout.addWidget(self.start_date)
        self.date_picker_layout.addWidget(self.end_date)
        self.start_date.setVisible(False)
        self.end_date.setVisible(False)
        
        self.radiobutton_all.toggled.connect(self.toggle_date_pickers)
        self.radiobutton_time.toggled.connect(self.toggle_date_pickers)


        main_layout.addLayout(self.date_picker_layout)

        # List for sheet names
        tests_label = QLabel("Листи з данними")
        main_layout.addWidget(tests_label)
        self.sheet_listview = QListView(self)
        self.sheet_model = QStandardItemModel(self.sheet_listview)
        self.sheet_listview.setModel(self.sheet_model)
        main_layout.addWidget(self.sheet_listview)

        # Save button
        self.save_button = QPushButton("Зберегти")
        self.save_button.clicked.connect(self.save_data)
        main_layout.addWidget(self.save_button)

        self.setLayout(main_layout)

    def toggle_date_pickers(self):
        if self.radiobutton_time.isChecked():
            self.start_date.setVisible(True)
            self.end_date.setVisible(True)
        else:
            self.start_date.setVisible(False)
            self.end_date.setVisible(False)

    
    def load_data(self):
        # Clear current list and load data from database
        self.model.clear()
        for row in self.cursor.execute("SELECT id, Name FROM TestStructures"):
            item = QStandardItem(row[1])
            item.setData(row[0], Qt.UserRole)
            self.model.appendRow(item)

    def open_file(self):
        # Let the user select an Excel file
        file_path, _ = QFileDialog.getOpenFileName(self, "Виберіть файл", "", "Excel files (*.xls *.xlsx)")
        if file_path:
            # Store the file path and update the label with the selected file's name
            self.selected_file_path = file_path
            file_name = file_path.split("/")[-1]
            self.selected_file_label.setText(f"Вибрано: {file_name}")

            # Clear the sheet_listview before adding any sheet names
            self.sheet_model.clear()
            self.open_excel_file(file_path)

    def open_excel_file(self, file_path):
        # Try opening with openpyxl (for .xlsx files)
        try:
            wb = openpyxl.load_workbook(file_path, keep_vba=False)
            sheet_names = wb.sheetnames
            self.populate_listview_with_sheet_names(sheet_names)
        except Exception as e:
            # If openpyxl fails, try xlrd for older .xls files
            try:
                wb = xlrd.open_workbook(file_path)
                sheet_names = wb.sheet_names()
                self.populate_listview_with_sheet_names(sheet_names)
            except Exception as e:
                QMessageBox.critical(self, "Помилка", f"Не вдалося відкрити файл: {e}")

    def populate_listview_with_sheet_names(self, sheet_names):
        # Populate sheet_listview with sheet names
        for sheet in sheet_names:
            item = QStandardItem(sheet)
            self.sheet_model.appendRow(item)

    def add_item(self):
        self.open_item_dialog()

    def edit_item(self):
        selected_item = self.listview.selectionModel().selectedIndexes()
        if selected_item:
            item_id = self.model.itemFromIndex(selected_item[0]).data(Qt.UserRole)
            self.cursor.execute("SELECT * FROM TestStructures WHERE id=?", (item_id,))
            item_data = self.cursor.fetchone()
            self.open_item_dialog(item_data)
        else:
            QMessageBox.warning(self, "Редагування", "Виберіть тест для редагування")


    def open_item_dialog(self, item_data=None):
        dialog = QDialog(self)
        dialog.setWindowTitle("Добавити / Редагувати")
        dialog.setFixedSize(300, 250)

        form_layout = QFormLayout(dialog)

        # Fields
        name_entry = QLineEdit(dialog)
        name_col_entry = QLineEdit(dialog)
        date_col_entry = QLineEdit(dialog)
        data_col_entry = QLineEdit(dialog)

        form_layout.addRow("Назва:", name_entry)
        form_layout.addRow("Колонка з ім'ям:", name_col_entry)
        form_layout.addRow("Колонка з датою:", date_col_entry)
        form_layout.addRow("Колонка з данними:", data_col_entry)

        # Pre-fill fields if editing
        if item_data:
            name_entry.setText(item_data[1])
            name_col_entry.setText(item_data[2])
            date_col_entry.setText(item_data[3])
            data_col_entry.setText(item_data[4])

        def save_item():
            name = name_entry.text()
            name_col = name_col_entry.text()
            date_col = date_col_entry.text()
            data_col = data_col_entry.text()

            if item_data:
                query = "UPDATE TestStructures SET Name=?, NameColumn=?, DateColumn=?, DataColumn=? WHERE id=?"
                self.cursor.execute(query, (name, name_col, date_col, data_col, item_data[0]))
            else:
                query = "INSERT INTO TestStructures (Name, NameColumn, DateColumn, DataColumn) VALUES (?, ?, ?, ?)"
                self.cursor.execute(query, (name, name_col, date_col, data_col))

            self.conn.commit()
            self.load_data()  # Reload data into the list
            dialog.accept()

        save_button = QPushButton("Save", dialog)
        save_button.clicked.connect(save_item)
        form_layout.addWidget(save_button)
        

        dialog.exec()

    def remove_item(self):
        selected_item = self.listview.selectionModel().selectedIndexes()
        if selected_item:
            item_id = self.model.itemFromIndex(selected_item[0]).data(Qt.UserRole)
            confirm = QMessageBox.question(self, "Підтвердити видалення", "Ви впевнені, що хочете видалити цей елемент?")
            if confirm == QMessageBox.Yes:
                self.cursor.execute("DELETE FROM TestStructures WHERE id=?", (item_id,))
                self.conn.commit()
                self.load_data()
        else:
            QMessageBox.warning(self, "Видалити елемент", "Будь ласка, виберіть елемент для видалення.")


    def save_data(self):
        selected_item = self.listview.selectionModel().selectedIndexes()
        
        if selected_item:
            item_id = self.model.itemFromIndex(selected_item[0]).data(Qt.UserRole)
            self.cursor.execute("SELECT * FROM TestStructures WHERE id=?", (item_id,))
            item_data = self.cursor.fetchone()

            selected_sheet = self.sheet_listview.selectionModel().selectedIndexes()
            if selected_sheet:
                sheet_name = self.sheet_listview.model().data(selected_sheet[0])
            else:
                QMessageBox.warning(self, "Виберіть лист", "Виберіть лист зі списку.")
                return
            
            default_filename = item_data[1] 
            save_path, _ = QFileDialog.getSaveFileName(self, "Зберегти файл", default_filename, "Word documents (*.docx)")
            
            if not save_path:
                return
            
            # Create structure for LasarusResults
            structure = {
                'name_column': item_data[2],  # NameColumn
                'date_column': item_data[3],  # DateColumn
                'data_columns': item_data[4]  # DataColumn
            }

            date_checked_id = self.date_button_group.checkedId()
            if date_checked_id == 1:
                lasarus_results = LasarusResults(self.selected_file_path)
            else:
                date_from = self.start_date.date()
                date_to = self.end_date.date()
                lasarus_results = LasarusResults(self.selected_file_path, date_from, date_to)    

            lasarus_results.save_results(sheet_name, save_path, structure)
            QMessageBox.information(self, "Збереження", "Дані успішно збережено.")
    
        else:
            QMessageBox.warning(self, "Зберегти дані", "Будь ласка, виберіть тест для збереження.")
    

    def closeEvent(self, event):
        self.conn.close()
        event.accept()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = App()
    window.show()
    sys.exit(app.exec())