from PyQt5.QtWidgets import QApplication, QSpacerItem, QWidget, QVBoxLayout, QLabel, QComboBox, QPushButton, QTextEdit, QPlainTextEdit, QMainWindow, QLineEdit, QMenuBar
from PyQt5 import QtCore
import sys
import json
from docxtpl import DocxTemplate
from datetime import datetime as dt
import os

class LetterApp(QWidget):
    def init(self, parent=None):
        super().init(parent)

        self.user_data = None
        self.initUI()

    def initUI(self):
        
        layout = QVBoxLayout()
        self.setWindowTitle('Letter Generator')
        self.setMinimumSize(400, 420)
        self.setGeometry(400, 420, 400, 420)

        self.company_label = QLabel('Выберите компанию:')
        self.company_combo = QComboBox()
        self.company_combo.setEditable(True)
        self.company_combo.setObjectName("comboBox")
        self.company_combo.currentIndexChanged.connect(self.update_employee_combo)
        self.company_combo.setMinimumSize(0, 24)
        self.company_combo.setMaximumSize(400, 24)

        self.employee_label = QLabel('Выберите сотрудника:')
        self.employee_combo = QComboBox()
        self.employee_combo.setMinimumSize(0, 24)
        self.employee_combo.setMaximumSize(400, 24)

        self.theme_lable = QLabel("Тема письма")
        self.theme_text = QTextEdit()
        self.theme_text.setMinimumSize(QtCore.QSize(0, 24))
        self.theme_text.setMaximumSize(400, 48)        

        self.body_lable = QLabel("Тело письма")
        self.body_text = QPlainTextEdit()

        self.fio_lable = QLabel("Введите своё ФИО")
        self.fio_text = QPlainTextEdit()
        self.fio_text.setMinimumSize(0, 24)
        self.fio_text.setMaximumSize(400, 24)

        self.position_lable = QLabel("Введите свою должность")
        self.position_text = QPlainTextEdit()
        self.position_text.setMinimumSize(QtCore.QSize(0, 24))
        self.position_text.setMaximumSize(QtCore.QSize(400, 24))

        self.generate_button = QPushButton("Сформировать", self)
        self.generate_button.clicked.connect(self.generate_letter)

        self.update_button = QPushButton('Обновить', self)
        self.update_button.clicked.connect(self.updateData)

        layout.addWidget(self.company_label)
        layout.addWidget(self.company_combo)
        layout.addWidget(self.employee_label)
        layout.addWidget(self.employee_combo)
        layout.addWidget(self.fio_lable)
        layout.addWidget(self.fio_text)
        layout.addWidget(self.position_lable)
        layout.addWidget(self.position_text)
        layout.addWidget(self.theme_lable)
        layout.addWidget(self.theme_text)
        layout.addWidget(self.body_lable)
        layout.addWidget(self.body_text)
        layout.addWidget(self.update_button)
        layout.addWidget(self.generate_button)

        layout.addSpacerItem(QSpacerItem(0, 5))
        self.setLayout(layout)

    def set_user_data(self, user_data):
        self.user_data = user_data
        self.company_combo.addItems([company['company_name'] for company in self.user_data['companies']])
        

    def update_employee_combo(self):
        self.employee_combo.clear()

        if self.user_data is not None:
            selected_company_index = self.company_combo.currentIndex()
            employees = self.user_data['companies'][selected_company_index]['employees']

            self.employee_combo.addItems([f"{employee['last_name']} {employee['first_name']} {employee['middle_name']}" for employee in employees])

    def updateData(self):
        
        with open(os.path.join(os.getcwd(), 'user_info.json'), 'r', encoding='utf-8') as file:
            user_data = json.load(file)
        
        self.company_combo.clear()
        self.employee_combo.clear()
        self.set_user_data(user_data)
        editor = EditDatabase(user_data)
        editor.data_saved.connect(self.updateData)
     
    def generate_letter(self):
        doc = DocxTemplate('template3.docx')
        data_base = json.load(open(os.path.join(os.getcwd(), 'user_info.json'), encoding='utf-8'))
        if not os.path.isdir(os.path.join(os.getcwd(), 'Письма')):
            os.mkdir(os.path.join(os.getcwd(), 'Письма'))

        max_file_number = 0
        # Create a valid file path by joining the directory and file name
        for file in os.listdir(os.path.abspath(r"Письма")):
            if os.path.isfile(os.path.join(os.getcwd(), "Письма", file)):
                # Изолируем имя файла и его расширение
                name, extension = os.path.splitext(file)
                # Убедимся, что имя начинается с числа
                if name[:3].isdigit():
                    # Преобразуем первые три символа в номер
                    file_number = int(name[:3])
                    # Обновим максимальный номер
                    if file_number > max_file_number:
                        max_file_number = file_number
                        letter_number = max_file_number + 1

        print(letter_number)
        print(max_file_number)
        
            # Get the selected company and employee
        selected_company_index = self.company_combo.currentIndex()
        selected_employee_index = self.employee_combo.currentIndex()

        selected_employee = self.user_data['companies'][selected_company_index]['employees'][selected_employee_index]
        rec_email = selected_employee['email']
        rec_position = selected_employee['position']
        rec_io = f"{selected_employee['first_name']} {selected_employee['last_name']}"



        data = {
            "letter_number":str(letter_number).zfill(3),
            "ans_letter_number":(),
            'year': dt.strftime(dt.now(), '%y'),
            "theme":self.theme_text.toPlainText(),
            "company_name": self.company_combo.currentText(),
            "recipient_fio": self.employee_combo.currentText(),
            "recipient_position":rec_position,
            "recipient_email":rec_email,
            "recipient_io":rec_io,
            "body":self.body_text.toPlainText(),
            "position":self.position_text.toPlainText(),
            "fio":self.fio_text.toPlainText()
        }

        doc.render(data)
       
        file_path = os.path.join(os.getcwd(), "Письма", f'{str(letter_number).zfill(3)} {self.company_combo.currentText()} {self.employee_combo.currentText()}.docx')

        doc.save(file_path)

        print("письмо готово")

class EditDatabase(QMainWindow):
    data_saved = QtCore.pyqtSignal()

    def __init__(self, user_data, parent=None):
        super().__init__(parent)

        self.user_data = user_data
        self.initUI()

    def initUI(self):
        self.setWindowTitle('DB Editor')
        self.setGeometry(800, 120, 400, 300)
        self.setMaximumSize(350,120)

        self.company_name_label = QLabel('Название компании:')
        # self.company_name_field = QLineEdit()
        self.company_name_field = QComboBox()
        self.company_name_field.setEditable(True)
        self.add_company_button = QPushButton('Добавить компанию в Базу', self)
        self.add_company_button.clicked.connect(self.addCompany)

        self.employee_label = QLabel('Данные сотрудника:')
        self.first_name_field = QLineEdit()
        self.middle_name_field = QLineEdit()
        self.last_name_field = QLineEdit()
        self.email_field = QLineEdit()
        self.position_field = QLineEdit()
        self.add_employee_button = QPushButton('Добавить сотрудника', self)
        self.add_employee_button.clicked.connect(self.addEmployee)

        self.save_button = QPushButton('Сохранить измененния', self)
        self.save_button.clicked.connect(self.saveDatabase)

        layout = QVBoxLayout()
        layout.addWidget(self.company_name_label)
        layout.addWidget(self.company_name_field)
        layout.addWidget(self.add_company_button)

        layout.addWidget(self.employee_label)
        layout.addWidget(QLabel('Фамилия:'))
        layout.addWidget(self.last_name_field)
        layout.addWidget(QLabel('Имя:'))
        layout.addWidget(self.first_name_field)
        layout.addWidget(QLabel('Отчество:'))
        layout.addWidget(self.middle_name_field)
        layout.addWidget(QLabel('Email:'))
        layout.addWidget(self.email_field)
        layout.addWidget(QLabel('Должность:'))
        layout.addWidget(self.position_field)
        layout.addWidget(self.add_employee_button)

        layout.addWidget(self.save_button)

        central_widget = QWidget()
        central_widget.setLayout(layout)
        self.setCentralWidget(central_widget)

        with open(os.path.join(os.getcwd(), 'user_info.json'), 'r', encoding='utf-8') as file:
            user_data = json.load(file)
        for company in user_data['companies']:
            self.company_name_field.addItem(company['company_name'])

    def addCompany(self):
        company_name = self.company_name_field.currentText()
        if company_name:
            self.user_data['companies'].append({
                "company_name": company_name,
                "employees": []
            })
            print(f"Added company: {company_name}")
        else:
            print("Company name cannot be empty.")

    def addEmployee(self):
        company_name = self.company_name_field.currentText()
        employee_data = {
            "first_name": self.first_name_field.text(),
            "middle_name": self.middle_name_field.text(),
            "last_name": self.last_name_field.text(),
            "email": self.email_field.text(),
            "position": self.position_field.text()
        }

        for company in self.user_data['companies']:
            if company['company_name'] == company_name:
                company['employees'].append(employee_data)
                print(f"Added employee: {employee_data}")
                break
        else:
            print(f"Company '{company_name}' not found.")
    
    def delCompany(self):
        pass

    def delEmployee(self):
        pass

    def saveDatabase(self):
        # Save the changes to the JSON file
        with open("user_info.json", "w", encoding='utf-8') as file:
            json.dump(self.user_data, file, ensure_ascii=False, indent=4)
        self.data_saved.emit()
        print("Changes saved to the database.")

def main():
    app = QApplication(sys.argv)

    # Загружаем данные из файла JSON
    with open(os.path.join(os.getcwd(), 'user_info.json'), 'r', encoding='utf-8') as file:
        user_data = json.load(file)


    window = LetterApp()
    window.initUI()
    window.set_user_data(user_data)
    window.show()
    
    editor = EditDatabase(user_data)
    editor.show()

    sys.exit(app.exec_())

if __name__ == '__main__':
    main()
