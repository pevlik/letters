from PyQt5.QtWidgets import QVBoxLayout, QWidget, QLabel, QComboBox, QTextEdit, QPlainTextEdit, QPushButton, QSpacerItem
from PyQt5 import QtCore
from docxtpl import DocxTemplate
from datetime import datetime as dt
from edit_database import EditDatabase
import json
import os

class LetterApp(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        
        self.user_data = None
        self.initUI()

        self.editor = EditDatabase(self.user_data)

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

        self.edit_database_button = QPushButton("Редактировать базу", self)
        self.edit_database_button.clicked.connect(self.open_editor)


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
        layout.addWidget(self.edit_database_button)

        layout.addSpacerItem(QSpacerItem(0, 5))
        self.setLayout(layout)

    def open_editor(self):
        self.editor = EditDatabase(self.user_data)
        self.editor.show()  

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
     
    def get_user_data(self):
        with open(('user_info.json'), encoding='utf-8') as file:
            return json.load(file)

    def generate_letter(self):
        doc = DocxTemplate('template3.docx')
        if not os.path.isdir(os.path.join(os.getcwd(), 'Письма')):
            os.mkdir(os.path.join(os.getcwd(), 'Письма'))
        
        file_number = 000
        letter_number = "001"
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
