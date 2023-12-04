from PyQt5.QtWidgets import QApplication, QSpacerItem, QWidget, QVBoxLayout, QLabel, QComboBox, QPushButton, QTextEdit, QPlainTextEdit
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
        self.setGeometry(300, 300, 300, 200)

        self.company_label = QLabel('Выберите компанию:')
        self.company_combo = QComboBox()
        self.company_combo.currentIndexChanged.connect(self.update_employee_combo)

        self.employee_label = QLabel('Выберите сотрудника:')
        self.employee_combo = QComboBox()

        self.theme_lable = QLabel("Тема письма")
        self.theme_text = QTextEdit()
        self.theme_text.setMinimumSize(QtCore.QSize(300, 24))
        self.theme_text.setMaximumSize(350, 48)        

        self.body_lable = QLabel("Тело письма")
        self.body_text = QTextEdit()

        self.fio_lable = QLabel("Введите своё ФИО")
        self.fio_text = QPlainTextEdit()
        self.fio_text.setMinimumSize(QtCore.QSize(0, 24))
        self.fio_text.setMaximumWidth(250)

        self.position_lable = QLabel("Введите свою должность")
        self.position_text = QPlainTextEdit()
        self.position_text.setMinimumSize(QtCore.QSize(0, 24))
        self.position_text.setMaximumSize(QtCore.QSize(250, 24))

        self.generate_button = QPushButton('Generate Letter', self)
        self.generate_button.clicked.connect(self.generate_letter)

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
            

    def generate_letter(self):
        doc = DocxTemplate('template3.docx')
        data_base = json.load(open(os.path.join(os.getcwd(), 'user_info.json'), encoding='utf-8'))
        if not os.path.isdir(os.path.join(os.getcwd(), 'Письма')):
            os.mkdir(os.path.join(os.getcwd(), 'Письма'))

        max_file_number = 0
        # Create a valid file path by joining the directory and file name
        for file in os.listdir(os.path.abspath("D:\Projects\letters\Письма")):
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

def main():
    app = QApplication(sys.argv)

    # Загружаем данные из файла JSON
    with open(os.path.join(os.getcwd(), 'user_info.json'), 'r', encoding='utf-8') as file:
        user_data = json.load(file)


    window = LetterApp()
    window.initUI()
    window.set_user_data(user_data)
    window.show()
    
    sys.exit(app.exec_())

if __name__ == '__main__':
    main()
