from PyQt5.QtWidgets import QMainWindow, QVBoxLayout, QWidget, QLabel, QComboBox, QPushButton, QLineEdit
from PyQt5.QtCore import pyqtSignal
import os
import json
import os

class EditDatabase(QMainWindow):
    data_saved = pyqtSignal()

    def __init__(self, user_data):
        super().__init__()

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
        self.close_button = QPushButton("Закрыть", self)
        self.close_button.clicked.connect(self.close)


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
        layout.addWidget(self.close_button)

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
    
    def close(self):
        self.hide()

