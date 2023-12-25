from PyQt5.QtWidgets import QMainWindow, QVBoxLayout, QHBoxLayout, QWidget, QLabel, QComboBox, QPushButton, QLineEdit, QMessageBox
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
        self.del_comp_button = QPushButton('Удалить компанию', self)
        self.del_comp_button.clicked.connect(self.delCompany)

        self.employee_label = QLabel('Данные сотрудника:')
        self.first_name_field = QComboBox()
        self.first_name_field.setEditable(True)
        self.middle_name_field = QLineEdit()
        self.last_name_field = QComboBox()
        self.last_name_field.setEditable(True)
        self.last_name_field.setObjectName("comboBox")
        self.last_name_field.currentIndexChanged.connect(self.update_last_name_combo)        
        self.email_field = QLineEdit()
        self.position_field = QLineEdit()
        self.add_employee_button = QPushButton('Добавить сотрудника', self)
        self.add_employee_button.clicked.connect(self.addEmployee)

        self.save_button = QPushButton('Сохранить измененния', self)
        self.save_button.clicked.connect(self.saveDatabase)
        self.del_empl_button = QPushButton('Удалить сотрудника', self)
        self.del_empl_button.clicked.connect(self.delEmployee)
        self.close_button = QPushButton("Закрыть", self)
        self.close_button.clicked.connect(self.close)


        layout = QVBoxLayout()
        Hlayout = QHBoxLayout()
        H1layout = QHBoxLayout()
        H2layout = QHBoxLayout()
        layout.addWidget(self.company_name_label)
        layout.addWidget(self.company_name_field)
        Hlayout.addWidget(self.del_comp_button)
        Hlayout.addWidget(self.add_company_button)
        layout.addLayout(Hlayout)

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
        H1layout.addWidget(self.del_empl_button)
        H1layout.addWidget(self.add_employee_button)
        layout.addLayout(H1layout)

        H2layout.addWidget(self.save_button)
        H2layout.addWidget(self.close_button)
        layout.addLayout(H2layout)

        central_widget = QWidget()
        central_widget.setLayout(layout)
        self.setCentralWidget(central_widget)

        with open(os.path.join(os.getcwd(), 'user_info.json'), 'r', encoding='utf-8') as file:
            user_data = json.load(file)
        for company in user_data['companies']:
            self.company_name_field.addItem(company['company_name'])

    def update_last_name_combo(self):
        self.last_name_field.clear()

        if self.user_data is not None:
            selected_company_index = self.company_name_field.currentIndex()
            employees = self.user_data['companies'][selected_company_index]['employees']
            
            self.last_name_field.addItems([f"{employee['last_name']}" for employee in employees])

    def addCompany(self):
        company_name = self.company_name_field.currentText()
        if company_name:
            existing_companies = [company['company_name'] for company in self.user_data['companies']]
            if company_name in existing_companies:
                QMessageBox.information(self, "Информация", f"Компания с именем '{company_name}' уже существует в базе данных.")
            else:
                self.user_data['companies'].append({
                    "company_name": company_name,
                    "employees": []
                })
                QMessageBox.information(self, "Успех", f"Добавлена компания: {company_name}")
        else:
            QMessageBox.critical(self, "Ошибка", "Название компании не может быть пустым.")



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
                full_name = f"{employee_data.get('last_name', '')} {employee_data.get('first_name', '')} {employee_data.get('last_name', '')}"
                QMessageBox.information(self, "Информация", f"Сотрудник {full_name} добавлен ")
                break
        else:
            QMessageBox.information(self, "Информация", f"Компания {company_name} не найдена.")
    
    def delCompany(self):
        company_name = self.company_name_field.currentText()
        for company in self.user_data['companies']:
            if company['company_name'] == company_name:
                self.user_data['companies'].remove(company)
                self.company_name_field.removeItem(self.company_name_field.currentIndex())
                QMessageBox.information(self, "Информация", f"Компания {company_name} удалена из базы данных.")
                break
        else:
            QMessageBox.information(self, "Информация", f"Компания {company_name} не найдена.")

    def delEmployee(self):
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
                for employee in company['employees']:
                    if all(employee[key] == value for key, value in employee_data.items()):
                        company['employees'].remove(employee)
                        QMessageBox.information(self, "Информация", f"Сотрудник {employee_data['last_name']} удален")
                        break
                else:
                    QMessageBox.information(self, "Информация", "Сотрудник не найден.")
                break
        else:
            QMessageBox.information(self, "Информация", f"Компания {company_name} не найдена.")


    def saveDatabase(self):
        # Save the changes to the JSON file
        with open("user_info.json", "w", encoding='utf-8') as file:
            json.dump(self.user_data, file, ensure_ascii=False, indent=4)
        self.data_saved.emit()

        QMessageBox.information(self, "Информация", "Changes saved to the database.")
    
    def close(self):
        self.hide()
