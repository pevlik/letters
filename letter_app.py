from PyQt5.QtWidgets import QVBoxLayout,  QHBoxLayout, QWidget, QLabel, QComboBox, QTextEdit, QPlainTextEdit, QPushButton, QSpacerItem, QMessageBox, QMenuBar, QAction, QFileDialog
from PyQt5 import QtCore
from docxtpl import DocxTemplate
from datetime import datetime as dt
from edit_database import EditDatabase
import pandas as pd
import json
import os

class LetterApp(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        
        self.user_data = None
        self.initUI()

        self.editor = EditDatabase(self.user_data)

        self.setupMenuBar()

    def setupMenuBar(self):
        mainMenu = QMenuBar()
        editMenu = mainMenu.addMenu('Настройки')
        editDatabaseAction = QAction('Редактировать базу данных', self)
        editDatabaseAction.triggered.connect(self.open_editor)
        addDatabaseAction = QAction('Загрузить базу данных', self)
        addDatabaseAction.triggered.connect(self.upload_new_base)
        editMenu.addAction(editDatabaseAction)
        editMenu.addAction(addDatabaseAction)
        self.layout().setMenuBar(mainMenu)

    def initUI(self):
        
        layout = QVBoxLayout()
        Hlayout =  QHBoxLayout()
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
        self.body_text.setMinimumSize(QtCore.QSize(0, 24))

        self.fio_lable = QLabel("Введите своё ФИО")
        self.fio_text = QPlainTextEdit()
        self.fio_text.setPlaceholderText("flds")
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

        self.close_button = QPushButton("Закрыть", self)
        self.close_button.clicked.connect(self.close)

        # self.edit_database_button = QPushButton("Редактировать базу", self)
        # self.edit_database_button.clicked.connect(self.open_editor)

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
        Hlayout.addWidget(self.update_button)
        Hlayout.addWidget(self.generate_button)
        Hlayout.addWidget(self.close_button)
        # layout.addWidget(self.edit_database_button)
        layout.addLayout(Hlayout)         

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

    def upload_new_base(self):
        file_path, _ = QFileDialog.getOpenFileName(None, "выберите файл xlsx", "", "XLSX files (*.xlsx)")
        # Загрузка данных из файла xlsx
        df = pd.read_excel(file_path)
        df = df.rename(columns={"Эл. адрес": "email", "Должность": "position", "Организация":"company_name"})
        df['email'] = df['email'].fillna('')
        df[['last_name', 'first_name', 'middle_name']] = df['ФИО'].str.split(' ', expand=True)

        # Чтение существующего файла json
        with open('user_info.json', 'r', encoding='utf-8') as file:
            existing_data = json.load(file)

        # Проверка наличия дубликатов перед добавлением новых данных
        new_json_data = {
            "companies": df.groupby('company_name').apply(lambda x: x[['first_name','middle_name','last_name','email', 'position']].to_dict('records')).reset_index().rename(columns={0:'employees'}).to_dict(orient='records')
        }

        existing_companies = {
            company['company_name']: {
            'employees': company['employees']
            } 
            for company in existing_data['companies']
        }

        for new_company in new_json_data['companies']:
            company_name = new_company['company_name']

            if company_name in existing_companies:
            # Компания уже есть
                existing_company = existing_companies[company_name]

            for employee in new_company['employees']:
                # Проверяем каждого сотрудника 
                exists = any(all(d == e for d, e in employee.items()) 
                            for d in existing_company['employees'])

                if not exists:
                # Сотрудник не дублируется, добавляем
                    existing_company['employees'].append(employee)

            else:
            # Новая компания, добавляем полностью  
                existing_data["companies"].append(new_company)

        # Запись обновленных данных
        with open('user_info.json', 'w', encoding='utf-8') as file:
            json.dump(existing_data, file, ensure_ascii=False, indent=4)
            
        self.updateData()
        # # Проверка наличия дубликатов
        # existing_companies = [company['company_name'] for company in existing_data['companies']]  
        # new_companies = [company['company_name'] for company in new_json_data['companies']]

        # duplicate_companies = set(existing_companies) & set(new_companies)

        # if duplicate_companies:
        #     print("Обнаружены дубликаты")
        #     # Обновление данных без дубликатов
        #     for new_company in new_json_data['companies']:
        #         if new_company['company_name'] not in existing_companies:
        #             existing_data["companies"].append(new_company)
        # else:
        #     # Обновление данных без дубликатов
        #     for new_company in new_json_data['companies']:
        #         if new_company['company_name'] not in existing_companies:
        #             existing_data["companies"].append(new_company)
                    
        #     print("Дубликаты не обнаружены, обновлены данные")
        
        # # Запись обновленных данных   
        # with open('user_info.json', 'w', encoding='utf-8') as file:
        #     json.dump(existing_data, file, ensure_ascii=False, indent=4)
        
        # self.updateData()

    def get_user_data(self):
        with open(('user_info.json'), encoding='utf-8') as file:
            return json.load(file)

    def generate_letter(self):
        doc = DocxTemplate('template4.docx')
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
        
            # Get the selected company and employee
        selected_company_index = self.company_combo.currentIndex()
        selected_employee_index = self.employee_combo.currentIndex()

        selected_employee = self.user_data['companies'][selected_company_index]['employees'][selected_employee_index]
        rec_email = selected_employee['email']
        rec_position = selected_employee['position']
        rec_io = f"{selected_employee['first_name']} {selected_employee['last_name']}"



        data = {
            "let_num":str(letter_number).zfill(3), # Номер текущего письма
            "ans_let_num":(), # Номер письма на которое отвечают
            'year': dt.strftime(dt.now(), '%y'), # Год отправки письма
            'date': dt.strftime(dt.now(), '%d.%m.%Y'), # Дата отправки письма
            "theme":self.theme_text.toPlainText(), # Тема письма
            "company_name": self.company_combo.currentText(), # Название компании получателя
            "rec_fio": self.employee_combo.currentText(), # Фамилия Имя Отчество получателся
            "rec_position":rec_position, # Должность получателя
            "rec_email":rec_email, # Электронная почта получателя
            # "rec_fax":rec_fax, # Факс получателя
            # "rec_phone":rec_phone, # Телефон получателя
            # "rec_enother":rec_enother, # Другое
            "rec_io":rec_io, # Фамилия Имя получателя
            "body":self.body_text.toPlainText(), # Текст письма
            "position":self.position_text.toPlainText(), # Должность отправителя
            "fio":self.fio_text.toPlainText(), # Фамилия Имя Отчество отправителя Щербаков Илья Владимирович
            # "o_mgr":o_manager, # Фамилия Имя Отчество исполнителя
            # "o_mgr_phone": o_manager_phone # Телефон исполнителя
        }

        doc.render(data)
       
        yeardata = dt.strftime(dt.now(), '%y' )

        file_path = os.path.join(os.getcwd(), "Письма", f'{str(letter_number).zfill(3)}-{yeardata} {self.company_combo.currentText()} {self.employee_combo.currentText()}.docx')

        doc.save(file_path)

        QMessageBox.information(self,"Информация", "письмо готово")
