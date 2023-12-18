import sys
import json
import os
from PyQt5.QtWidgets import QApplication
from letter_app import LetterApp


def main():

  app = QApplication(sys.argv)

  # Загружаем данные из файла JSON
  with open(os.path.join(os.getcwd(), 'user_info.json'), 'r', encoding='utf-8') as file:
      user_data = json.load(file)

  letter_app = LetterApp()
  letter_app.show()
  letter_app.set_user_data(user_data)

  sys.exit(app.exec_())

if __name__ == '__main__':
  main()
