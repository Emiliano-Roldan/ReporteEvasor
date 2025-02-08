import os
from datetime import datetime
import tkinter.messagebox as messagebox

class logger:
    def __init__(self):
        #self.log_folder = "C:\\nationalsoft\\FacturaPronta\\log"
        self.log_folder = "log"
        self.log_filename = f"{self.log_folder}\\log_{datetime.now().strftime('%Y-%m-%d')}.log"
        self.create_folder()
        self.create_log_file()

    def create_folder(self):
        if not os.path.exists(self.log_folder):
            os.makedirs(self.log_folder)

    def create_log_file(self):
        if not os.path.exists(self.log_filename):
            with open(self.log_filename, 'w') as file:
                file.write(f"Log creado el: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")

    def write_to_log(self, message):
        with open(self.log_filename, 'a') as file:
            file.write(f"=============================================================\n[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] {message}\n")