from tkinter import *
from tkinter.ttk import *
from tkinter import filedialog
from tkinter import messagebox
import get_excel_list
import json 
import os

#Функция для изменения конфига. key - ключ в конфиге по которому нужно изменить value - значение в конфиге
def edit_config(key, value):
    with open("assets/config.json", 'r+') as f:
        config = json.load(f)
        config[key] = value
        f.seek(0)
        json.dump(config, f, indent=4)
        f.truncate()


def create_folder():
    if not os.path.exists("result/"):
        os.makedirs("result/")

#Функция, возвращающая конфиг
def load_config():
    with open("assets/config.json", 'r') as f:
        config = json.load(f)
        return config

#Функция, обновляющая строки имён файлов в конфиге (просто обновляет на актуальные)
def update_filenames():
    config = load_config()
    try:
        with open(config['filepath_xls']) as _:
            pass
        filename_xls = str(config['filepath_xls'][((config['filepath_xls'].rfind('/'))+1):-1]+config['filepath_xls'][-1])
        edit_config("filename_xls", filename_xls)
    except:
        edit_config("filename_xls", "Файл отсутствует")

    try:
        with open(config['filepath_pptx']) as _:
            pass
        filename_pptx = str(config['filepath_pptx'][((config['filepath_pptx'].rfind('/'))+1):-1]+config['filepath_pptx'][-1])
        edit_config("filename_pptx", filename_pptx)
    except:
        edit_config("filename_pptx", "Файл отсутствует")    

#Функция, создающая конфиг при его отсутствии
def create_json():
    with open('assets/config.json', 'w') as f:
        data = {
                "filename_xls": "\u0424\u0430\u0439\u043b \u043e\u0442\u0441\u0443\u0442\u0441\u0442\u0432\u0443\u0435\u0442",
                "filename_pptx": "\u0424\u0430\u0439\u043b \u043e\u0442\u0441\u0443\u0442\u0441\u0442\u0432\u0443\u0435\u0442",
                "submitted": False,
                "button_slide_enabled": False,
                "button_class_enabled": False,
                "slides": 0
                }
        json.dump(data, f)

#Функция, записывающая значение выбранного класса
def enablebutton_class(val):
    if val != "Не выбрано":
        edit_config("button_class_enabled", True)
        edit_config("class", val)

#Функция, записывающая значение выбранного класса
def enablebutton_slide(val):
    if val != "Не выбрано":
        edit_config("button_slide_enabled", True)
        edit_config("slide", val)

#Функция кнопки подтверждения. Закрывает окно при нажатии и сохраняет данные
def submit(window):
    config = load_config()
    if config["button_slide_enabled"] and config["button_class_enabled"]:
        edit_config("submitted", True)
        window.destroy()
    else:
        messagebox.showerror("Ошибка", "Вы не выбрали класс/слайд")

#Функция кнопки отметы. Закрывает окно при нажатии
def cancel(window):
    window.destroy()

#Функция для отключения меню классов
def switch_dropdown(var, option_menu):
    option_menu.configure(state=var)
    if var == "disabled":
        edit_config("all", True)
        edit_config("button_class_enabled", True)
    else:
        edit_config("all", False)
        edit_config("button_class_enabled", False)


def show_error(header, text):
    messagebox.showerror(header, text)


