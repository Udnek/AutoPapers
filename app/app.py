from tkinter import *
from tkinter.ttk import *
from tkinter import messagebox
from get_excel_list import fetch_list_classes
from libs_gleb import slides_amount
from libs_dima import *


#Функция, записывающая системный путь к выбранному файлу списка участников
def new_file_xls(text_to_update, window):
    filepath = filedialog.askopenfilename()
    if str(filepath[-4:]) != ".xls":
        messagebox.showerror("Ошибка", "Выбран не верный файл.\nФайл должен иметь расширение .xls\n(Microsoft Excel 1997-2003)")
    else:
        edit_config("filepath_xls", filepath)
        update_filenames()
        config = load_config()
        text_to_update.configure(text="{}".format(config["filename_xls"]))
        window.destroy()
        interface_badges()

#Функция, записывающая системный путь к выбранному файлу шаблона
def new_file_pptx(text_to_update, window):
    filepath = filedialog.askopenfilename()
    if str(filepath[-5:]) != ".pptx":
        messagebox.showerror("Ошибка", "Выбран не верный файл.\nФайл должен иметь расширение .pptx\n(Microsoft Power Point 2007)")
    else:
        edit_config("filepath_pptx", filepath)
        update_filenames()
        config = load_config()
        text_to_update.configure(text="{}".format(config["filename_pptx"]))
        window.destroy()
        interface_badges()

#Интерфейс для заполнения бейджиков
def interface_badges():
    try:
        config = load_config()
    except:
        create_json()
        config = load_config()
    create_folder()
    update_filenames()
    config = load_config()
    edit_config("submitted", False)
    edit_config("all", False)
    edit_config("button_slide_enabled", False)
    edit_config("button_class_enabled", False)

    #Блок отвечающий за основное окно программы
    window = Tk()
    window.title("Auto Papers")
    try:
        window.iconbitmap("assets/icon.ico")
    except:
        pass
    screen_width = window.winfo_screenwidth()
    screen_height = window.winfo_screenheight()
    width = 320
    height = 300
    x = (screen_width/2) - (width/2)
    y = (screen_height/2) - (height/2)
    window.geometry('%dx%d+%d+%d' % (width, height, x, y))
    window.resizable(False, False)

    #Блок отвечающий за стили
    button_style = Style()
    button_style.configure("TButton")

    #Блок, отвечающий за окно с файлом таблицы
    frame_addfile_xls = LabelFrame(text='Файл списка участников [формат .xls]')
    frame_addfile_xls.place(width=300, height=78, x=10, y=4)
    addfile_button_xls = Button(frame_addfile_xls, style="TButton", text="Изменить файл", command=lambda:new_file_xls(current_file_xls, window), cursor="hand2")
    addfile_button_xls.place(width=120, height=51, x=173, y=2)
    current_file_xls = Label(frame_addfile_xls, text="{}".format(config["filename_xls"]), relief="groove", anchor=CENTER)
    current_file_xls.place(x=3, y=3, width=170, height=50)
    
    #Блок, отвечающий за окно с файлом шаблона
    frame_addfile_pptx = LabelFrame(text='Файл шаблонов [формат .pptx]')
    frame_addfile_pptx.place(width=300, height=78, x=10, y=82)
    addfile_button_pptx = Button(frame_addfile_pptx, style="TButton", text="Изменить файл", command=lambda:new_file_pptx(current_file_pptx, window), cursor="hand2")
    addfile_button_pptx.place(width=120, height=51, x=173, y=2)
    current_file_pptx = Label(frame_addfile_pptx, text="{}".format(config["filename_pptx"]), relief="groove", anchor=CENTER)
    current_file_pptx.place(x=3, y=3, width=170, height=50)

    #Блок, отвечающий за окно с выбором класса
    try:
        classes_raw = fetch_list_classes(config['filepath_xls'])
    except:
        classes_raw = [["Нет данных"]]
    classes = ["Не выбрано"]
    for i in classes_raw:
        classes.append(i)
    frame_grade = LabelFrame(text='Класс')
    frame_grade.place(width=150, height=45, x=10, y=160)
    variable_class = StringVar()
    variable_class.set(classes[0])
    choose_grade = OptionMenu(frame_grade, variable_class,  command=enablebutton_class, *classes)
    choose_grade.place(width=135, x=5)

    #Блок, отвечающий за чекбокс для дипломов
    frame_diplomes = LabelFrame(text="Для дипломов")
    frame_diplomes.place(width=150, height=45, x=160, y=160)
    var_diplomes = StringVar()
    diplomes = Checkbutton(frame_diplomes, text="Все классы", variable=var_diplomes, cursor="hand2", 
                           onvalue="disabled", offvalue="normal", command=lambda:switch_dropdown(var_diplomes.get(), choose_grade))
    diplomes.place(width=135, x=5, y=1)

    #Блок, отвечающий за окно с выбором слайда
    slidescount = slides_amount()
    slides = ["Не выбрано"]
    if slidescount != 0:  
        for i in range(1, slidescount+1):
            slides.append(i)
    else:
        slides.append("{Нет данных}")
    
    frame_slide = LabelFrame(text='Номер слайда с нужным шаблоном')
    frame_slide.place(width=300, height=50, x=10, y=210)
    variable_slide = StringVar()
    variable_slide.set(classes[0])
    choose_slide = OptionMenu(frame_slide, variable_slide,  command=enablebutton_slide, *slides)
    choose_slide.place(width=285, x=5, y=3)

    #Блок, отвечающий за кнопку подтверждения настроек
    submit_button = Button(window, text="Готово", command=lambda:submit(window), cursor="hand2")
    submit_button.place(width=100, height=25, x=211, y=266)

    #Блок, отвечающий за кнопку отмены
    cancel_button = Button(window, text="Отмена", command=lambda:cancel(window), cursor="hand2")
    cancel_button.place(width=100, height=25, x=110, y=266)

    window.mainloop()

"""#Интерфейс для заполнения дипломов
def interface_diplomes():
    pass

#Интерфейс для выбора программы
def interface():
    update_filenames()
    edit_config("submitted", False)
    edit_config("button_slide_enabled", False)
    edit_config("button_class_enabled", False)
    
    window_choose = Tk()
    window_choose.title("Auto Papers")
    window_choose.iconbitmap("images\icon.ico")
    screen_width = window_choose.winfo_screenwidth()
    screen_height = window_choose.winfo_screenheight()
    width = 420
    height = 120
    x = (screen_width/2) - (width/2)
    y = (screen_height/2) - (height/2)
    window_choose.geometry('%dx%d+%d+%d' % (width, height, x, y))
    window_choose.resizable(False, False)

    Style().configure("TButton", cursor="hand2")

    frame_main = LabelFrame(window_choose, text="Выберите программу")
    frame_main.place(width=400, height=112, x=10, y=3)

    choose_autobadges = Button(frame_main, text="Авто-бейджики", command=lambda:chosen_badges(window_choose))
    choose_autobadges.place(width=190, height=90, x=5, y=0)
    choose_autodiplomes = Button(frame_main, text="Авто-дипломы",command=lambda:messagebox.showerror("Не доступно", "Функция пока не работает"))
    choose_autodiplomes.place(width=190, height=90, x=202, y=0)

    window_choose.mainloop()"""

if __name__ == "__main__":
    interface_badges()