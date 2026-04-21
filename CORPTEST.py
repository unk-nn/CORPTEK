import tkinter as tk
from tkinter import messagebox as msg
from TestUtils import Utils
import threading, queue, pythoncom
from queue import Queue
import os, sys

#Возможно переработать GUI result в таблицу

# Этот класс отвечает за наполнение окна элементами,
# а также содержит всю логику кнопок и экрана загрузки.
# В нём есть методы для быстрого добавления готовых элементов на экран.
class GUI_Widget(tk.Frame):
    def __init__(self, master, version:str, size_title:int, size_sub_title:int, size_text:int, size_button:int):
        # Передаем в родительский класс 'окно' "self.__root"
        super().__init__(master)
        self.pack()

        # Устанавливаем версию приложения
        self.version = version

        # Устанавливаем размеры для всего контента на экране
        self.__size_title = size_title
        self.__size_sub_title = size_sub_title
        self.__size_text = size_text
        self.__size_btn = size_button

        #Кнопки главного экрана
        self.__size_start_test_btn = 13
        self.__size_info_and_settings_btn = 31


        #Параметры настроек

        #Почта
        self.__mail_node = 'imap.yandex.ru'
        self.__imap_port = 993
        self.__smtp_port = 143

        #Корп узлы
        self.__trust_node1 = 'yandex.cloud'
        self.__trust_node2 = 'direct.yandex.ru'

        #Результат тестирования
        self.__result_test = []
    # Метод для очистки всех элементов на экране
    def clear_window(self):
        for widget in self.master.winfo_children():
            widget.destroy()
        return True

    # Метод наполнения элементами главного экрана
    def main_page(self):
        self.clear_window()

        self.add_image(os.path.join(sys._MEIPASS, 'name14.png'), 215, 1080)

        self.add_button('НАЧАТЬ ТЕСТИРОВАНИЕ', height=5, width=30, fg='white',
                                 bg='#004C99', font=('Arial', self.__size_start_test_btn, 'bold italic'),
                        command=self.diagnostics)
        self.add_button('⚙', height=1, width=5, fg='#004C99', bg='white',
                        font=('Segoe UI Symbol', self.__size_info_and_settings_btn), side="right", x=10, y=1,
                        command=self.__settings)
        self.add_button('ⓘ', height=1, width=5, fg='#004C99', bg='white',
                        font=('Segoe UI Symbol', self.__size_info_and_settings_btn), side="left", x=10, y=1,
                        command=self.__info)

        self.update()
    # Метод наполнения элементами экрана кнопки 'info'
    def __info(self):
        self.clear_window()

        self.__size_title += 2
        self.__size_sub_title += 1
        self.__size_text += 1

        self.add_text('О программе:', height=2, width=400, size=self.__size_title, bold=True)
        self.add_text(f'Программа {self.master.title()} разработана в \n в 2025 году. '
                      f'\nПрограмма {self.master.title()} создана для \nбыстрой предварительной проверки '
                      '\nсостояния компьютера. \nБлагодаря этой проверке \n(или тестированию) специалист '
                      '\nтехподдержки может быстро оценить \nсостояние компьютера и получить '
                      '\nминимальные данные \nдля более углубленной \nдиагностики устройства.',
                      height=13, width=30, size=self.__size_text, side='top', justify='left')
        self.add_text(f'Version: {self.version}', height=1, width=40, size=self.__size_sub_title, side='top')
        self.add_button('Вернуться', height=2, width=10, fg='#004C99', bg='white',
                        font=('Time New Romans', self.__size_btn), side="top", x=50, command=self.main_page)

        self.__size_title -= 2
        self.__size_sub_title -= 1
        self.__size_text -= 1

        self.update()
    # Метод наполнения элементами экрана кнопки 'settings'
    def __settings(self):
        self.clear_window()

        self.__size_text += 2
        self.__size_btn += 2

        self.add_text('НАСТРОЙКИ', height=1, width=400, size=self.__size_sub_title, bold=True)
        self.add_separate()

        self.add_text('Введите почтовый сервер: ', height=2, width=400, size=self.__size_text)
        entry1 = tk.Entry(self.master, width=30, justify='center')
        entry1.insert(0, self.__mail_node)
        entry1.config(font=('Times New Roman', self.__size_btn))
        entry1.pack()

        self.add_text('Введите IMAP порт: ', height=2, width=400, size=self.__size_text)
        entry2 = tk.Entry(self.master, width=30, justify='center')
        entry2.insert(0, str(self.__imap_port))
        entry2.config(font=('Times New Roman', self.__size_btn))
        entry2.pack()

        self.add_text('Введите SMTP порт: ', height=2, width=400, size=self.__size_text)
        entry3 = tk.Entry(self.master, width=30, justify='center')
        entry3.insert(0, str(self.__smtp_port))
        entry3.config(font=('Times New Roman', self.__size_btn))
        entry3.pack()

        self.add_text('Введите первый корп узел: ', height=2, width=400, size=self.__size_text)
        entry4 = tk.Entry(self.master, width=30, justify='center')
        entry4.insert(0, str(self.__trust_node1))
        entry4.config(font=('Times New Roman', self.__size_btn))
        entry4.pack()

        self.add_text('Введите второй корп узел: ', height=2, width=400, size=self.__size_text)
        entry5 = tk.Entry(self.master, width=30, justify='center')
        entry5.insert(0, str(self.__trust_node2))
        entry5.config(font=('Times New Roman', self.__size_btn))
        entry5.pack()

        # Логика кнопки подтвердить
        def set_settings():
            try:
                self.__mail_node = str(entry1.get())
                self.__imap_port = int(entry2.get())
                self.__smtp_port = int(entry3.get())
                self.__trust_node1 = str(entry4.get())
                self.__trust_node2 = str(entry5.get())
                msg.showinfo('Successfully', 'Изменения успешно сохранены!')
            except ValueError:
                msg.showerror('Error_Value', 'Введен неправильный тип данных!')
            except TypeError:
                msg.showerror('Error_Type', 'Введен неправильный тип данных!')
            except Exception as e:
                msg.showerror('Error', f'Ошибка: {e}')

        self.__size_text -= 2
        self.__size_btn -= 2

        self.add_button('Вернуться', height=2, width=10, fg='#004C99', bg='white',
                        font=('Time New Romans', self.__size_btn), side="left", x=50, command=self.main_page)
        self.add_button('Применить', height=2, width=10, fg='#004C99', bg='white',
                        font=('Time New Romans', self.__size_btn), side="right", x=50, command=set_settings)

        self.update()

    # Метод с логикой загрузочного экрана и наполнением его контентом
    def diagnostics(self):  # {}
        queue_set = Queue()

        # Отдельный поток выполняющий методы тестирования и возвращающий их результаты
        def run_diagnostics():
            pythoncom.CoInitialize()

            try:
                uUtils = Utils()
                methods = [
                    ('Начало тестирования', 1, uUtils.GetHostName),
                    ('Получение IP', 25, uUtils.GetHostIP),
                    ('Проверка сети', 50, uUtils.CheckNet),
                    ('Проверка почты', 75, lambda: uUtils.CheckMail(int(4), str(self.__mail_node),
                                                                    int(self.__imap_port), int(self.__smtp_port))),
                    ('Проверка корпоративной сети', 90, lambda: uUtils.CheckCorpNet(trust_node1=self.__trust_node1,
                                                                                    trust_node2=self.__trust_node2)),
                    ('Проверка устройств ввода', 99, uUtils.CheckInputDevice),
                    ('Тестирование завершено', 100, int)
                ]

                for info, percent, method in methods:
                    queue_set.put((len(methods), info, percent))
                    func = method()
                    self.__result_test.append(func)
            finally:
                pythoncom.CoUninitialize()
        # Метод обновления экрана загрузки
        def main_diagnostics():
            local_count_methods = 0

            try:
                data = queue_set.get_nowait()
                local_count_methods = int(data[0])
                self.__content_diagnostics(data[1], data[2])

                if int(len(self.__result_test)) == local_count_methods:
                    self.__content_diagnostics(f'Тестирование завершено', 100)
                    self.update()
                    self.result_page()
                else:
                    self.after(100, main_diagnostics)
            except queue.Empty:
                if int(len(self.__result_test)) == local_count_methods:
                    self.__content_diagnostics(f'Тестирование завершено', 100)
                    self.update()
                    self.result_page()
                else:
                    self.after(100, main_diagnostics)

        agree = msg.askyesno('Confirmation', 'Тестирование может длиться до трех минут! '
                                             '\n\nЕсли тестирование длится больше трёх минут, \nперезапустите программу и запустите \nтестирование ещё раз.'
                                             '\n\nПодтверждаете ли вы начало тестирования?')
        if agree == True:
            self.__result_test = []
            threading.Thread(target=run_diagnostics).start()
            self.after(100, main_diagnostics)

    # Метод наполнения элементами экрана загрузки
    def __content_diagnostics(self, event:str, percent:int):
            self.clear_window()

            self.add_text(' ', height=3, width=400, size=self.__size_title)
            self.add_text(text=event, height=3, width=400, size=self.__size_title, bold=True)
            self.add_text(f'{percent} %', height=1, width=400, size=self.__size_sub_title)

            # Анимация загрузки
            width = int(self.master.winfo_width())
            height = int(self.master.winfo_height())
            radius = 30 if width >= 800 and height >= 600 else 18
            speed = 6

            self.__dx_left = speed
            self.__dx_right = -speed


            canvas = tk.Canvas(self.master, width=width, height=height, highlightthickness=0)
            canvas.pack()

            # Левый круг
            circle_left = canvas.create_oval(
                10, 40,
                10 + radius * 2, 40 + radius * 2,
                fill="#004C99", outline=""
            )
            # Правый круг
            circle_right = canvas.create_oval(
                width - 10 - radius * 2, 40,
                width - 10, 40 + radius * 2,
                fill="#004C99", outline=""
            )

            # Логика анимация загрузки
            def animate():
                if canvas.winfo_exists():
                    canvas.move(circle_left, self.__dx_left, 0)
                    canvas.move(circle_right, self.__dx_right, 0)

                    # Координаты
                    x1, _, x2, _ = canvas.coords(circle_left)
                    x3, _, x4, _ = canvas.coords(circle_right)

                    # Отскок от левой/правой стенки
                    if x1 <= 0 or x2 >= width:
                        self.__dx_left *= -1
                    if x3 <= 0 or x4 >= width:
                        self.__dx_right *= -1
                    self.master.after(10, animate)  # ~60 FPS

            animate()
            self.update()

    # Метод вывода результатов тестирования
    def result_page(self):
        self.clear_window()

        status_hostname = bool(self.__result_test[0]['status'])
        status_ip = bool(self.__result_test[1]['status'])
        status_network = bool(self.__result_test[2]['status'])
        status_co_mail = bool(self.__result_test[3]['status'])
        status_co_network = bool(self.__result_test[4]['status'])
        status_device = bool(self.__result_test[5]['status'])
        # Метод определения состояния систем
        def get_EndStatus():
            if status_hostname == False or  status_ip == False or status_network == False or status_co_mail == False:
                return ('ПЛОХО', '#FF0000')
            elif status_co_network == False or status_device == False:
                return ('НОРМАЛЬНО', '#fedd30')
            else:
                return ('ОТЛИЧНО', '#00FF00')

        data_hostname = self.__result_test[0]['data']
        data_ip = self.__result_test[1]['data']
        data_network = self.__result_test[2]['data']
        data_co_mail = self.__result_test[3]['data']
        data_co_network = self.__result_test[4]['data']
        data_device = self.__result_test[5]['data']

        error_hostname = self.__result_test[0]['msg']
        error_ip = self.__result_test[1]['msg']
        error_network = self.__result_test[2]['msg']
        error_co_mail = self.__result_test[3]['msg']
        error_co_network = self.__result_test[4]['msg']
        error_device = self.__result_test[5]['msg']

        #Наполнения элементами второй страницы результатов тестирования
        def result_page_2():
            self.clear_window()

            self.add_text('РЕЗУЛЬТАТ', height=1, width=400, size=self.__size_title, bold=True)
            self.add_separate()

            self.add_text(f'Почта', height=1, width=400, size=self.__size_sub_title, bold=True,
                          color=f'{"#004C99" if status_co_mail == True else "#FF0000"}')
            self.add_text(f'Доступ к почте: {"Присутствует" if status_co_mail == True else "Отсутствует"} ',
                          height=1, width=400, size=self.__size_text)
            self.add_text(
                f'Доступ к серверу почты: {"Присутствует" if data_co_mail["mail_node"] == True else "Отсутствует"}'
                f'\nДоступ к IMAP порту: {"Присутствует" if data_co_mail["imap_port"] == True else "Отсутствует"}'
                f'\nДоступ к SMTP порту: {"Присутствует" if data_co_mail["smtp_port"] == True else "Отсутствует"}',
                height=3, width=400, size=self.__size_text)
            self.add_text(f"Проблема: {error_co_mail if error_co_mail != '' else 'Отсутствует'}",
                          height=2, width=400, size=self.__size_text)
            self.add_separate()

            self.add_text(f'Корпоративная сеть', height=1, width=400, size=self.__size_text, bold=True,
                          color=f'{"#004C99" if status_co_network == True else "#FF0000"}')
            self.add_text(f'Доступ в корп сеть: {"Присутствует" if status_co_network == True else "Отсутствует"}'
                          f"\nПроблема: {error_co_network if error_co_network != '' else 'Отсутствует'}",
                          height=3, width=400, size=self.__size_text)
            self.add_separate()

            self.add_text(f'Обнаруженные устройства ввода', height=1, width=400, size=self.__size_sub_title, bold=True,
                          color=f'{"#004C99" if status_device == True else "#FF0000"}')
            self.add_text(f'Кол-во комп мышей: {data_device["mouse"]} | Кол-во клавиатур: {data_device["keyboard"]}',
                          height=1, width=400, size=self.__size_text)
            self.add_text(f"Проблема: {error_device if error_device != '' else 'Отсутствует'}",
                          height=2, width=400, size=self.__size_text)
            self.add_separate()

            self.add_button('На главную', height=2, width=10, fg='#004C99', bg='white',
                            font=('Time New Romans', 15), side="left", x=50, command=self.main_page)
            self.add_button('<', height=2, width=5, fg='#004C99', bg='white',
                            font=('Time New Romans', 15), side="right", x=50, command=self.result_page)

            self.update()

        self.add_text('РЕЗУЛЬТАТ', height=1, width=400, size=self.__size_title+4, bold=True)
        self.add_separate()

        if status_hostname == True:
            self.add_text(f'Имя компьютера: {data_hostname}', height=1, width=400, size=self.__size_sub_title+2, bold=True)
            self.add_separate()
        else:
            self.add_text(f'Имя компьютера: {error_hostname}', height=1, width=400, size=self.__size_sub_title+2, bold=True)
            self.add_separate()

        self.add_text(f'Состояние систем компьютера: ', height=1, width=400, size=self.__size_sub_title)
        StatusEnd = get_EndStatus()
        self.add_text(f'{StatusEnd[0]}', height=1, width=400, size=self.__size_text, color=f'{StatusEnd[1]}')
        self.add_separate()

        self.add_text(f'Cетевое подключение', height=1, width=400, size=self.__size_sub_title, bold=True,
                      color=f'{"#004C99" if status_ip == True else "#FF0000"}')
        type_connect = f'{data_ip.keys()}'.split("'")[1] if status_ip == True else 'Неизвестен'
        self.add_text(f"Тип подключения: {type_connect}"
                      f'\nIPv4 адрес: ' 
                      f'{data_ip[type_connect] if type_connect != "Неизвестен" else "Недоступен"}'
                      , height=2, width=400, size=self.__size_text)
        self.add_text(f"Проблема: {error_ip if error_ip != '' else 'Отсутствует'}", height=2, width=400, size=self.__size_text)
        self.add_separate()

        self.add_text(f'Доступ в интернет', height=1, width=400, size=self.__size_sub_title, bold=True,
                      color=f'{"#004C99" if status_network == True else "#FF0000"}')
        self.add_text(f'Доступ в сеть интернет: {"Присутствует" if status_network == True else "Отсутствует"}',
                      height=1, width=400, size=self.__size_text)
        self.add_text(
            f'Проверочный узел "google.com": {"Доступен" if data_network["google.com"] == True else "Недоступен"}'
                f'\nПроверочный узел "yandex.ru": {"Доступен" if data_network["yandex.ru"] == True else "Недоступен"}',
            height=2, width=400, size=self.__size_text)
        self.add_text(f"Проблема: {error_network if error_network != '' else 'Отсутствует'}",
                      height=2, width=400, size=self.__size_text)
        self.add_separate()

        self.add_button('На главную', height=2, width=10, fg='#004C99', bg='white',
                        font=('Time New Romans', 15), side="left", x=50, command=self.main_page)
        self.add_button('>', height=2, width=5, fg='#004C99', bg='white',
                        font=('Time New Romans', 15), side="right", x=50, command=result_page_2)

        self.update()

    # Добавить на экран изображение
    def add_image(self, path_image:str, height:int=1080, width:int=1080):
        self.__img = tk.PhotoImage(file=path_image)
        tk.Label(self.master, image=self.__img, height=height, width=width).pack()
    # Добавить на экран кнопку
    def add_button(self, button, font, bg:str, fg:str, height:int=1080, width:int=1080, command=None,
                   side:str="top", x:int=5, y:int=5):
        self.__btn = tk.Button(self.master, text=button, command=command, height=height,
                               width=width, bg=bg, fg=fg, font=font, relief="flat")
        self.__btn.pack(side=side, padx=x, pady=y)
    # Добавить на экран текст
    def add_text(self, text:str, height:int=1080, width:int=1080, size:int=45,
                 color:str='#004C99', font:str='Arial',
                 side:str="top", justify:str="center", anchor:str="center", bold:bool=False):
        self.__txt = tk.Label(self.master, text=text, height=height, width=width,
                              font=(font, size, 'bold' if bold == True else ''), fg=color,
                              justify=justify, anchor=anchor)
        self.__txt.pack(side=side, anchor=anchor, fill=tk.X)
    # Добавить на экран разделительную полосу
    def add_separate(self, height:int=5, width:int=7, color:str='#004C99'):
        canvas = tk.Canvas(self.master, width=1080, height=height, highlightthickness=0)
        canvas.pack(fill='x')

        canvas.create_line(0, 5, 1080, 5, width=width, fill=color)

# Класс для создания окна, настройки параметров окна и его запуска
class GUI_root():
    def __init__(self, widget, title:str, geometry:str, version:str,
                 size_title_text:int, size_sub_title_text:int, size_text:int, size_button:int):
        self.version = version
        self.__geometry = geometry

        # Инициализация окна и класса GUI_Widget
        self.__root = tk.Tk()
        self.__widget = widget(self.__root, self.version, size_title_text, size_sub_title_text, size_text, size_button)

        # Установка параметров окна
        self.__root.title(title)
        x, y = self.__display_center()
        self.__root.geometry(f'{geometry}+{x}+{y}')
        self.__root.resizable(False, False)
    # Этот метод помогает окну открываться всегда в центре экрана
    def __display_center(self):
        self.__root.update_idletasks()

        width = int(self.__geometry.split('x')[0])
        height = int(self.__geometry.split('x')[1])

        screen_width = self.__root.winfo_screenwidth()
        screen_height = self.__root.winfo_screenheight()

        x = (screen_width // 2) - (width // 2)
        y = (screen_height // 2) - (height // 2)
        return x, y

    # Запуск окна и наполнение его контентом
    def main(self):
        self.__widget.main_page()
        self.__root.mainloop()
        return True
    # Метод для установки иконки приложения
    def set_icon(self, path_icon:str):
        icon = tk.PhotoImage(file=path_icon)
        self.__root.iconphoto(True, icon)
        return True

if __name__ == '__main__':
    # Создание экземпляра класса и запуск приложения
    app = GUI_root(GUI_Widget, 'CORPTEK', '500x550', '0.1.5 beta',
                   22, 19,17, 13)
    app.set_icon(os.path.join(sys._MEIPASS, 'image11.png'))
    app.main()