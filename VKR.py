import openpyxl
import tkinter as tk
from tkinter import *
from tkinter import messagebox
import pandas as pd
def exit_window():
    windowAdm.destroy()
def exit_akk():
    windowAdm.destroy()
    avtor()
def zakrformu():
    windowDSotr.destroy()
    admWind()
def zakrformu1():
    windowDSotr.destroy()
    sotrWind()
def dobav_sotr():
    windowAdm.destroy()
    global windowDSotr
    windowDSotr = Tk()
    windowDSotr.title("Добавление сотрудника")
    frm_form = tk.Frame(relief=tk.SUNKEN, borderwidth=3)
    frm_form.pack()
    global ent_first_name
    lbl_first_name = tk.Label(master=frm_form, text="Имя и Фамилия:")
    ent_first_name = tk.Entry(master=frm_form, width=50)
    lbl_first_name.grid(row=0, column=0, sticky="e")
    ent_first_name.grid(row=0, column=1)
    global ent_otdel
    lbl_otdel = tk.Label(master=frm_form, text="Код отдела:")
    ent_otdel = tk.Entry(master=frm_form, width=50)
    lbl_otdel.grid(row=1, column=0, sticky="e")
    ent_otdel.grid(row=1, column=1)
    global ent_opit
    lbl_opit = tk.Label(master=frm_form, text="Опыт работы:")
    ent_opit = tk.Entry(master=frm_form, width=50)
    lbl_opit.grid(row=2, column=0, sticky="e")
    ent_opit.grid(row=2, column=1)

    frm_buttons = tk.Frame()
    frm_buttons.pack(fill=tk.X, ipadx=5, ipady=5)

    btn_submit = tk.Button(master=frm_buttons, text="Отправить", command=dobavit_sotr)
    btn_submit.pack(side=tk.RIGHT, padx=10, ipadx=10)
    btn_clear = tk.Button(master=frm_buttons, text="Закрыть", command=zakrformu)
    btn_clear.pack(side=tk.RIGHT, ipadx=10)
    windowDSotr.mainloop()
def dobavit_sotr():
    name=ent_first_name.get()
    otdel=ent_otdel.get()
    opt=ent_opit.get()
    excel_sheet = excel_file['Сотрудники']
    row = (name, otdel, opt)
    excel_sheet.append(row)
    excel_file.save('Ds.xlsx')
    messagebox.showinfo("Ура!", "Сотрудник " + name + " добавлен!")
def sdelka():
    global windowDSotr
    windowAdm.destroy()
    windowDSotr = Tk()
    windowDSotr.title("Добавление сделки")
    frm_form = tk.Frame(relief=tk.SUNKEN, borderwidth=3)
    frm_form.pack()
    global ent_name
    lbl_name = tk.Label(master=frm_form, text="Сотрудник:")
    ent_name = tk.Entry(master=frm_form, width=50)
    ent_name.insert(0, username)
    ent_name.configure(state=tk.DISABLED)

    lbl_name.grid(row=0, column=0, sticky="e")
    ent_name.grid(row=0, column=1)
    global ent_data
    lbl_data= tk.Label(master=frm_form, text="Дата:")
    ent_data = tk.Entry(master=frm_form, width=50)
    lbl_data.grid(row=1, column=0, sticky="e")
    ent_data.grid(row=1, column=1)
    global ent_profit
    lbl_profit = tk.Label(master=frm_form, text="Доход:")
    ent_profit = tk.Entry(master=frm_form, width=50)
    lbl_profit.grid(row=2, column=0, sticky="e")
    ent_profit.grid(row=2, column=1)
    frm_buttons = tk.Frame()
    frm_buttons.pack(fill=tk.X, ipadx=5, ipady=5)
    btn_submit = tk.Button(master=frm_buttons, text="Отправить", command=dobavit_sdelki)
    btn_submit.pack(side=tk.RIGHT, padx=10, ipadx=10)
    btn_clear = tk.Button(master=frm_buttons, text="Закрыть", command=zakrformu1)
    btn_clear.pack(side=tk.RIGHT, ipadx=10)
    windowDSotr.mainloop()

def dobavit_sdelki():
    emp_name=ent_name.get()
    date=ent_data.get()
    profit=ent_profit.get()
    excel_sheet = excel_file['Сделки']
    row = (date, emp_name, profit)
    excel_sheet.append(row)
    excel_file.save('Ds.xlsx')
    messagebox.showinfo("Мои поздравления!", "Сделка прошла успешно!")
def fil():
    global windowDSotr
    windowAdm.destroy()
    windowDSotr = Tk()
    windowDSotr.title("Добавление филиала")
    frm_form = tk.Frame(relief=tk.SUNKEN, borderwidth=3)
    frm_form.pack()
    global ent_pdr
    lbl_pdr = tk.Label(master=frm_form, text="Подразделение:")
    ent_pdr = tk.Entry(master=frm_form, width=50)
    lbl_pdr.grid(row=0, column=0, sticky="e")
    ent_pdr.grid(row=0, column=1)
    global ent_city
    lbl_city = tk.Label(master=frm_form, text="Город:")
    ent_city = tk.Entry(master=frm_form, width=50)
    lbl_city.grid(row=1, column=0, sticky="e")
    ent_city.grid(row=1, column=1)

    frm_buttons = tk.Frame()
    frm_buttons.pack(fill=tk.X, ipadx=5, ipady=5)

    btn_submit = tk.Button(master=frm_buttons, text="Отправить", command=dobavit_fil)
    btn_submit.pack(side=tk.RIGHT, padx=10, ipadx=10)
    btn_clear = tk.Button(master=frm_buttons, text="Закрыть", command=zakrformu1)
    btn_clear.pack(side=tk.RIGHT, ipadx=10)
    windowDSotr.mainloop()
def dobavit_fil():
    pdr = ent_pdr.get()
    city = ent_city.get()
    excel_sheet = excel_file['Филиалы']
    row = (city, podr)
    excel_sheet.append(row)
    excel_file.save('Ds.xlsx')
def clicked(): #Обработка авторизации
    global username
    username = username_entry.get()
    password = password_entry.get()

    # выводим в диалоговое окно введенные пользователем данные
    if (username=='A') & (password=='123'):
        window.destroy()
        admWind()
    else:
        if (username == 'B') & (password == '123'):
            window.destroy()
            sotrWind()
        else:
            messagebox.showinfo("Хм","Проверьте правильность введённых данных!")
def open_sdelki():
    global windowDSotr
    windowAdm.destroy()
    windowDSotr = Tk()
    text_box = tk.Text()
    text_box.pack()
    excel_file_pd = pd.read_excel('Ds.xlsx', sheet_name='Сделки')
    data = pd.DataFrame(excel_file_pd, columns=['Date', 'Employee Name', 'Profit'])
    text_box.insert(tk.END, data)
def open_sotr():
    global windowDSotr
    windowAdm.destroy()
    windowDSotr = Tk()
    text_box = tk.Text()
    text_box.pack()
    excel_file_pd = pd.read_excel('Ds.xlsx', sheet_name='Сотрудники')
    data = pd.DataFrame(excel_file_pd, columns=[ 'Employee Name', 'Un_code','Work Experience'])
    text_box.insert(tk.END, data)
def sotrWind():
    global windowAdm
    windowAdm = tk.Tk()
    windowAdm.title('Меню сотрудника')
    windowAdm.geometry('400x500')
    windowAdm.configure(bg='black')
    label1 = tk.Label(
        text="Добро пожаловать, " +username + '!',
        fg="white",
        bg="black",
        width=50,
        height=5,
        font='Arial 14'
    )
    label2 = tk.Label(
        text="МЕНЮ",
        fg="white",
        bg="black",
        width=50,
        height=2,
        font='Arial 14'
    )
    button3 = tk.Button(
        text="Провести сделку",
        width=35,
        bg="white",
        fg="black",
        command=sdelka
    )

    label1.pack()
    label2.pack()
    button3.pack()
    frm_buttons = tk.Frame()
    frm_buttons.configure(bg='black')
    frm_buttons.pack(fill=tk.X, ipadx=15, ipady=34)
    btn_close = tk.Button(master=frm_buttons, text="Закрыть", command=exit_window)
    btn_close.pack(side=tk.RIGHT, padx=10, ipadx=10)
    btn_exit = tk.Button(master=frm_buttons, text="Выйти", command=exit_akk)
    btn_exit.pack(side=tk.RIGHT, ipadx=10)
    windowAdm.mainloop()
def admWind():
    global windowAdm
    windowAdm = tk.Tk()
    windowAdm.title('Меню директора')
    windowAdm.geometry('400x500')
    windowAdm.configure(bg='#708090')
    label1 = tk.Label(
        text="Добро пожаловать, " +username + '!',
        fg="white",
        bg="#708090",
        width=50,
        height=5,
        font='Arial 14'
    )
    label2 = tk.Label(
        text="МЕНЮ",
        fg="white",
        bg="#708090",
        width=50,
        height=2,
        font='Arial 24'
    )
    '''provesti_sdelku = PhotoImage(file="provesti_sdelku.png")
    button1 = Button(windowAdm,image=provesti_sdelku)
    button1["border"] = "0"'''
    izmbud = PhotoImage(file="izmbud.png")
    button1 = tk.Button(
        image=izmbud,
        text="Изменить годовой бюджет подразделения",
        bg="#708090",
        border="0",
        )
    dobav_sotrudnika = PhotoImage(file="dobav_sotrudnika.png")
    button2 = tk.Button(
        image=dobav_sotrudnika,
        bg="#708090",
        border="0",
        command=dobav_sotr
    )
    provesti_sdelku = PhotoImage(file="provesti_sdelku.png")
    button3 = tk.Button(
        image=provesti_sdelku,
        bg="#708090",
        border="0",
        command=sdelka
    )
    dob_fil = PhotoImage(file="dob_fil.png")
    button4 = tk.Button(
        image=dob_fil,
        bg="#708090",
        border="0",
        command=fil
    )
    uvolit = PhotoImage(file="uvolit.png")
    button5 = tk.Button(
        image=uvolit,
        bg="#708090",
        border="0",
    )
    perevod = PhotoImage(file="perevod.png")
    button6 = tk.Button(
        image=perevod,
        bg="#708090",
        border="0",
    )

    label1.pack()
    label2.pack()
    button3.pack()
    button2.pack()
    button5.pack()
    button6.pack()
    button1.pack()
    button4.pack()

    spis_sotr=PhotoImage(file="spis_sotr.png.png")
    spis_sdelki = PhotoImage(file="spis_sdelki.png.png")

    frm_buttons1 = tk.Frame()
    frm_buttons1.pack(fill=tk.X, ipadx=15, ipady=15)
    btn_sotrinf = tk.Button(master=frm_buttons1, bg="#708090",border="0",image=spis_sotr, command=open_sotr)
    btn_sotrinf.pack(side=tk.RIGHT, padx=20, ipadx=20)
    btn_sdelkinf = tk.Button(master=frm_buttons1, bg="#708090",border="0", command=open_sdelki)
    frm_buttons1.configure(bg='#708090')
    btn_sdelkinf.pack(side=tk.RIGHT, padx=20)

    frm_buttons = tk.Frame()
    frm_buttons.configure(bg='#708090')
    frm_buttons.pack(fill=tk.X, ipadx=15, ipady=34)
    btn_close = tk.Button(master=frm_buttons, text="Закрыть", command=exit_window)
    btn_close.pack(side=tk.RIGHT, padx=10, ipadx=10)
    btn_exit = tk.Button(master=frm_buttons, text="Выйти", command=exit_akk)
    btn_exit.pack(side=tk.RIGHT, ipadx=10)
    windowAdm.mainloop()

def avtor():
    global window
    window = Tk()
    # заголовок окна
    window.title('Авторизация')
    # размер окна
    window.geometry('450x230')
    # можно ли изменять размер окна - нет
    window.resizable(False, False)

    # кортежи и словари, содержащие настройки шрифтов и отступов
    font_header = ('Arial', 15)
    font_entry = ('Arial', 12)
    label_font = ('Arial', 11)
    base_padding = {'padx': 10, 'pady': 8}
    header_padding = {'padx': 10, 'pady': 12}

    # заголовок формы: настроены шрифт (font), отцентрирован (justify), добавлены отступы для заголовка
    # для всех остальных виджетов настройки делаются также
    main_label = Label(window, text='Авторизация', font=font_header, justify=CENTER, **header_padding)
    # помещаем виджет в окно по принципу один виджет под другим
    main_label.pack()

    # метка для поля ввода имени
    username_label = Label(window, text='Имя пользователя', font=label_font , **base_padding)
    username_label.pack()

    # поле ввода имени
    global username_entry
    username_entry = Entry(window, bg='#fff', fg='#444', font=font_entry)
    username_entry.pack()

    # метка для поля ввода пароля
    password_label = Label(window, text='Пароль', font=label_font , **base_padding)
    password_label.pack()

    # поле ввода пароля
    global password_entry
    password_entry = Entry(window, bg='#fff', fg='#444', font=font_entry)
    password_entry.pack()

    # кнопка отправки формы
    send_btn = Button(window, text='Войти', command=clicked)
    send_btn.pack(**base_padding)

    # запускаем главный цикл окна
    window.mainloop()
excel_file = openpyxl.load_workbook('Ds.xlsx')
avtor()
