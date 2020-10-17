from tkinter import Tk, Entry, Label, Button, Frame, LabelFrame, LEFT, TOP, StringVar
from tkinter.ttk import Combobox
from tkinter import filedialog as fd
import xlrd
import configparser
import os
from counter import calculate
from datetime import time


def choosefile(flag):
    global files, column
    if flag == 0:
        files[flag].set(fd.askopenfilename(defaultextension='.xlsx',
                                           filetypes=[('MS Excel files', '*.xlsx')]))
        column['values'] = definecols(files[0].get())
        column.current(0)
    elif flag == 1:
        files[flag].set(fd.askopenfilename(defaultextension='.csv',
                                           filetypes=[('CSV files', '*.csv')]))
    else:
        files[flag].set(fd.asksaveasfilename(defaultextension='.xlsx',
                                           filetypes=[('MS Excel files', '*.xlsx')]))


def definecols(filename):
    wb = xlrd.open_workbook(filename)
    ws = wb.sheet_by_index(0)
    return [ws.cell(0, i).value for i in range(1, ws.ncols)]


def on_closing(w):
    global files
    config = configparser.ConfigParser()
    config.add_section("Files")
    config.set("Files", "intervals", files[0].get())
    config.set("Files", "values", files[1].get())
    config.set('Files', 'calculations', files[2].get())
    with open('config.cfg', "w") as config_file:
        config.write(config_file)
    w.destroy()


def on_load():
    global files
    if os.path.exists('config.cfg'):
        config = configparser.ConfigParser()
        config.read('config.cfg')
        files[0].set(config.get("Files", "intervals"))
        files[1].set(config.get("Files", "values"))
        files[2].set(config.get("Files", "calculations"))
        if files[0]:
            column['values'] = definecols(files[0].get())
            column.current(0)


def convert_time(excel_time):
    if not excel_time:
        excel_time = 0
    excel_time = excel_time / 60
    x = int(excel_time * 24 * 3600)  # convert to number of seconds
    return time(x // 3600, (x % 3600) // 60, x % 60)  # hours, minutes, seconds


def run_script():
    global column
    wb = xlrd.open_workbook(files[0].get())
    ws = wb.sheet_by_index(0)
    data = [convert_time(ws.cell(i, column.current() + 1).value) for i in range(1, ws.nrows)]
    calculate(data, files[1].get(), files[2].get())


def ui():
    global files, column
    window = Tk()
    window.geometry("330x175")
    window.title('CSV parser')
    files = [StringVar(window), StringVar(window), StringVar(window)]
    Label(window, text='Файл интервалов').grid(column=0, row=0, padx=(10, 5), pady=(10, 3), sticky='w')
    Label(window, text='Столбец интервалов').grid(column=0, row=1, padx=(10, 5), pady=3, sticky='w')
    Label(window, text='Файл значений').grid(column=0, row=2, padx=(10, 5), pady=3, sticky='w')
    Label(window, text='Файл вывода').grid(column=0, row=3, padx=(10, 5), pady=(3, 10), sticky='w')
    Entry(window, textvariable=files[0], width=25).grid(column=1, row=0, sticky='w', pady=(10, 3))
    column = Combobox(window, values=[], width=22)
    column.grid(column=1, row=1, sticky='w')
    Entry(window, textvariable=files[1], width=25).grid(column=1, row=2, sticky='w')
    Entry(window, textvariable=files[2], width=25).grid(column=1, row=3, sticky='w')
    Button(window, text='...', command=lambda x=0: choosefile(x)).grid(column=2, row=0, padx=(10, 5), pady=(10, 3))
    Button(window, text='...', command=lambda x=1: choosefile(x)).grid(column=2, row=2, padx=(10, 5), pady=3)
    Button(window, text='...', command=lambda x=2: choosefile(x)).grid(column=2, row=3, padx=(10, 5), pady=3)
    buttons = Frame(window)
    Button(buttons, text='Сформировать', width=17, command=run_script).pack(side=LEFT, padx=(10,5), pady=(5, 10))
    Button(buttons, text='Закрыть', width=17, command=lambda x=window: on_closing(x)).pack(side=LEFT, padx=(0,10), pady=(5, 10))
    buttons.grid(column=0, row=4, columnspan=3)
    on_load()
    window.protocol("WM_DELETE_WINDOW", lambda f=window: on_closing(f))
    window.mainloop()


# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    ui()
