from tkinter import LEFT, BOTTOM, W, S, Toplevel, BOTH
from tkinter.ttk import Combobox, Frame, Label
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk
from matplotlib.figure import Figure


df = ''


def graph_report(window, data):
    """ графическое представление результатов """
    global df, cmb

    df = data

    graph = Toplevel(window)
    graph.title('Графическое представление')
    graph.minsize(width=400, height=200)

    choose_box = Frame(graph)
    lbl = Label(choose_box, text='Выбор параметра')
    lbl.pack(side=LEFT, anchor=W, padx=10, pady=10)

    cmb = Combobox(choose_box, values=list(df.index), width=35)
    cmb.current(0)
    cmb.pack(side=LEFT, anchor=W, padx=10, pady=10)

    choose_box.pack(anchor=S, padx=10, pady=10)

    fig = Figure(figsize=(5, 4), dpi=100)
    make_fig(fig, cmb.current())
    canvas = FigureCanvasTkAgg(fig, master=graph)
    canvas.draw()
    canvas.get_tk_widget().pack(side=BOTTOM, fill=BOTH, expand=1)

    cmb.bind('<<ComboboxSelected>>', lambda event, canvas=canvas, ax=fig: change_canvas(event, canvas, ax))

    toolbar = NavigationToolbar2Tk(canvas, graph)
    toolbar.update()
    canvas.get_tk_widget().pack(fill=BOTH, expand=1)


def make_fig(fig, flag):
    global df
    """ функция формирования фигуры matplotlib согласно выбора пользователя в Combobox """
    fig.clear()
    ax = fig.subplots()
    ax.plot(['{!s}:{!s}-{!s}:{!s}'.format(c[0].minute, c[0].second, c[1].minute, c[1].second) for c in df], df.iloc[flag])
    fig.autofmt_xdate()

def autolabel(rects, ax):
    """ функция размещения меток на столбцах диаграммы """
    for rect in rects:
        height = rect.get_height()
        ax.annotate('{:d}'.format(int(height) if height > 0 else 0),
                    xy=(rect.get_x() + rect.get_width() / 2, height),
                    xytext=(0, 3),
                    textcoords="offset points",
                    ha='center', va='bottom')


def change_canvas(event, canvas, fig):
    global cmb
    """ обработчик события изменения значения Combobox """
    if event:
        make_fig(fig, cmb.current())
        canvas.draw()
        canvas.get_tk_widget().pack(side=BOTTOM, fill=BOTH, expand=1)