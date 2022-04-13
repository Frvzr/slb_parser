import tkinter
from tkinter import messagebox, RIGHT, END, StringVar, Event, LEFT, SINGLE
import tkinter.ttk as ttk

def langSelect(selection):
    print()

#функция очистки
def clear():
    srxShoulder_entry.delete(0, END)
    xoverLength_entry.delete(0, END)
    label1.place_forget()
    label_InMin.place_forget()
    label_Min.place_forget()
    label_InMax.place_forget()
    label_Max.place_forget()
    
def srxShoulder(event):
    srxShoulder_entry.delete(0,END)
    usercheck=True

def xoverLength(event):
    xoverLength_entry.delete(0, END)
    passcheck=True
    
#функция для подсчета рзультата    
def result():
    a = float(srxShoulder_entry.get())
    b = float(xoverLength_entry.get())
    Min = float(79)
    e = float(25.4)
    Max = float(82)
    
    label1.place(x=40, y=260)
    label_Min.place(x = 170, y = 80)
    label_InMin.place(x=170, y=100)
    label_Max.place(x=170, y=140)
    label_InMax.place(x=170, y=160)
    
    res_min = float(a+b+Min)
    label_Min["text"] = res_min
    
    
    resIn_min = float(res_min/e)
    resIn_min = round(resIn_min, 3)
    label_InMin["text"] = resIn_min
    label_InMin.config(font=('Arial',10,'bold'))

    res_max = float(a+b+Max)
    label_Max["text"] = res_max
    
    resIn_max = float(res_max/e)
    resIn_max = round(resIn_max, 3)
    label_InMax["text"] = resIn_max
    label_InMax.config(font=('Arial',10,'bold'))

    res_minError = resIn_min
    if float(res_minError)<41.22:
        label1["text"] = "Меньше минимально допустимой длины 41.28in"
        label1.config(font=('Arial', 14, 'bold'), width=40, fg="red")
    elif float(res_minError)>44.02:
        label1["text"] = "Больше максимально допустимой длины 44.02in"
        label1.config(font=('Arial', 14, 'bold'), width=40, fg="red")
    else:
        label1["text"] = "Длина в допустимых пределах"
        label1.config(font=('Arial', 15, 'bold'), width=40, fg="green")

    if srxShoulder_entry.get() == "0":
        label1["text"] = "Введите корректные данные"
        label1.config(font=('Arial', 15, 'bold'), width=40, fg='red')
        label_InMin.place_forget()
        label_Min.place_forget()
        label_InMax.place_forget()
        label_Max.place_forget()
    elif xoverLength_entry.get() == "0":
        label1["text"] = "Введите корректные данные"
        label1.config(font=('Arial', 15, 'bold'), width=40, fg="red")
        label_InMin.place_forget()
        label_Min.place_forget()
        label_InMax.place_forget()
        label_Max.place_forget()
        
def resultReturn(event):
    result()
            
#меню About RTLMcalc
def insert_text():
    messagebox.showinfo("ADS SRX/RTLM calculator", "Version 2.0\nUpdate 05.12.2019\n(c) IKoposhilov, 2019")

class tableinmm(tkinter.Tk):
    def __init__(self):
        tkinter.Tk.__init__(self)
        self.title("in = mm")
        self.geometry('220x305')
        self.resizable(0, 0)
        t = SimpleTable(self, 16,2)
        t.pack(side="top", fill="x")
        t.set(0,0,"Inch")
        t.set(0,1, "Millimeter")
        t.set(1,0, "1/16")
        t.set(1,1, "0.0625")
        t.set(2,0, "2/16")
        t.set(2,1, "0.125")
        t.set(3,0, "3/16")
        t.set(3,1, "0.1875")
        t.set(4,0, "4/16")
        t.set(4,1, "0.25")
        t.set(5,0, "5/16")
        t.set(5,1, "0.3125")
        t.set(6,0, "6/16")
        t.set(6,1, "0.375")
        t.set(7,0, "7/16")
        t.set(7,1, "0.435")
        t.set(8,0, "8/16")
        t.set(8,1, "0.5")
        t.set(9,0, "9/16")
        t.set(9,1, "0.5625")
        t.set(10,0, "10/16")
        t.set(10,1, "0.625")
        t.set(11,0, "11/16")
        t.set(11,1, "0.6875")
        t.set(12,0, "12/16")
        t.set(12,1, "0.75")
        t.set(13,0, "13/16")
        t.set(13,1, "0.8125")
        t.set(14,0, "14/16")
        t.set(14,1, "0.875")
        t.set(15,0, "15/16")
        t.set(15,1, "0.9375")
        
class SimpleTable(tkinter.Frame):
    def __init__(self, parent, rows=16, columns=2):
        tkinter.Frame.__init__(self, parent, background="black")
        self._widgets = []
        for row in range(rows):
            current_row = []
            for column in range(columns):
                label = tkinter.Label(self, text="%s/%s" % (row, column), 
                                 borderwidth=0, width=10)
                label.grid(row=row, column=column, sticky="nsew", padx=1, pady=1)
                current_row.append(label)
            self._widgets.append(current_row)

        for column in range(columns):
            self.grid_columnconfigure(column, weight=1)


    def set(self, row, column, value):
        widget = self._widgets[row][column]
        widget.configure(text=value)

#menu Exit
def menu_exit():
    root.destroy()
    exit()
    
def buttonexit(event):
    menu_exit()

root = tkinter.Tk()
root.title("ADS SRX/RTLM calculator") #название в шапке
root.geometry('550x340')      #размер окна
root.resizable(0, 0) #запрет на изменение размера окна

#меню
main_menu = tkinter.Menu()
root.config(menu=main_menu)

file_menu = tkinter.Menu()
file_menu.add_command(label="About", command = insert_text)
file_menu.add_command(label="in = mm", command = tableinmm)
file_menu.add_separator()
file_menu.add_command(label="Exit", command=menu_exit)

main_menu.add_cascade(label="File", menu = file_menu)

#задаем название лэйблов
srxShoulder_label = tkinter.Label(text="Расстояние от торца SRX до торца картриджа (без плага)(мм):")
xoverLength_label = tkinter.Label(text="Длина переводника между торцами (мм):")
resLenghtMin_label = tkinter.Label(text="Минимальная длина в мм:")
resLenghtInMin_label = tkinter.Label(text="Минимальная длина в in:")
resLenghtMax_label = tkinter.Label(text="Максимальная длина в мм")
resLenghtInMax_label = tkinter.Label(text="Максимальная длина в in")
label_el = tkinter.Label(text="____________________________________________________________________________________",font ='arial 8')
label1 = tkinter.Label()
label_Min = tkinter.Label()
label_InMin = tkinter.Label()
label_Max = tkinter.Label()
label_InMax = tkinter.Label()
label_el2 = tkinter.Label(text="____________________________________________________________________________________",font ='arial 8')

#задаем положение лэйблов
srxShoulder_label.place(x = 10, y = 10)
xoverLength_label.place(x = 10, y = 35)
resLenghtMin_label.place(x = 10, y = 80)
resLenghtInMin_label.place(x = 10, y = 100)
resLenghtMax_label.place(x = 10, y = 140)
resLenghtInMax_label.place(x = 10, y = 160)
label_el.place(x=10, y = 55)
label1.place(x = 40, y = 270)
label_Min.place(x = 170, y = 80)
label_InMin.place(x = 170, y = 100)
label_Max.place(x = 170, y = 140)
label_InMax.place(x = 170, y = 160)
label_el2.place(x = 10, y = 180)

z=StringVar()
x=StringVar()
usercheck=False
passcheck=False

#задаем поле для ввода
srxShoulder_entry = tkinter.Entry(root,textvariable=z, justify=RIGHT)
xoverLength_entry = tkinter.Entry(root,textvariable=x, justify=RIGHT)

#выпадающий список v3
langLabel = tkinter.Label(text = 'Выбор языка:')
langLabel.place(x = 380, y = 300)
combobox = ttk.Combobox(values = [u"Rus",u"Eng"],state='readonly', width = 7)
combobox.set(u"Rus")
combobox.place(x = 465, y = 300)

#расположение поля ввода
srxShoulder_entry.place(x=370, y = 10)
xoverLength_entry.place(x=370, y = 35)

#вставка начальных данных
srxShoulder_entry.insert(0,0)
xoverLength_entry.insert(0,0)

#кнопки
display_button = tkinter.Button(text = "Результат", command = result)
clear_button = tkinter.Button(text = "Очистить", command = clear)

#расположение кнопок
display_button.place(x = 200, y = 220)
clear_button.place(x = 280, y = 220)

srxShoulder_entry.bind("<Button>",srxShoulder)
xoverLength_entry.bind("<Button>",xoverLength)

root.bind("<Return>", resultReturn)
root.bind("<Escape>", buttonexit)

root.mainloop()
