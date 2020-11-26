import tkinter as tk
import numpy as np
import pandas as pd
from datetime import date
import shutil
import os
from os.path import join
import path as my_p

test_sch = '0010;20-3060500;100\n'

now_data = str(date.today())
root = tk.Tk()
root.title('Program do skanowania AREX-Schmitd')
# root.geometry('500x500')

df = pd.DataFrame(columns=['Lokalizacja Schmitd', 'Numer Schmitd', 'Ilość','J.m.'])     # tworzenie DataFrame z kolumnami
is_okej = [False,False]

# funkcjia rozdzielająca słowa w linii tekstu odzielonych znakiem ';'
def separate_text_from_scaner(string):
    new_text = string.split(';')                    # wstawianie do tablicy
    return (new_text[0],new_text[1],new_text[2])    # return tuple

# funkcja usuwająca tekst z pierwszego widżetu textu
def del_text():
    global is_okej
    my_text.delete(1.0, 'end')
    df.drop(df.index, inplace=True)                 # czyszczenie zawarrtości DataFrame (zostawianie struktury)
    my_text_df.delete(1.0, 'end')  # usuwaniie zawartości my_text_df
    my_button.configure(bg='white')
    my_button_to_exel.configure(bg='white')
    my_button_to_office.configure(bg='white')
    is_okej = [False,False]

# funkcja zamieniająca ciąg znaków z pierwszego widżetttu Text na DataFrame
# oraz wstawia DataFrame do drugiego widżetu w celach poglądowych
# oraz dokonuje drobnej modyfikacji
def set_text():
    global all_text
    global all_text_sep
    global is_okej
    my_text_df.delete(1.0,'end')                                # usuwaniie zawartości my_text_df
    all_text = my_text.get(1.0,'end')                           # wstawianie do atblicy zawartości my_text
    all_text_sep = all_text.split('\n')                         # rozdzielanie tablicy liniami
    if all_text_sep[-1] == '':                                  # usuwanie ostatniego pustego elementu w tabeli
        all_text_sep.remove('')
    for x,y in enumerate(all_text_sep):                         # wstawianie danych do DataFrame
        if len(y) != 0:
            (first,second,third) = separate_text_from_scaner(y)     # wydobywanie (lok.; num. sch.; ilo.)
            df.loc[x,'Lokalizacja Schmitd'] = first                 # wstawianie
            df.loc[x,'Numer Schmitd'] = second                      # -----
            df.loc[x,'Ilość'] = float(third) / 100.0                # -----
            df.loc[x,'J.m.'] = '100szt.'                            # -----v
    df.index += 1
    my_text_df.insert(tk.END, str(df))                          # wstawienie do widżetu my_text_df struktury DataFrame (poukładanej)
    df.index -= 1
    my_button.configure(bg='green')
    is_okej[0]=True


# funkcja zapisująca DataFrame do exela
def write_to_exel():
    global is_okej
    if is_okej[0] == True:
        # Zapisywanie pliku do Exela
        global name
        name = now_data + '-KANBAN.xlsx'
        writer = pd.ExcelWriter(name)#'test.xlsx')#now_data+'-KANBAN.xlsx')
        df.to_excel(writer)
        writer.save()
        writer.close()
        my_button_to_exel.configure(bg='green')
        is_okej[1]=True

def send_to_office():
    global is_okej
    if is_okej[1] == True and is_okej[0] == True:
        path = os.path.abspath(name)
        # path2 =r'\\DESKTOP-70MER87\Users\arex-\Desktop\Folder udostępniony BIURO'
        try:
            shutil.move(path,my_p.office)
            my_button_to_office.configure(bg='green')
        except:
            label_error.configure(text='plik o tej nazwie już istnieje')
            my_button_to_office.configure(bg='red')

# C:\Users\kwiec\PycharmProjects\Test_zczytywania_scanera\2020-11-06-KANBAN.xlsx

my_text = tk.Text(root, width=60, height=20)
my_text.grid(column=0,row=0,columnspan=2)

my_text_df = tk.Text(root, width=60, height=20)
my_text_df.grid(column=0,row=2,columnspan=2)

my_button_del = tk.Button(root, text='Usuń',command=del_text)
my_button_del.grid(column=1,row=1)
my_button_del.configure(font=15)

my_button = tk.Button(root, text='1.Zatwierdź',command=set_text)
my_button.grid(column=0,row=1)
my_button.configure(font=15)

my_button_to_exel = tk.Button(root, text='2.Zapisz do exela',command=write_to_exel)
my_button_to_exel.grid(column=0,row=3)
my_button_to_exel.configure(font=15)

my_button_to_office = tk.Button(root, text='3.Prześlij do biura',command=send_to_office)
my_button_to_office.grid(column=1,row=3)
my_button_to_office.configure(font=15)

label_error = tk.Label(root,padx=4,text='komunikaty o błędach',fg='red',pady=4,bg='white',font=15)
label_error.grid(row=4,columnspan=2)

root.mainloop()

#   Przykład
# 0010;0202848-8;20
# 0050;20-0203139-1;100
# 0010;0202848-8;20
# 0050;0204203-4;50
# 0010;0208124-8;100
# 0050;0288450-0;100
