from tkinter import *
import pandas as pd
from tkinter import ttk, filedialog

root = Tk()
root.title('Excel to TreeView with Pandas')
# root.iconbitmap("palma.png")
root.geometry("700x500")

# Create frame
my_frame = Frame(root)
my_frame.pack(pady=20)

# Create treeview
my_tree = ttk.Treeview(my_frame)

# File open function
# Модуль filedialog – диалоговые окна открытия и сохранения файлов
"""
Рассмотрим две функции из модуля filedialog – askopenfilename и asksaveasfilename. 
Первая предоставляет диалоговое окно для открытия файла, вторая – для сохранения. 
Обе возвращают имя файла, который должен быть открыт или сохранен, 
но сами они его не открывают и не сохраняют. 
Делать это уже надо программными средствами самого Python.
"""


def file_open():
    filename = filedialog.askopenfilename(
        initialdir="/home",
        title="Open A File",
        filetypes=(("xlsx files", "*.xlsx"), ("xls files", "*.xls"), ("All Files", "*.*"))
    )
    if filename:
        try:
            filename = r"{}".format(filename)  # raw-string, нужен чтобы слеш \ не вызывал экранирование символов
            df = pd.read_excel(filename)
        except ValueError:
            my_label.config(text="File Couldn't Be Opened")
        except FileNotFoundError:
            my_label.config(text="File Couldn't Be Found")

    # Clear old treeview
    clear_tree()

    # Set up new Treeview
    my_tree["column"] = list(df.columns)
    my_tree["show"] = "headings"

    # Loop thru column list for headers
    for column in my_tree["column"]:
        my_tree.heading(column, text=column)

    # Put data in treeview
    df_rows = df.to_numpy().tolist()
    for row in df_rows:
        my_tree.insert("", "end", values=row)

    # Pack the treeview finally
    my_tree.pack()


def clear_tree():
    my_tree.delete(*my_tree.get_children())


# Add a menu
my_menu = Menu(root)
root.config(menu=my_menu)

# Add menu dropdown
file_menu = Menu(my_menu, tearoff=False)
my_menu.add_cascade(label="Spreadsheets", menu=file_menu)
file_menu.add_command(label="Open", command=file_open)

my_label = Label(root, text='')
my_label.pack(pady=20)

root.mainloop()
