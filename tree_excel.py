from tkinter import *
from tkinter import ttk

root = Tk()
root.title('TreeView')
# root.iconbitmap('palma.ico')
root.geometry("500x500+500+400")

my_tree = ttk.Treeview(root)

# Define Columns
my_tree['columns'] = ("Name", "ID", "Favorite Pizza")

# Formate Columns
my_tree.column("#0", width=120, minwidth=25)
my_tree.column("Name", anchor=W, width=120)
my_tree.column("ID", anchor=CENTER, width=80)
my_tree.column("Favorite Pizza", anchor=W, width=120)

# Create Headings
my_tree.heading("#0", text="Label", anchor=W)
my_tree.heading("Name", text="Tree Name", anchor=W)
my_tree.heading("ID", text="Tree ID", anchor=CENTER)
my_tree.heading("Favorite Pizza", text="Tree Fav Pizza", anchor=W)

# Add Data
my_tree.insert(parent='', index='end', iid=0, text="Parent",
               values=("John", 1, "Pepperoni"))
my_tree.insert(parent='', index='end', iid=1, text="Parent",
               values=("Mary", "2", "Cheese"))
my_tree.insert(parent='', index='end', iid=2, text="Parent",
               values=("Tina", "3", "Ham"))
my_tree.insert(parent='', index='end', iid=3, text="Parent",
               values=("Bob", "4", "Supreme"))
my_tree.insert(parent='', index='end', iid=4, text="Parent",
               values=("Erin", "5", "Cheese"))
my_tree.insert(parent='', index='end', iid=5, text="Parent",
               values=("Wes", "6", "Margaritta"))
my_tree.insert(parent='', index='end', iid=6, text="Parent",
               values=("Armani", "7", "Four Seasons"))
# Pack to the screen
my_tree.pack(pady=20)

root.mainloop()
