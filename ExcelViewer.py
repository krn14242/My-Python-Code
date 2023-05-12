import tkinter as tk
from tkinter import ttk
import openpyxl as op


def load_data():
    path = "Child population by age group.xlsx"
    workbook = op.load_workbook(path)
    sheet = workbook["ExcelViewer"]

    list_values = list(sheet.values)
    for col_name in list_values[0]:
        treeview.heading(col_name, text=col_name)

    for value_tuple in list_values[1:]:
        treeview.insert('', tk.END, values=value_tuple)


def insert_row():
    name = name_entry.get()
    age = int(age_spin_box.get())
    subscription_status = status_combobox.get()
    employment_status = "Employed" if a.get() else "Unemployed"

#  Insert row into Excel Sheet
    path = "Child population by age group.xlsx"
    workbook = op.load_workbook(path)
    sheet = workbook["ExcelViewer"]
    row_values = [name, age, subscription_status, employment_status]
    sheet.append(row_values)
    workbook.save(path)

#  Insert row into treeview
    treeview.insert('', tk.END, values=row_values)

#  Clear the values
    name_entry.delete(0, "end")
    name_entry.insert(0, "Name")
    age_spin_box.delete(0, "end")
    age_spin_box.insert(0, "Age")
    status_combobox.set(combo_list[0])
    checkbutton.state(["!selected"])


def toggle_mode():
    if mode_switch.instate(["selected"]):
        style.theme_use("forest-light")
    else:
        style.theme_use("forest-dark")


root = tk.Tk()
root.title("Excel Viewer")
style = ttk.Style(root)
root.tk.call("source", "forest-dark.tcl")
root.tk.call("source", "forest-light.tcl")
style.theme_use("forest-dark")

combo_list = ["Subscribed", "Not Subscribed", "Other"]

frame = ttk.Frame(root)
frame.pack()

label_frame = ttk.LabelFrame(frame, text="Insert Row")
label_frame.grid(row=0, column=0, padx=20, pady=10)

name_entry = ttk.Entry(label_frame)
name_entry.insert(0, "Name")
name_entry.bind("<FocusIn>", lambda e: name_entry.delete('0', 'end'))
name_entry.grid(row=0, column=0, padx=5, pady=(0, 5), sticky="ew")

age_spin_box = ttk.Spinbox(label_frame, from_=18, to=100)
age_spin_box.insert(0, "Age")
age_spin_box.grid(row=1, column=0, padx=5, pady=5, sticky="ew")

status_combobox = ttk.Combobox(label_frame, values=combo_list)
status_combobox.current(0)
status_combobox.grid(row=2, column=0, padx=5, pady=5, sticky="ew")

a = tk.BooleanVar()
checkbutton = ttk.Checkbutton(label_frame, text="Employed", variable=a)
checkbutton.grid(row=3, column=0, padx=5, pady=5, sticky="news")

button = ttk.Button(label_frame, text="Insert", command=insert_row)
button.grid(row=4, column=0, padx=5, pady=5, sticky="news")

separator = ttk.Separator(label_frame)
separator.grid(row=5, column=0, padx=(20, 10), pady=10, sticky="ew")

mode_switch = ttk.Checkbutton(label_frame, text="Mode (Dark or Light)", style="Switch", command=toggle_mode)
mode_switch.grid(row=6, column=0, padx=5, pady=10, sticky="news")

treeFrame = ttk.Frame(frame)
treeFrame.grid(row=0, column=1, pady=10, sticky="ew")
treeScroll = ttk.Scrollbar(treeFrame)
treeScroll.pack(side="right", fill="y")

cols = ("Date", "Total Deaths", "Male", "Female")
treeview = ttk.Treeview(treeFrame, show="headings", columns=cols, yscrollcommand=treeScroll.set, height=13)
treeview.column("Date", width=75)
treeview.column("Total Deaths", width=75)
treeview.column("Male", width=75)
treeview.column("Female", width=75)
treeview.pack()
treeScroll.config(command=treeview.yview)
load_data()

root.mainloop()
