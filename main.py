import tkinter as tk
from tkinter import ttk
import openpyxl

combo_list = ["Subscribed", "Not Subscribed", "Other"]

def toggle_mode():
    if mode_switch.instate(["selected"]):
        style.theme_use("forest-light")
    else:
        style.theme_use("forest-dark")

def load_data():
    path = "people.xlsx"
    workbook = openpyxl.load_workbook(path)
    sheet = workbook.active

    list_values = list(sheet.values)
    print( list_values)
    for col_name in list_values[0]:
        treeview.heading(col_name, text=col_name)

    for value_tuple in list_values[1:]:
        treeview.insert('', tk.END, values=value_tuple)

def insert_row():
    name = name_entry.get()
    age = int(age_spinbox.get())
    subscription_status = status_combobox.get()
    employement_status = "Employed" if a.get() else "Unemployed"

    # insert row into excel sheet
    path = "people.xlsx"
    workbook = openpyxl.load_workbook(path)
    sheet = workbook.active
    row_values = [name, age, subscription_status, employement_status]
    sheet.append(row_values)
    workbook.save(path)

    # insert row into tree view
    treeview.insert('',tk.END, values=row_values)

    # clear form
    name_entry.delete(0, "end")
    name_entry.insert(0, "Name")
    age_spinbox.delete(0, "end")
    age_spinbox.insert(0, "Age")
    status_combobox.set(combo_list[0])
    checkbutton.state(["!selected"])

root = tk.Tk()

style = ttk.Style(root)
root.tk.call("source", "forest-light.tcl")
root.tk.call("source", "forest-dark.tcl")
style.theme_use("forest-dark")

# inside which we will have all our widgets or other frames - it is invisible
frame = ttk.Frame(root)
frame.pack() #needed to be enabled in the gui - gets added sequentially

# labelled frame 
widgets_frame = ttk.LabelFrame(frame, text="Insert Row")
widgets_frame.grid(row=0, column=0, padx=20, pady=10) # padx, pady -> padding x and y axis

# name entry
name_entry = ttk.Entry(widgets_frame)
name_entry.insert(0, "Name") #insert placeholder at the very start of entry
name_entry.bind("<FocusIn>", lambda e: name_entry.delete('0', 'end')) #if clicked in entry then delete the placeholder(from 0 to end)
name_entry.grid(row=0, column=0, padx=5, pady=(0, 5), sticky="ew") # ew - expand to fill up the label frame i.e the root, pady=(0, 5) is 0 at top and 5 at bottom

# age spinbox - to enter a number by incrementing and decrementing
age_spinbox = ttk.Spinbox(widgets_frame, from_=18, to=100) # number entry from 18 - 100
age_spinbox.insert(0, "Age")
age_spinbox.grid(row=1, column=0, padx=5, pady=5, sticky="ew")

# subscribed dropdown menu
status_combobox = ttk.Combobox(widgets_frame, values=combo_list)
status_combobox.current(0) #placeholder combo_list[0]
status_combobox.grid(row=2,column=0, padx=5, pady=5, sticky="ew")

# employed checkbox
a = tk.BooleanVar()
checkbutton = ttk.Checkbutton(widgets_frame, text="Employed", variable=a)
checkbutton.grid(row=3, column=0, padx=5, pady=5, sticky="nsew") #north south east west

# insert button
button = ttk.Button(widgets_frame, text="Insert", command=insert_row)
button.grid(row=4, column=0, padx=5, pady=5, sticky="nsew")

# separator
separator = ttk.Separator(widgets_frame)
separator.grid(row=5, column=0, padx=(20,10), pady=10, sticky="ew")

# mode switch
mode_switch = ttk.Checkbutton(
    widgets_frame, text="Mode", style="Switch", command=toggle_mode) # toggle_mode -> function
mode_switch.grid(row=6, column=0, padx=5, pady=10, sticky="ew")

# labelled frame
treeFrame = ttk.Frame(frame)
treeFrame.grid(row=0, column=1, pady=10)
treeScroll = ttk.Scrollbar(treeFrame)
treeScroll.pack(side="right", fill="y")

# treeview
cols = ("Name", "Age", "Subscription", "Employment")
treeview = ttk.Treeview(treeFrame, show="headings", 
                        yscrollcommand=treeScroll.set, columns=cols, height=13)
treeview.column("Name", width=100)
treeview.column("Age", width=50)
treeview.column("Subscription", width=100)
treeview.column("Employment", width=100)
treeview.pack()
treeScroll.config(command=treeview.yview)

load_data()


root.mainloop()