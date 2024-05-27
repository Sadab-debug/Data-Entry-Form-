import tkinter as tk
from tkinter import ttk
from function import onEntryClick, onFocusOut, createDataFrame, clearAll, convertToExcel
import pandas as pd
from openpyxl import Workbook
from openpyxl import load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows


root = tk.Tk()

#constant variables
SCREEN_WIDTH = 980
SCREEN_HEIGHT = 720
FRAME_WIDTH = 700
FRAME_HEIGHT = 500
BG_COLOR = "#008080"  # teal 9
TITLE = "Data Entry Form"
FONT_STYLE = ("Helvetica", 20, "bold")
FG = "white"
BUTTON_FONT = ("Aerial", 15)


# Create frame for holding entry forms
main_frame = tk.Frame(root, bg='#008B8B', width=FRAME_WIDTH, height=FRAME_HEIGHT, relief="sunken", borderwidth=5, highlightbackground="black")
main_frame.place(relx=0.5, rely=0.5, anchor=tk.CENTER)
main_frame.pack_propagate(0)  # Ensure the frame maintains its size

# Labels
data_entry = tk.Label(root, text="Data Entry", font=FONT_STYLE, fg=FG, bg=BG_COLOR)
data_entry.pack(padx=100, pady=20)

label_name = tk.Label(main_frame, text="Name :", font=FONT_STYLE, fg=FG, bg="#008B8B")
label_name.place(x=10, y=20)

label_age = tk.Label(main_frame, text="Age :", font=FONT_STYLE, fg=FG, bg="#008B8B")
label_age.place(x=10, y=80)

label_phone = tk.Label(main_frame, text="Phone:", font=FONT_STYLE, fg=FG, bg="#008B8B")
label_phone.place(x=10, y=140)

label_birth = tk.Label(main_frame, text="Birth:", font=FONT_STYLE, fg=FG, bg="#008B8B")
label_birth.place(x=10, y=200)

label_blood_group = tk.Label(main_frame, text="Blood:", font=FONT_STYLE, fg=FG, bg="#008B8B")
label_blood_group.place(x=400, y=200)

label_email = tk.Label(main_frame, text="Email:", font=FONT_STYLE, fg=FG, bg="#008B8B")
label_email.place(x=10, y=320)

label_address = tk.Label(main_frame, text="Address:", font=FONT_STYLE, fg=FG, bg="#008B8B")
label_address.place(x=10, y=260)



# entry vars
name_var = tk.StringVar()
age_var = tk.StringVar()
phone_var = tk.IntVar()
birth_var = tk.StringVar()
blood_group_var = tk.StringVar()
address_var = tk.StringVar()
email_var = tk.StringVar()

# entry
name_entry = tk.Entry(main_frame, textvariable=name_var, font=FONT_STYLE, width=25).place(x=110, y=20)
age_entry = tk.Entry(main_frame, textvariable=age_var, font=FONT_STYLE, width=25).place(x=110, y=80)
phone_entry = tk.Entry(main_frame, textvariable=phone_var, font=FONT_STYLE, width=25).place(x=110, y=140)
address_entry = tk.Entry(main_frame, textvariable=address_var, font=FONT_STYLE, width=35).place(x=140, y=260)
email_entry = tk.Entry(main_frame, textvariable=email_var, font=FONT_STYLE, width=25).place(x=140, y=320)

# Create birth entry widget
placeholder_birth_entry = "dd/mm/yyyy"  
birth_entry = tk.Entry(main_frame, textvariable=birth_var, font=FONT_STYLE, width=15)
birth_entry.insert(0, placeholder_birth_entry)  # Set placeholder text
birth_entry.bind("<FocusOut>", lambda event: onFocusOut(birth_entry, placeholder_birth_entry))  # Bind focus out event
birth_entry.bind("<FocusIn>", lambda event: onEntryClick(birth_entry, placeholder_birth_entry))  # Bind click event
birth_entry.place(x=110, y=200)

#create drop down menu for blood
blood_group_var.set('A+') #set default blood group
blood_group_options = ["A+", "A-", "B+", "B-", "AB+", "AB-", "O+", "O-"]
blood_group_dropdown = ttk.OptionMenu(main_frame, blood_group_var, *blood_group_options,)
blood_group_dropdown.config(width=15)
# Configure the font separately
blood_group_dropdown["menu"].config(font=('Arial', 12))
blood_group_dropdown.place(x=500, y=203)

# buttons
clear_button = tk.Button(main_frame, text="Clear All", width=7, height=2, font=BUTTON_FONT, bg="#0000FF", cursor="arrow", command= lambda: clearAll(name_var, age_var, phone_var, birth_var, blood_group_var, address_var, email_var))
clear_button.place(x=70, y=400)


def saveToExcel():
    # Create DataFrame from tkinter variables
    df = createDataFrame(pd, name_var, age_var, phone_var, birth_var, blood_group_var, address_var, email_var)
    if df is not None:  # Ensure DataFrame is not None
        convertToExcel(Workbook, load_workbook, dataframe_to_rows, df)
    else:
        print("DataFrame is None. Data not saved.")

submit_button = tk.Button(main_frame, text="Submit", width=7, height=2, font=BUTTON_FONT, bg="#FF0000", cursor="arrow", command=saveToExcel)
submit_button.place(x=500, y=400)


root.geometry(f"{SCREEN_WIDTH}x{SCREEN_HEIGHT}")
# root.resizable(False,False)
root.configure(bg=BG_COLOR)
root.title(TITLE)
root.mainloop()