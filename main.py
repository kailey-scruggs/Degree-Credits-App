import tkinter
from tkinter import ttk
import os
import openpyxl 
import customtkinter
from tkinter import ttk

app = customtkinter.CTk()
app.geometry("350x300")
app.title("Degree Registration")
customtkinter.set_appearance_mode("system")
customtkinter.set_default_color_theme("green")


def enter_data():
    # Course info
    css = class_info_entry.get() 
    credits = credits_info_combobox.get()
    registration = reg_status_var.get()
    progress = prog_status_var.get()

# Finds and opens excell workbook 
    filepath = "C:\\Users\\Kailey\\OneDrive\\Documents\\Python\\data.xlsx" 
    if not os.path.exists(filepath):
       workbook = openpyxl.Workbook()
       sheet = workbook.active
       heading = ["Class", "Credits", "Registration", "Progress"] 
       sheet.append(heading)
       workbook.save(filepath)
    workbook = openpyxl.load_workbook(filepath)
    sheet = workbook.active
    sheet.append([css, credits, registration, progress])
    workbook.save(filepath)
  

frame = customtkinter.CTkFrame(app)
frame.pack() 

# Saving Class Info
class_info_frame =tkinter.LabelFrame(frame, text="Class Information")
class_info_frame.grid(row= 0, column= 0, padx=20, pady= 20)

class_info_label = customtkinter.CTkLabel(class_info_frame, text="Class Name", text_color="black")
class_info_label.grid(row=0, column=0) 
# Credits Info
credits_info_label = customtkinter.CTkLabel(class_info_frame, text="Credits", text_color="black")
credits_info_label.grid(row=0, column=1)
# _entry allows input from user 
class_info_entry = customtkinter.CTkEntry(class_info_frame)
credits_info_combobox = customtkinter.CTkComboBox(class_info_frame, values=["1", "2", "3", "4", "5", "6"])

class_info_entry.grid(row=1, column=0)
credits_info_combobox.grid(row=1, column=1)
# Creates space around each grid for cleaner look
for widget in class_info_frame.winfo_children():
  widget.grid_configure(padx=10, pady=5)

# Saving Course Info
course_taken_frame = tkinter.LabelFrame(frame,)
course_taken_frame.grid(row=1, column=0, sticky="news", padx=20, pady=20)

registered_label = customtkinter.CTkLabel(course_taken_frame, text="Registration Status", text_color="black")  
# Stringvar stores information from check button
reg_status_var = customtkinter.StringVar(value="Not Registered")
registered_check = customtkinter.CTkCheckBox(course_taken_frame, text="Currently Registered", variable=reg_status_var, onvalue="Registered", offvalue="Not Registered", text_color="red") 

registered_label.grid(row=0, column=0)
registered_check.grid(row=1, column=0)

prog_status_var = customtkinter.StringVar(value="Not In-Progress")
progress_check = customtkinter.CTkCheckBox(course_taken_frame, text="In-Progress", variable=prog_status_var, onvalue="In-progress", offvalue="Not in progress", text_color="red")
progress_check.grid(row=1, column=1)

for widget in course_taken_frame.winfo_children():
  widget.grid_configure(padx=10, pady=5)

# Import Data 
button = customtkinter.CTkButton(frame, text="Import Data", command= enter_data)
button.grid(row=3, column=0, sticky="news", padx=20, pady=20) 





app.mainloop() 