import tkinter
from tkinter import ttk
import os
import openpyxl 
import customtkinter
from tkinter import ttk
import pandas as pd

app = customtkinter.CTk()
app.geometry("500x500")
app.title("Credits Needed To Finish")
customtkinter.set_appearance_mode("light")
customtkinter.set_default_color_theme("dark-blue")
frame = customtkinter.CTkFrame(app)
frame.pack() 


# Subtracting columns with pandas
data = pd.read_excel(r"c:\\Users\\Kailey\\OneDrive\\Documents\\Python\data.xlsx", engine="openpyxl")
dataframe = pd.DataFrame(data, columns=["Credits", "Total Creds Needed"])
dataframe['Total Creds Needed'] = dataframe['Credits'] - dataframe['Total Creds Needed']



def enter_data():
    # Course info
    css = class_info_entry.get() 
    credits = credits_info_combobox.get()
    registration = reg_status_var.get()
    progress = prog_status_var.get()
    completed = complete_status_var.get()
    grade = grade_entry.get()
    total_creds = total_credits_entry.get()

# Finds and opens excell workbook 
    filepath = r"c:\\Users\\Kailey\\OneDrive\\Documents\\Python\data.xlsx"
    if not os.path.exists(filepath):
       workbook = openpyxl.Workbook()
       sheet = workbook.active
       heading = ["Class", "Credits", "Registration", "Progress", "Total Creds Needed", "Final Grade", "Completed"] 
       sheet.append(heading)
       workbook.save(filepath)
    workbook = openpyxl.load_workbook(filepath)
    sheet = workbook.active
    sheet.append([css, credits, registration, progress, total_creds, grade, completed])
    workbook.save(filepath)
    workbook.close()



# Saving Class Info
class_info_frame =tkinter.LabelFrame(frame, text="Class Information")
class_info_frame.grid(row= 0, column= 0, padx=20, pady=20)
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
# Grade input
grade = customtkinter.CTkLabel(class_info_frame, text="Final Grade", text_color="black")
grade.grid(row=2, column=1)
grade_entry = customtkinter.CTkEntry(class_info_frame)
grade_entry.grid(row=3, column=1)
# Total Credits Needed
total_credits_label = customtkinter.CTkLabel(class_info_frame, text="Total Credits Needed", text_color="black")
total_credits_label.grid(row=2, column=0)
total_credits_entry = customtkinter.CTkEntry(class_info_frame)
total_credits_entry.grid(row=3, column=0)

# Creates space around each Frame and Grid for cleaner look
for widget in class_info_frame.winfo_children():
  widget.grid_configure(padx=10, pady=5)

# Saving Course Info
course_taken_frame = tkinter.LabelFrame(frame)
course_taken_frame.grid(row=1, column=0, sticky="news", padx=20, pady=20)

registered_label = customtkinter.CTkLabel(course_taken_frame, text="Registration Status", text_color="black")  
# Stringvar stores information from check button
reg_status_var = customtkinter.StringVar(value="Not Registered")
registered_check = customtkinter.CTkCheckBox(course_taken_frame, text="Currently Registered", variable=reg_status_var, onvalue="Registered", offvalue="Not Registered", text_color="black") 

registered_label.grid(row=0, column=0)
registered_check.grid(row=1, column=0)

prog_status_var = customtkinter.StringVar(value="Not In-Progress")
progress_check = customtkinter.CTkCheckBox(course_taken_frame, text="In-Progress", variable=prog_status_var, onvalue="In-progress", offvalue="Not in progress", text_color="black")
progress_check.grid(row=1, column=1)

complete_status_var = customtkinter.StringVar(value="Not Completed")
complete_check = customtkinter.CTkCheckBox(course_taken_frame, text="Completed", variable=complete_status_var, onvalue="Completed", offvalue="Not Completed", text_color="black")
complete_check.grid(row=1, column=3)


for widget in course_taken_frame.winfo_children():
  widget.grid_configure(padx=10, pady=5)

# Import Data 
button = customtkinter.CTkButton(frame, text="Import Data", command= enter_data)
button.grid(row=3, column=0, sticky="news", padx=20, pady=20) 


app.mainloop() 