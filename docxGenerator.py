from docxtpl import DocxTemplate
import tkinter as tk
from tkinter import simpledialog

ROOT = tk.Tk()
doc = DocxTemplate("template.docx")

ROOT.withdraw()
# the input dialog
trim_number = simpledialog.askstring(title="Test",
                                  prompt="What's your trimester number?")

company_name = simpledialog.askstring(title="Test",
                                  prompt="What's your company name?")
                                
title = simpledialog.askstring(title="Test",
                                  prompt="What's your title?")

teacher_name = simpledialog.askstring(title="Test",
                                  prompt="What's your teacher name?")

total_grade = simpledialog.askstring(title="Test",
                                  prompt="What's your total grade?")

context = { 'trimester_number' : trim_number , 
            'company_name' : company_name, 
            'title' : title, 
            'teacher_name' : teacher_name, 
            'total_grade' : total_grade}

doc.render(context)
doc.save("generated_doc.docx")

##Automatic-Sendi