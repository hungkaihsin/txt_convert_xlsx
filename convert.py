import aspose.cells 
from aspose.cells import Workbook
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
from tkinter.messagebox import showinfo

window = tk.Tk()
window.title("Convert")
window.geometry("250x100")
window.resizable(False, False)
window.iconbitmap("icon.ico")

En_file_path = ttk.Entry(width=28)
En_file_path.place(x=0, y=2)
En_save_path = ttk.Entry()


def load_file():
    file_route = filedialog.askopenfilename(filetypes=(("txt", "*.txt"), ("all", "*.*")), title="Open file")
    En_file_path.insert(0, file_route)



def save_location():
    if En_save_path.get() is None:
        save_path = filedialog.askdirectory()
        En_save_path.insert(0, save_path)
    else:
        save_path = filedialog.askdirectory()
        En_save_path.delete(0, "end")
        En_save_path.insert(0, save_path)

        
def convert():
    if len(En_file_path.get()) == 0:
        showinfo(title="Error", message= "Please select file")
    elif len(En_save_path.get()) == 0:
        showinfo(title="Error", message= "Please choose save location")

    workbook = Workbook(En_file_path.get())
    save_destination = En_save_path.get()
    save_route = save_destination + "\Output.xlsx"
    workbook.save(save_route)
    showinfo(title= "Done", message="Finish!")

Btn_choose_file = ttk.Button(text="Select file", command= load_file)
Btn_choose_file.place(x=175, y=0)
Btn_convert = ttk.Button(text="Convert", command= convert)
Btn_convert.place(x=150, y=60)
Btn_save_loaction = ttk.Button(text="Save location", command= save_location)
Btn_save_loaction.place(x=30, y=60)


window.mainloop()






