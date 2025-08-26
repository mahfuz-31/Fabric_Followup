from tkinter import *
from tkinter.ttk import Progressbar
import spc_search_color_wise
from tkinter import messagebox

root = Tk()
root.title('Mahfuz\'s Fabric Followup System')
w = Label(root, text='Mahfuz\'s Fabric Followup System', fg='gray', font=("Playfair Display", 16, 'bold'))
w.pack(pady=5)

frs_label = Label(root, text="Enter your FRS No list:", font=("Calibri", 11))
frs_label.place(x=5, y=40)
# Create a frame to hold text + scrollbar together
text_frame = Frame(root)
text_frame.place(x=5, y=65)

# Text widget
text_area = Text(text_frame, wrap=WORD, width=20, height=25, font=("Arial", 11))
text_area.pack(side=LEFT, fill=BOTH, expand=True)

# Add a Scrollbar
# Scrollbar (placed to the right of text)
scrollbar = Scrollbar(text_frame, command=text_area.yview)
scrollbar.pack(side=RIGHT, fill=Y)

text_area.config(yscrollcommand=scrollbar.set)

progress = Progressbar(root, orient=HORIZONTAL, length=500, mode='determinate')
def on_download():
    spc_search_color_wise.spc_search_color_wise(
        root, progress, text_area.get("1.0", END).strip().split('\n')
    )
    messagebox.showinfo("Completed", "Color wise fabric data download completed!")
    text_area.delete("1.0", END)

download_btn = Button(
    root, 
    text='Download Data', 
    bg='green', 
    fg='white',
    activebackground='lightgreen', 
    activeforeground='white',
    command=on_download
)
download_btn.pack(side=LEFT, anchor='sw', pady=20, padx=20)


root.geometry('800x600')
root.mainloop()