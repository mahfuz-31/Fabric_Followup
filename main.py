from tkinter import *
from tkinter.ttk import Progressbar
import spc_search_color_wise
from tkinter import messagebox
import threading
import fabric_data_window

root = Tk()
root.title('Mahfuz\'s Special Followup System')
w = Label(root, text='Mahfuz\'s Special Followup System', fg='#054270', font=("Playfair Display", 16, 'bold'))
w.pack(pady=5)

fabricL = Label(root, text="Fabric Status Download", fg='#2B6936', font=("Calibri", 12, 'bold'))
fabricL.place(x=35, y=100)
frs_label = Label(root, text="Enter your list of FRS No:", fg='#2B6936', font=("Calibri", 9, 'italic'))
frs_label.place(x=35, y=120)

# Create a frame to hold text + scrollbar together
text_frame = Frame(root)
text_frame.place(x=35, y=145)

# Text widget
text_area = Text(text_frame, wrap=WORD, width=20, height=23, font=("Arial", 11))
text_area.pack(side=LEFT, fill=BOTH, expand=True)

# Add a Scrollbar
scrollbar = Scrollbar(text_frame, command=text_area.yview)
scrollbar.pack(side=RIGHT, fill=Y)
text_area.config(yscrollcommand=scrollbar.set)

# sewing area
sewingL = Label(root, text="Sewing Status Download", fg='#2B6469', font=("Calibri", 12, 'bold'))
sewingL.place(x=280, y=100)
sfrs_label = Label(root, text="Enter your list of FRS No:", fg='#2B6469', font=("Calibri", 9, 'italic'))
sfrs_label.place(x=280, y=120)
# Create a frame to hold text + scrollbar together
stext_frame = Frame(root)
stext_frame.place(x=280, y=145)
# Text widget
stext_area = Text(stext_frame, wrap=WORD, width=20, height=23, font=("Arial", 11))
stext_area.pack(side=LEFT, fill=BOTH, expand=True)
# Add a Scrollbar
sscrollbar = Scrollbar(stext_frame, command=stext_area.yview)
sscrollbar.pack(side=RIGHT, fill=Y)
stext_area.config(yscrollcommand=sscrollbar.set)

# Progress bar
progress = Progressbar(root, orient=HORIZONTAL, length=500, mode='determinate')


def download_thread():
    """Function to run in separate thread"""
    try:
        # Update UI to show processing state
        root.after(0, lambda: download_btn.config(state='disabled', text='Processing...'))
        
        # Get the orders from text area
        orders = text_area.get("1.0", END).strip().split('\n')
        
        # Run the main function
        spc_search_color_wise.spc_search_color_wise(root, progress, orders)
        
        # Success - update UI on main thread
        root.after(0, lambda: messagebox.showinfo("Completed", "Color wise fabric data download completed!"))
        root.after(0, lambda: text_area.delete("1.0", END))
        
    except Exception as e:
        # Error handling - update UI on main thread
        root.after(0, lambda: messagebox.showerror("Error", f"An error occurred: {str(e)}"))
    
    finally:
        # Re-enable button and reset text
        root.after(0, lambda: download_btn.config(state='normal', text='Download Data'))
        root.after(0, cancel_btn.pack_forget)

def on_download():
    global cancel_btn  # make it accessible later to hide
    cancel_btn = Button(
        root, 
        text='Cancel', 
        bg='red', 
        fg='white',
        activebackground='lightcoral', 
        activeforeground='white',
        command=cancel_operation
    )
    cancel_btn.pack(side=LEFT, anchor='sw', pady=20, padx=5)

    """Start download in separate thread"""
    if text_area.get("1.0", END).strip():
        thread = threading.Thread(target=download_thread, daemon=True)
        thread.start()
    else:
        messagebox.showwarning("Warning", "Please enter some FRS numbers!")
        # only hide if warning
        cancel_btn.pack_forget()


download_btn = Button(
    root, 
    font=('Calibri', 9, 'bold'),
    text='Download Data', 
    bg='#2E763A',     
    fg='white',
    activebackground='lightgreen', 
    activeforeground='white',
    command=on_download
)
download_btn.pack(side=LEFT, anchor='sw', pady=20, padx=20)

# Add a cancel button (optional)
def cancel_operation():
    """Note: This is a simple implementation. For true cancellation, 
    you'd need to modify the spc_search_color_wise function"""
    download_btn.config(state='normal', text='Download Data')
    progress.pack_forget()

view_data_btn = Button(root, text='View Fabric Data', font=('Calibri', 9, 'bold'), bg='#3B974A', fg='white',
                       activebackground='lightblue', activeforeground='white', command=lambda: fabric_data_window.fabric_data_window(root))
view_data_btn.pack(side=LEFT, anchor='se', pady=20, padx=0)
root.geometry('500x600')
root.mainloop()