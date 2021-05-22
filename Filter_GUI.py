import tkinter as tk
from profile_filter import Profile_filter
from tkinter import filedialog
from tkinter.ttk import Progressbar
import os
import threading

HEIGHT = 800
WIDTH = 1000

def get_result(keys):
    xl_frame.place_forget()
    if len(keys) <2 or keys == 'Enter the search keys here separated by comma':
        text_area.delete(1.0, tk.END)
        text_area.insert(tk.INSERT, "Please provide correct keys like python, java")
        return
    out_xl = xl_box.get()

    keys = keys.split(',')
    pattern = [i.strip().lower() for i in keys]
    path = label_folder['text']
    if path == "Please select folder or a file":
        text_area.delete(1.0, tk.END)
        text_area.insert(tk.INSERT, "Please select a folder or a file before search")
        return

    progress.place(relx=0.5, rely=0.2, anchor='n')
    pf = Profile_filter(pattern,path)
    pf.xl = out_xl
    pf.filter_filenames()
    t1 = threading.Thread(target=pf.start_profile_filter, kwargs={'filter_file':False})
    t1.start()
    start_progress_cp(pf, progress)
    t1.join()

    table = pf.table
    text_area.configure(state='normal')
    text_area.delete(1.0, tk.END)
    text_area.insert(tk.INSERT, table)
    text_area.configure(state='disabled')
    if out_xl:
        xl_frame_place(pf)
    progress.place_forget()


def start_progress_cp(pf_obj, progress_obj):
    total = len(pf_obj.Profile_Files_pdf) + len(pf_obj.Profile_Files_docx)
    progress_val = 0
    progress_obj['value'] = progress_val
    root.update_idletasks()
    print("Brfore Loop", total)
    while True:
        count=pf_obj.progress_count
        if total:
            progress_val = round((count/total)*100,2)
        if progress['value'] == progress_val:
            continue
        else:
            progress['value'] = progress_val
            root.update_idletasks()
            print(f"Progress:{progress_val}%,File count: {count}")
        if total == count:
            break
    progress['value'] = 100
    root.update_idletasks()

def xl_frame_place(pf_obj):
    xl_path = pf_obj.output_file_xl
    xl_frame.place(relx=0.445, rely=0.15, relwidth=0.64, relheight=0.05, anchor='n')
    button_xl = tk.Button(xl_frame, text="open result", command=lambda: os.startfile(xl_path))
    button_xl.place(relx=0.82, relheight=1,relwidth=0.18)

    label_xl = tk.Label(xl_frame, text=xl_path, fg='green')
    label_xl.place(relheight=1, relwidth=0.82)

def get_folder():
    root.directory = filedialog.askdirectory()
    dir = root.directory
    if len(dir) > 1:
        dir = dir.replace('/', '\\')
        label_folder['text'] = dir
        label_folder['fg'] = 'green'

def get_file():
    file = filedialog.askopenfile(initialdir='/', title="Select file", filetypes=(("Word Document", ".docx .pdf"),
                                                                                  ("all files", "*.*")))
    if len(file)>1 :
        file= file.replace('/', '\\')
        label_folder['text'] = file
        label_folder['fg'] = 'green'

######MAIN############

root = tk.Tk()
root.title("Profile Filter")
canvas = tk.Canvas(root, height=HEIGHT, width=WIDTH)
canvas.pack()

dir = str(__file__)
dir = (dir.split(os.path.basename(dir))[0])

icon_img = tk.PhotoImage(file= dir.replace('/','\\')+ 'icon.png')
root.iconphoto(False, icon_img)
background_image = tk.PhotoImage(file= dir.replace('/','\\')+ 'blue1.png')
background_label = tk.Label(root, image=background_image)
background_label.place(relwidth=1, relheight=1)

##############
path_frame = tk.Frame(root)
path_frame.place(relx=0.5, rely=0.05, relwidth=0.75, relheight=0.05, anchor='n')

button_folder = tk.Button(path_frame, text="Select Folder", command=get_folder)
button_folder.place(relx=0.85, relheight=1, relwidth=0.15)
button_file = tk.Button(path_frame, text="Select File", command=get_file)
button_file.place(relx=0.70, relheight=1, relwidth=0.15)

label_folder = tk.Label(path_frame, text='Please select folder or file', fg='red', font=8 )
label_folder.place(relheight=1, relwidth=0.70)

#################

key_frame = tk.Frame(root)
key_frame.place(relx=0.5, rely=0.1, relwidth=0.75, relheight=0.05, anchor='n')

entry = tk.Entry(key_frame, font=40, bd=2)
entry.place(relwidth=0.85, relheight=1)
entry.insert(0, "Enter the search keys here separated by comma")

xl_box = tk.IntVar()
checkbox_frame = tk.Frame(root, bg='#0d1a26', bd=1)
checkbox_frame.place(relx=0.82, rely=0.15, relwidth=0.11, relheight=0.05, anchor='n')
checkbutton = tk.Checkbutton(checkbox_frame, text="Excel_Result", variable=xl_box, bd=2)
checkbutton.place(relx=0.4, rely=-0.1, relwidth=1.2, relheight=1.1, anchor='n')

progress = Progressbar(root, orient=tk.HORIZONTAL, length=100, mode='determinate')

xl_frame= tk.Frame(root, bg='#669999', bd=1)

###############

result_frame = tk.Frame(root, bg='#3973ac', bd=2)
result_frame.place(relx=0.5, rely=0.25, relwidth=0.75, relheight=0.6, anchor='n')

h= tk.Scrollbar(result_frame, orient= 'horizontal')
h.pack(side=tk.BOTTOM, fill=tk.X)
v=tk.Scrollbar(result_frame)
v.pack(side=tk.RIGHT, fill=tk.Y)
text_area = tk.Text(result_frame, width=40, height=40, font= ('Consolas', 10), wrap = tk.NONE, xscrollcommand=h.set,
                    yscrollcommand=v.set)
text_area.place(relwidth=1, relheight=1)
text_area.pack(side=tk.TOP, fill=tk.X)
h.config(command=text_area.xview)
v.config(command=text_area.yview)

button_search = tk.Button(key_frame, text="Search Profiles", command=lambda: get_result(entry.get()))
button_search.place(relx=0.85, relheight=1, relwidth=0.15)
#quit_button
tk.Button(root, text='Quit', command=root.quit).place(relx=0.85, rely=0.85, relwidth=0.1, relheight=0.05, anchor='n')

root.mainloop()
