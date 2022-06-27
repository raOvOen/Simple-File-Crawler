import os
import tkinter.messagebox

import docx
import pandas as pd
import tkinter as tk
import threading
import tkinter.filedialog as fd
import tkinter.ttk as ttk
from datetime import datetime

stopped = False
res_adress = []
res_info = []

def block_widgets(state):
    if state:
        btn_start["state"] = tk.DISABLED
        btn_stop["state"] = tk.NORMAL
        btn_ch_dir["state"] = tk.DISABLED
        btn_add_keyword["state"] = tk.DISABLED
        btn_clear["state"] = tk.DISABLED
        btn_clear_keywords["state"] = tk.DISABLED
        entry_add_keyword["state"] = tk.DISABLED
        entry_dir_path["state"] = tk.DISABLED
        checkbox_reg_check["state"] = tk.DISABLED
        checkbox_check_excel["state"] = tk.DISABLED
    else:
        btn_start["state"] = tk.NORMAL
        btn_stop["state"] = tk.DISABLED
        btn_ch_dir["state"] = tk.NORMAL
        btn_add_keyword["state"] = tk.NORMAL
        btn_clear["state"] = tk.NORMAL
        btn_clear_keywords["state"] = tk.NORMAL
        entry_add_keyword["state"] = tk.NORMAL
        entry_dir_path["state"] = tk.NORMAL
        checkbox_reg_check["state"] = tk.NORMAL
        checkbox_check_excel["state"] = tk.NORMAL

def btn_ch_dir_command():
    directory = fd.askdirectory(title="Choose directory", initialdir="/")
    if directory:
        entry_dir_path.delete(0,tk.END)
        entry_dir_path.insert(0,directory)

def btn_add_keyword_command():
    if entry_add_keyword.get() != "":
        listbox_keywords.insert(tk.END,entry_add_keyword.get())
        entry_add_keyword.delete(0, tk.END)

def btn_start_command():
    global stopped
    stopped = False
    block_widgets(True)
    listbox_result.delete(0, tk.END)
    th = threading.Thread(target=thread_start_command)
    th.daemon = True
    th.start()

def btn_stop_command():
    global stopped
    stopped = True

def btn_cp_results_command():
    if listbox_result.size() > 0:
        label_cp_path["text"] = "All results were successfully copied!"
        label_cp_path["justify"]=tk.LEFT
        label_cp_path["fg"] = 'purple'
        result_of_scan = ""
        for i,text in enumerate(listbox_result.get(0, tk.END)):
            result_of_scan = result_of_scan + str(text) + "\n"
        result_of_scan = result_of_scan.rstrip()
        root.clipboard_clear()
        root.clipboard_append(result_of_scan)
        root.update()
        label_cp_path.place(x=240, y=551, height=30)

def listbox_del(event):
    if btn_add_keyword["state"] == tk.NORMAL:
        selected = list(listbox_keywords.curselection())
        selected.reverse()
        for i in selected:
            listbox_keywords.delete(i)

def listbox_open(event):
    if listbox_result.size() > 0:
        temp_str = listbox_result.selection_get()
        temp_pos = listbox_result.curselection()
        if temp_str[0][0] == 'M':
            os.startfile(res_adress[temp_pos[0]-2])

def listbox_copy_path(event):
    if listbox_result.size() > 0:
        label_cp_path["text"] = "The selected path was successfully copied!"
        label_cp_path["justify"]=tk.LEFT
        label_cp_path["fg"] = 'green'
        temp_pos = listbox_result.curselection()
        root.clipboard_clear()
        root.clipboard_append(res_adress[temp_pos[0]-2])
        root.update()
        label_cp_path.place(x=240, y=551, height=30)

def btn_clear_command():
    listbox_keywords.delete(0,tk.END)
    listbox_result.delete(0,tk.END)
    label_mentioned["text"] = "Ready to Start!"
    label_cp_path["text"] = ""
    global res_adress
    res_adress = []
    global res_info
    res_info = []

def btn_clear_keywords_command():
    listbox_keywords.delete(0,tk.END)

def entry_search_command(event):
    val = event.widget.get()
    if val == '':
        data = res_info
    else:
        data = []
        for item in res_info:
            if val.lower() in item.lower():
                data.append(item)

    search_update(data)

def search_update(data):
    listbox_result.delete(0, 'end')
    for item in data:
        listbox_result.insert('end', item)

def btn_help_command():
    tkinter.messagebox.showinfo(title="About",message="Simple Doc and Xls Crawler\n"+
                                                      "\n"+
                                                      "This program was designed to search for the required information (keywords) in Word and Excel documents.\n"+
                                                      "\n"+
                                                      "How to use the program:\n"+
                                                      "1. Specify or select, using the corresponding button, the root directory of the scan.\n"+
                                                      "2. Using a special window, specify the necessary keywords which will be further searched in the documents. If necessary, you can set auxiliary flags to add phrases of different case and/or search in Excel documents. If any keyword was added unintentionally - you can delete it by double-clicking it.\n"+
                                                      "3. Clicking the \"Start\" button will start the scanning process. After its completion, you can select the desired item to copy it or open it - by double clicking on it.\n"+
                                                      "If necessary, you can copy all the results of the scan.\n"+
                                                      "\n"+
                                                      "Made by raOvOen")

def thread_start_command():
    label_mentioned["text"] = "Scanning..."
    progress_bar_search = ttk.Progressbar(root, orient="horizontal", mode="determinate", value=0)
    path = os.walk(entry_dir_path.get())
    cnt_dirs = count_dirs(path)
    progress_bar_search["maximum"] = cnt_dirs
    progress_bar_search.place(x=80, y=556, height=20)
    ins_text = listbox_keywords.get(0,tk.END)
    global stopped
    text = []
    global res_adress
    res_adress = []
    global res_info
    res_info = []
    for i in ins_text:
        text.append(i)
    path = os.walk(entry_dir_path.get())
    now = datetime.now()
    current_time = now.strftime("%H:%M:%S")
    temp_str = f'[{current_time}] Start scanning'
    res_info.append(temp_str)
    listbox_result.insert(tk.END, temp_str)
    temp_str = 'Keywords: '+str(text)
    res_info.append(temp_str)
    listbox_result.insert(tk.END, temp_str)
    for adress, dirs, files in path:
        for file in files:
            if not stopped:
                lookfor_text(adress, file, check_excel.get(),text)
                progress_bar_search['value'] += 1
    now = datetime.now()
    current_end_time = now.strftime("%H:%M:%S")
    if not stopped:
        temp_str = f'[{current_end_time}] Scanning Finished'
        listbox_result.insert(tk.END,temp_str)
        res_info.append(temp_str)
    else:
        temp_str = f'[{current_end_time}] Scanning Stopped'
        listbox_result.insert(tk.END, temp_str)
        res_info.append(temp_str)
    progress_bar_search.stop()
    progress_bar_search.destroy()
    label_mentioned["text"] = f"The keywords are mentioned {len(res_adress)} time(s)!"
    block_widgets(False)

def count_dirs(path):
    counter = 0
    for adress, dirs, files in path:
        counter += len(files)
    return counter

def lookfor_text(adr, filename, excel, text):
    file_ext = os.path.splitext(filename)
    if (file_ext[-1] == ".docx" or file_ext[-1] == ".doc"):
        count = 0
        full_adr = adr + os.sep + filename
        try:
            doc_file = docx.Document(full_adr)
            all_paras = doc_file.paragraphs
            act_keyword = "|"
            for para in all_paras:
                for keyword in text:
                    if check_reg.get():
                        if keyword.lower() in para.text.lower():
                            count += 1
                            if keyword not in act_keyword:
                                act_keyword += keyword + "|"
                    else:
                        if keyword in para.text:
                            count += 1
                            if keyword not in act_keyword:
                                act_keyword += keyword + "|"
        except (Exception):
            pass
        finally:
            if count > 0:
                temp_str = 'Mentioned: ' + act_keyword + ' for ' + str(count) + ' times! [' + full_adr + ']'
                if entry_search.get != "" and entry_search.get() in temp_str:
                    listbox_result.insert(tk.END, temp_str)
                res_info.append(temp_str)
                res_adress.append(full_adr)
    if excel:
        if (file_ext[-1] == ".xlsx" or file_ext[-1] == ".xls"):
            count = 0
            full_adr = adr + os.sep + filename
            try:
                xls_file = pd.read_excel(full_adr)
                xls_text = xls_file.head()
                act_keyword = "|"
                for temp_text in xls_text:
                    for keyword in text:
                        if check_reg.get():
                            if keyword.lower() in temp_text.lower():
                                count += 1
                                if keyword not in act_keyword:
                                    act_keyword += keyword + "|"
                        else:
                            if keyword in temp_text:
                                count += 1
                                if keyword not in act_keyword:
                                    act_keyword += keyword + "|"
            except (Exception):
                pass
            finally:
                if count > 0:
                    temp_str = 'Mentioned: ' + act_keyword + ' for ' + str(count) + ' times! [' + os.path.join(adr, filename) + ']'
                    if entry_search.get != "" and entry_search.get() in temp_str:
                        listbox_result.insert(tk.END, temp_str)
                    listbox_result.insert(tk.END, temp_str)
                    res_adress.append((full_adr))
                    res_info.append(temp_str)

if __name__ == '__main__':
    root = tk.Tk()
    root.title("Simple Doc and Xls Crawler by raOvOen")
    root.resizable(width=False, height=False)
    window_width = 800
    window_height = 600
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    x_cordinate = int((screen_width / 2) - (window_width / 2))
    y_cordinate = int((screen_height / 2) - (window_height / 2))
    root.geometry("{}x{}+{}+{}".format(window_width, window_height, x_cordinate, y_cordinate))

    btn_ch_dir = tk.Button(root)
    btn_ch_dir["text"] = "Choose directory"
    btn_ch_dir.place(x=10, y=20, width=117, height=30)
    btn_ch_dir["command"] = btn_ch_dir_command

    entry_dir_path = tk.Entry(root)
    entry_dir_path["text"] = "Dir"
    entry_dir_path.place(x=140, y=20, width=600, height=30)

    listbox_keywords = tk.Listbox(root)
    listbox_keywords.place(x=320, y=60, width=470, height=80)
    listbox_keywords.bind('<Double-Button-1>', listbox_del)

    btn_add_keyword = tk.Button(root)
    btn_add_keyword["text"] = "Add keyword"
    btn_add_keyword.place(x=140, y=105, width=170, height=35)
    btn_add_keyword["command"] = btn_add_keyword_command

    btn_clear = tk.Button(root)
    btn_clear["text"] = "Clear all"
    btn_clear.place(x=140, y=148, width=80, height=32)
    btn_clear["command"] = btn_clear_command

    btn_clear_keywords = tk.Button(root)
    btn_clear_keywords["text"] = "Clear keys"
    btn_clear_keywords.place(x=230, y=148, width=80, height=32)
    btn_clear_keywords["command"] = btn_clear_keywords_command

    entry_add_keyword = tk.Entry(root)
    entry_add_keyword["text"] = "Entry"
    entry_add_keyword.place(x=140, y=63, width=170, height=32)

    btn_start = tk.Button(root)
    btn_start["text"] = "Start"
    btn_start.place(x=25, y=116, width=88, height=32)
    btn_start["command"] = btn_start_command

    btn_stop = tk.Button(root)
    btn_stop["text"] = "Stop"
    btn_stop.place(x=25, y=154, width=88, height=32)
    btn_stop["state"] = tk.DISABLED
    btn_stop["command"] = btn_stop_command

    check_reg = tk.BooleanVar()
    check_reg.set(True)
    checkbox_reg_check = tk.Checkbutton(root)
    checkbox_reg_check["text"] = "Simple Reg check"
    checkbox_reg_check.place(x=0, y=55, width=140, height=30)
    checkbox_reg_check["variable"] = check_reg
    checkbox_reg_check["offvalue"] = False
    checkbox_reg_check["onvalue"] = True

    check_excel = tk.BooleanVar()
    check_excel.set(True)
    checkbox_check_excel = tk.Checkbutton(root)
    checkbox_check_excel["text"] = "Check excel"
    checkbox_check_excel.place(x=0, y=80, width=110, height=30)
    checkbox_check_excel["variable"] = check_excel
    checkbox_check_excel["offvalue"] = False
    checkbox_check_excel["onvalue"] = True

    scrollbar = tk.Scrollbar(root, orient="horizontal")
    scrollbar.pack(side="bottom",fill="x")
    listbox_result = tk.Listbox(root,xscrollcommand=scrollbar.set,height=20)
    listbox_result.place(x=10, y=190, width=780, height=360)
    listbox_result.bind('<Double-Button-1>', listbox_open)
    listbox_result.bind('<<ListboxSelect>>',listbox_copy_path)
    scrollbar.config(command=listbox_result.xview)

    btn_help = tk.Button(root)
    btn_help["text"] = "?"
    btn_help.place(x=750, y=10, width=40, height=40)
    btn_help["command"] = btn_help_command

    btn_cp_results = tk.Button(root)
    btn_cp_results["text"] = "Copy all results"
    btn_cp_results.place(x=670, y=551, width=120, height=30)
    btn_cp_results["command"] = btn_cp_results_command

    label_mentioned = tk.Label(root)
    label_mentioned["text"] = "Ready to Start!"
    label_mentioned["justify"]=tk.LEFT
    label_mentioned.place(x=10, y=551, height=30)

    label_cp_path = tk.Label(root)

    entry_search = tk.Entry(root)
    entry_search.place(x=375, y=150, width=415, height=30)
    entry_search.bind('<KeyRelease>', entry_search_command)

    label_search = tk.Label(root)
    label_search["text"] = "Search:"
    label_search["justify"]=tk.LEFT
    label_search.place(x=325, y=150, height=30)

    stopped = False
    root.mainloop()
