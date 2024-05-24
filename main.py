import re
import time
import tkinter as tk
from tkinter import *
from tkinter import filedialog
from tkinter import ttk
import pandas as pd
import os
import subprocess
import threading
from tkinter import messagebox

import xlsxwriter
from PIL import Image, ImageTk

root = tk.Tk()
root.title("Slwave Automation")

noteframe = tk.Frame(root)
mainframe = tk.Frame(root)

logmessageframe = tk.Frame(root, bg="white", borderwidth=2, relief="groove")

canvas = tk.Canvas(logmessageframe, bg="white", borderwidth=2, relief="groove")
canvas.pack(side="left", fill="both", expand=True)

scrollbar = tk.Scrollbar(logmessageframe, orient="vertical", command=canvas.yview)
scrollbar.pack(side="right", fill="y")

canvas.configure(yscrollcommand=scrollbar.set)
canvas.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))

inner_frame = tk.Frame(canvas, bg="white")
inner_frame.pack(fill="both", expand=True)

canvas.create_window((0, 0), window=inner_frame, anchor="nw")
logmes = tk.Label(inner_frame, text="Log message", font=("Arial", 10), bg="white")
logmes.pack(pady=3, padx=2, anchor='nw')

imageframe = tk.Frame(root)
Slwave_file_path = ""
envVars = os.environ
ansysVers = {}
version_selected = ""
version_pathof_exe = ""
Excel_file_path = ""
checktrue = True


def notewin():
    root.geometry("400x200")

    label = tk.Label(noteframe, text='Note!', font=("Arial", 13))
    label.pack(pady=2)

    stri = "SIwave file should be obtained from Cadence’s Ansys Auto,\n please make sure only power rails selected for the simulation\n are included (i.e. user will be prompted to select desired \npower rails in Ansys Auto, to include in SIwave file)"
    label2 = tk.Label(noteframe, text=stri, font=("Arial", 10))
    label2.pack(pady=5)

    but = tk.Button(noteframe, text="Start", font=("Arial", 10), fg="black", height=2, width=20,
                    command=lambda: gotomain())
    but.pack(pady=3)


def gotomain():
    if noteframe.winfo_ismapped():
        noteframe.pack_forget()
        mainframe12()
        mainframe.pack(side="top", fill="both", expand=True)
        logmessageframe.pack(side="bottom", fill="both", expand=True)


def mainframe12():
    global version_pathof_exe, Slwave_file_path, version_selected, ansysVers, checktrue

    root.geometry("410x500")
    root.maxsize(500, 500)
    for ver in envVars:
        if ver.startswith('ANSYSEM'):
            ansyspath = str(os.environ[ver])
            verName = re.search(r'(\d{3}$)', ver)
            verName = "AnsysEM" + verName[0]
            if verName == '':
                continue
            verName = str(verName)
            ansysVers[verName] = ansyspath
    verlist = list(ansysVers.keys())
    verlist = ["ans", "ques"]
    print(verlist)
    print(type(verlist[0]))
    verpath = list(ansysVers.values())

    label1 = tk.Label(mainframe, text="      Slwave Automation", font=("Arial", 10))
    label1.grid(row=0, column=2, pady=10, sticky="n", columnspan=2, padx=100)

    label1 = tk.Label(mainframe, text="      Ansys Slwave version", font=("Arial", 10))
    label1.grid(row=1, column=2, pady=5)

    selected_option = tk.StringVar(value=verlist[0])
    dropdown = tk.OptionMenu(mainframe, selected_option, *verlist,
                             command=lambda value: on_select(value, selected_option))
    dropdown.grid(row=1, column=3)

    # if checktrue:
    #     version_pathof_exe=ansysVers[verlist[0]]
    #     print("path", version_pathof_exe)

    label1 = tk.Label(mainframe, text="      Slwave File               ", font=("Arial", 10))
    label1.grid(row=2, column=2, pady=5)

    upload_button = tk.Button(mainframe, text="...", command=upload, width=5)
    upload_button.grid(row=2, column=3, pady=5, padx=30)

    run_button = tk.Button(mainframe, text="Run Automation", command=runautomation, width=15)
    run_button.grid(row=3, column=2, pady=5, padx=10, columnspan=2)

    canvas = tk.Canvas(mainframe, bg='gray', height=1, width=370)
    # Draw a horizontal line from (50, 100) to (150, 100)
    canvas.create_line(50, 100, 150, 100)
    canvas.grid(row=4, column=2, columnspan=2, padx=20)

    label1 = tk.Label(mainframe, text="Excel Configuration input                                        ",
                      font=("Arial", 10))
    label1.grid(row=5, column=2, pady=5, columnspan=2)

    var1 = tk.StringVar()
    var1.set('option1')

    style = ttk.Style()
    style.configure("TRadiobutton", background=mainframe.cget('background'),
                    foreground='black', borderwidth=0, focuscolor=mainframe.cget('background'))

    rb1 = ttk.Radiobutton(mainframe, text="    Select excel file (already in correct format).      ", variable=var1,
                      value='option1', command=on_radio_select1,style="TRadiobutton")
    rb1.grid(row=6, column=2, padx=5, columnspan=2)

    rb1 = ttk.Radiobutton(mainframe, text="    No pre-set excel configuration, input now.      ", variable=var1,
                      value='option2', command=on_radio_select,style="TRadiobutton")
    rb1.grid(row=7, column=2, columnspan=2)

    canvas = tk.Canvas(mainframe, bg='gray', height=1, width=370)
    # Draw a horizontal line from (50, 100) to (150, 100)
    canvas.create_line(50, 100, 150, 100)
    canvas.grid(row=8, column=2, columnspan=2, padx=20)

    reset_button = tk.Button(mainframe, text="Reset", command=restart, width=15)
    reset_button.grid(row=9, column=3, pady=10)

    genrate_button = tk.Button(mainframe, text="Generate Report", command=genratepptx, width=15)
    genrate_button.grid(row=9, column=2, pady=10)


def runautomation():
    if not Slwave_file_path:
        print("slwave file path",Slwave_file_path)
        logmesafe1 = tk.Label(inner_frame, text="No Slwave file selected", font=("Arial", 10), bg="white")
        logmesafe1.pack(pady=2, padx=2, anchor='nw')
        canvas.update_idletasks()
        canvas.configure(scrollregion=canvas.bbox("all"))
        return

    logmesafe1 = tk.Label(inner_frame, text="Running Autosim python file", font=("Arial", 10), bg="white")
    logmesafe1.pack(pady=2, padx=2, anchor='nw')
    logmesafe1 = tk.Label(inner_frame, text="Please waite while execution complete!", font=("Arial", 10), bg="white")
    logmesafe1.pack(pady=2, padx=2, anchor='nw')

    # result = runcommand()
    # logmesafe1 = tk.Label(inner_frame, text="Execution completed please continue now!", font=("Arial", 10), bg="white")
    # logmesafe1.pack(pady=2, padx=2, anchor='nw')

    def runcommand_thread():
        process = runcommand()
        process.wait()
        logmesafe1 = tk.Label(inner_frame, text="Auto simulation execution completed. Please input excel configuration",
                              font=("Arial", 10),
                              bg="white")
        logmesafe1.pack(pady=2, padx=2, anchor='nw')


        logmesafe1 = tk.Label(inner_frame, text="Slwave file Closed", font=("Arial", 10), bg="white")
        logmesafe1.pack(pady=2, padx=2, anchor='nw')
        canvas.update_idletasks()
        canvas.configure(scrollregion=canvas.bbox("all"))


    thread = threading.Thread(target=runcommand_thread)
    thread.start()

    # if result.returncode != 0:
    #     logmesafe1 = tk.Label(inner_frame, text="Error while running Autosim python file", font=("Arial", 10),
    #                           bg="white")
    #     logmesafe1.pack(pady=2, padx=2, anchor='nw')
    #
    #     canvas.update_idletasks()
    #     canvas.configure(scrollregion=canvas.bbox("all"))
    #     return

    # logmesafe1 = tk.Label(inner_frame, text="Autosim python file executed successfully", font=("Arial", 10), bg="white")
    # logmesafe1.pack(pady=2, padx=2, anchor='nw')

    canvas.update_idletasks()
    canvas.configure(scrollregion=canvas.bbox("all"))


def genratepptx():
    global Slwave_file_path, Excel_file_path
    if not Excel_file_path:
        logmesafe1 = tk.Label(inner_frame, text="No Excel file selected or Excel file still opened", font=("Arial", 10), bg="white")
        logmesafe1.pack(pady=2, padx=2, anchor='nw')
        return

    file_name = "post_process.py"
    pytonfilepath = os.path.join(os.getcwd(), file_name)
    logmesafe1 = tk.Label(inner_frame, text="Running post process python file", font=("Arial", 10), bg="white")
    logmesafe1.pack(pady=2, padx=2, anchor='nw')
    parent_dir = os.path.dirname(Slwave_file_path)
    print("parent_dir: ",parent_dir)
    result = subprocess.run(['python', pytonfilepath, Excel_file_path,parent_dir])
    if result.returncode != 0:
        logmesafe1 = tk.Label(inner_frame, text="Error while running post process python file", font=("Arial", 10),
                              bg="white")
        logmesafe1.pack(pady=2, padx=2, anchor='nw')

        canvas.update_idletasks()
        canvas.configure(scrollregion=canvas.bbox("all"))
        return

    logmesafe1 = tk.Label(inner_frame, text="Running of post process python file completed", font=("Arial", 10),
                          bg="white")
    logmesafe1.pack(pady=2, padx=2, anchor='nw')

    logmesafe1 = tk.Label(inner_frame, text="Report file generated", font=("Arial", 10), bg="white")
    logmesafe1.pack(pady=2, padx=2, anchor='nw')
    canvas.update_idletasks()
    canvas.configure(scrollregion=canvas.bbox("all"))


def restart():
    logmesafe1 = tk.Label(inner_frame, text="Restarting", font=("Arial", 10), bg="white")
    logmesafe1.pack(pady=2, padx=2, anchor='nw')
    global Slwave_file_path, envVars, ansysVers, version_pathof_exe, version_selected, checktrue, Excel_file_path
    mainframe.pack_forget()
    for widget in mainframe.winfo_children():
        widget.destroy()
    logmessageframe.pack_forget()

    Slwave_file_path = ""
    envVars = os.environ
    ansysVers = {}
    version_selected = ""
    version_pathof_exe = ""
    checktrue = True
    Excel_file_path = ""

    mainframe12()
    mainframe.pack(side="top", fill="both", expand=True)
    logmessageframe.pack(side="bottom", fill="both", expand=True)
    logmesafe1 = tk.Label(inner_frame, text="Restarting complete", font=("Arial", 10), bg="white")
    logmesafe1.pack(pady=2, padx=2, anchor='nw')

    canvas.update_idletasks()
    canvas.configure(scrollregion=canvas.bbox("all"))


def on_radio_select1():
    # upload excel file
    uploadexcel()


def on_radio_select():
    createxcel()


def uploadexcel():
    gotodisplay()


def oupload_exclefile():
    global Excel_file_path
    for widget in imageframe.winfo_children():
        widget.destroy()
    try:
        logmesafe1 = tk.Label(inner_frame, text="Uploading Excel file", font=("Arial", 10), bg="white")
        logmesafe1.pack(pady=2, padx=2, anchor='nw')
        file_path = filedialog.askopenfilename(title="Select Excel file",
                                               filetypes=(("Excel files", "*.xlsx"), ("All files", "*.*")))
        logmesafe1 = tk.Label(inner_frame, text="Upload successful", font=("Arial", 10), bg="white")
        logmesafe1.pack(pady=2, padx=2, anchor='nw')

        canvas.update_idletasks()
        canvas.configure(scrollregion=canvas.bbox("all"))
    except:
        logmesafe1 = tk.Label(inner_frame, text="Error while uploading Excel FIle", font=("Arial", 10), bg="white")
        logmesafe1.pack(pady=2, padx=2, anchor='nw')

        canvas.update_idletasks()
        canvas.configure(scrollregion=canvas.bbox("all"))
        return

    if file_path:
        Excel_file_path = file_path

        df = pd.read_excel(Excel_file_path)

        logmesafe1 = tk.Label(inner_frame, text="Checking file format", font=("Arial", 10), bg="white")
        logmesafe1.pack(pady=2, padx=2, anchor='nw')
        if 'Power Rail Name' in df.columns:
            if "Current (A)" in df.columns:
                logmesafe1 = tk.Label(inner_frame, text="File format checked", font=("Arial", 10), bg="white")
                logmesafe1.pack(pady=2, padx=2, anchor='nw')
            else:
                logmesafe1 = tk.Label(inner_frame, text="File format incorrect!", font=("Arial", 10), bg="white")
                logmesafe1.pack(pady=2, padx=2, anchor='nw')

                canvas.update_idletasks()
                canvas.configure(scrollregion=canvas.bbox("all"))
                return
        else:
            logmesafe1 = tk.Label(inner_frame, text="File format incorrect!", font=("Arial", 10), bg="white")
            logmesafe1.pack(pady=2, padx=2, anchor='nw')

            canvas.update_idletasks()
            canvas.configure(scrollregion=canvas.bbox("all"))
            return


    else:
        logmesafe1 = tk.Label(inner_frame, text="No Excel FIle selected", font=("Arial", 10), bg="white")
        logmesafe1.pack(pady=2, padx=2, anchor='nw')

    canvas.update_idletasks()
    canvas.configure(scrollregion=canvas.bbox("all"))


def gotodisplay():
    if mainframe.winfo_ismapped():
        mainframe.pack_forget()
        logmessageframe.pack_forget()
        display_image()


def display_image():
    # Open image file
    root.geometry("500x270")
    image = Image.open("image.png")
    photo = ImageTk.PhotoImage(image)
    # Create a Label widget with the image
    label = tk.Label(imageframe, image=photo)
    label.pack()
    imageframe.pack(side="top", fill="both", expand=True)
    text12 = "Please ensure that Power Rail Names are entered accurately, without spaces before, within, or after the name"
    messagebox.showinfo("Image Pop-up", text12, icon='info', master=imageframe)

    gotouploadexcel()


def gotouploadexcel():
    root.geometry("410x500")
    if imageframe.winfo_ismapped():
        imageframe.pack_forget()
        mainframe.pack(side="top", fill="both", expand=True)
        logmessageframe.pack(side="bottom", fill="both", expand=True)

        oupload_exclefile()


def createxcel():
    global Excel_file_path
    # creating excel
    logmesafe1 = tk.Label(inner_frame, text="Creating Excel file", font=("Arial", 10), bg="white")
    logmesafe1.pack(pady=2, padx=2, anchor='nw')

    # Define column headers
    column_headers = ['Power Rail Name', 'Current (A)']

    # Create a new Excel workbook and add a worksheet
    workbook = xlsxwriter.Workbook('post_process.xlsx')
    worksheet = workbook.add_worksheet()

    # Define cell format for the header row with borders
    header_format = workbook.add_format({
        'border': 1,
        'bold': True,
        'align': 'center',
        'valign': 'vcenter'
    })

    # Write the column headers to the worksheet
    worksheet.write_row(0, 0, column_headers, header_format)

    # Save and close the workbook
    workbook.close()

    logmesafe1 = tk.Label(inner_frame, text="Excel file Created", font=("Arial", 10), bg="white")
    logmesafe1.pack(pady=2, padx=2, anchor='nw')

    logmesafe1 = tk.Label(inner_frame, text="Opening Excel file", font=("Arial", 10), bg="white")
    logmesafe1.pack(pady=2, padx=2, anchor='nw')
    file_name = 'post_process.xlsx'
    file_path = os.path.join(os.getcwd(), file_name)


    def runcommand_thread1():
        process = subprocess.Popen(file_path, shell=True)
        # Wait for the program to finish
        process.wait()
        global Excel_file_path
        Excel_file_path = file_path
        logmesafe1 = tk.Label(inner_frame, text="Excel file closed", font=("Arial", 10), bg="white")
        logmesafe1.pack(pady=2, padx=2, anchor='nw')
        canvas.update_idletasks()
        canvas.configure(scrollregion=canvas.bbox("all"))

    thread = threading.Thread(target=runcommand_thread1)
    thread.start()

def runcommand():
    global version_pathof_exe, Slwave_file_path
    print(version_pathof_exe, "Slwave_file_path ", Slwave_file_path)

    file_name = "autosim.py"
    file1_path = os.path.join(os.getcwd(), file_name)

    version_pathof_exe = version_pathof_exe + "\\siwave.exe"

    parent_dir = os.path.dirname(Slwave_file_path)
    print("parent_dir: ",parent_dir)

    with open("path.txt", 'w') as f:
        f.write(parent_dir)
        f.close()


    cmd_commad1 = f"\"{version_pathof_exe}\" {Slwave_file_path} -RunScriptAndExit {file1_path}"

    try:
        logmesafe1 = tk.Label(inner_frame, text="Opening Slwave file", font=("Arial", 10), bg="white")
        logmesafe1.pack(pady=2, padx=2, anchor='nw')

        print("cmd command 1= ", cmd_commad1)

        result = subprocess.Popen(f'cmd /c {cmd_commad1}', shell=True)


        print("cmd command 1= ", cmd_commad1)

        print(f"Slwave file path ={Slwave_file_path}")

        logmesafe1 = tk.Label(inner_frame, text="Slwave file opened successfully", font=("Arial", 10), bg="white")
        logmesafe1.pack(pady=2, padx=2, anchor='nw')
    except:
        logmesafe1 = tk.Label(inner_frame, text="Error while opeing Slwave file", font=("Arial", 10), bg="white")
        logmesafe1.pack(pady=2, padx=2, anchor='nw')


    canvas.update_idletasks()
    canvas.configure(scrollregion=canvas.bbox("all"))

    return result


def on_select(value, selected_option):
    global version_selected, ansysVers, version_pathof_exe, checktrue

    if value != selected_option.get():
        version_selected = value
        print("ver", version_selected)
        version_pathof_exe = ansysVers[version_selected]
        print("path", version_pathof_exe)
        checktrue = False
    else:
        version_selected = value
        print("ver", version_selected)
        version_pathof_exe = ansysVers[version_selected]
        print("path", version_pathof_exe)
        checktrue = False


def upload():
    global Slwave_file_path
    try:
        logmesafe1 = tk.Label(inner_frame, text="Uploading Slwave file.....", font=("Arial", 10), bg="white")
        logmesafe1.pack(pady=3, padx=2, anchor='nw')
        file_path = filedialog.askopenfilename(title="Select slwave file",
                                               filetypes=(("Slwave files", "*.siw"), ("All files", "*.*")))
    except:
        logmesafe1 = tk.Label(inner_frame, text="Error while uploading Slwave file", font=("Arial", 10), bg="white")
        logmesafe1.pack(pady=3, padx=3, anchor='nw')

        canvas.update_idletasks()
        canvas.configure(scrollregion=canvas.bbox("all"))
        return

    if file_path:

        Slwave_file_path = file_path
        logmesafe1 = tk.Label(inner_frame, text="Upload successful", font=("Arial", 10), bg="white")
        logmesafe1.pack(pady=3, padx=2, anchor='nw')



    else:
        logmesafe1 = tk.Label(inner_frame, text="No Slwave file selected", font=("Arial", 10), bg="white")
        logmesafe1.pack(pady=2, padx=2, anchor='nw')

    canvas.update_idletasks()
    canvas.configure(scrollregion=canvas.bbox("all"))


if __name__ == "__main__":
    notewin()
    noteframe.pack()

    root.mainloop()
