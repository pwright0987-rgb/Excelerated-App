###############################################
#
#
#
#     Excelerated written by Peter Wright
#
#   written for faster probe certification
#
#                   6/22/22
#
#
#
###############################################
import os
import tkinter as tk
from tkinter import *
from tkinter import ttk
import tkinter.font as tkFont
from tkinter.filedialog import askopenfiles
from tkinter.filedialog import askdirectory
import shutil
from openpyxl import load_workbook
from win32com import client

global files
global folder
global folder2
global count
count=0

global converting_in_progress
converting_in_progress = False

global progress
global books

'''
insert(str)
inserts the string formatted for progress box and resets disabled status
Param-> str: string to be added to progress box
Returns-> adds string to progress box formatted correctly
'''
def insert(str):
    progress.configure(state='normal')
    if (str[:5] == 'Error'):
        progress.insert(tk.END, "  " + str +'\n', 'red')
    else:
        progress.insert(tk.END, "  "+ str + "\n")
    progress.configure(state='disabled')
# end of insert(str)


'''
browse_click()
responds when browse button is clicked and assigns a file list and a folder for the updated files
Params-> none
Returns-> assigns global file and folder, adjusts count to show if successful
'''
def browse_click():
    if converting_in_progress:
        return
    global count
    count = 0
    global files
    files = askopenfiles(parent=root, mode='r+', title="Choose Excel files to use", 
        filetype=[("Excel file", "*.xlsx")]) # asks the user to pick Excel files
    if files: # if the files are chosen correctly
        count += 1
        if (len(files)==1):
            insert("File successfully loaded")
        else:
            insert("Files successfully loaded")
        global folder
        folder = askdirectory(title="Choose a directory to place new excel files in") # asks for folder to place files
        if folder:
            insert("Excel folder successfully selected")
            current_excel.set(folder)
            count += 1 # counts if file and folder are successful, if they are, count == 2)
            global folder2
            folder2 = askdirectory(title="Choose a directory to place new pdf certificates") # asks for folder to place pdf's in
            if folder2:
                insert("PDF folder successfully selected")
                current_pdf.set(folder2)
                count +=1
    if (count != 3):
        insert("Error: Browse unsuccessful. Try again.")
        current_excel.set('')
        current_pdf.set('')
# end of browse_click()


'''
convert_click()
starts the process of creating new Excel files and PDFs
'''
def convert_click():
    # CHECKS BEFORE STARTING
    global books
    global converting_in_progress
    if (converting_in_progress): # sees if button is currently working
        converting_in_progress = False
        convert_text.set('Convert')
        convert_btn.configure(bg=gray)
        root.update()
        return
    converting_in_progress = True
    convert_text.set('Stop')
    convert_btn.configure(bg='red')
    root.update()

    m = current_month.get()
    d = current_day.get()
    y = current_year.get()
    if (len(m)!=2 or len(d)!=2 or len(y)!=4): # makes sure date is entered properly
        insert('Error: Incorrect date (## / ## / ####)')
        converting_in_progress = False
        convert_text.set('Convert')
        convert_btn.configure(bg=gray)
        root.update()
        return
    global count
    if (count != 3): # makes sure browsing is complete
        insert('Error: Must browse first')
        converting_in_progress = False
        convert_text.set('Convert')
        convert_btn.configure(bg=gray)
        root.update()
        return

    # PROCESS BEGINS
    
    progbar.tkraise()
    times = 3*len(files)
    differential = (100+times) / times

    for i in range(len(files)): # goes through every excel file chosen
        original = files[i].name
        target = folder + '/' + os.path.basename(files[i].name)[:-5] + ' (Updated ' + m +'-' + d + '-' + y + ').xlsx' 
        if (not converting_in_progress): # checks to make sure no conversion is happening, otherwise stops this process
            converting_in_progress = False
            convert_text.set('Convert')
            convert_btn.configure(bg=gray)
            insert("Creation of new files interrupted")
            books.Close(True)
            progbar.lower()
            progbar['value'] = 0
            root.update()
            return  
        shutil.copyfile(original, target)                       # creates copies of the Excel files in the folder designated
        wb = load_workbook(filename = target, data_only=True) # opens copy
        sheet = wb['TPTTest']                                   # grabs data sheet
        sheet['N45'] = m + '/' + d + '/' + y                    # inserts date

        for sh in wb.worksheets[1:16]:                          # updates the cert dates
            sh['J7'] = '=DATE(' + y + ', ' + m + ', ' + d + ')'
            sh['J32'] = '=DATE(YEAR(TPTTest!N45)+1,MONTH(TPTTest!N45)+1,DAY(TPTTest!N45))'

        wb.save(filename = target)

        insert(os.path.basename(target) + " created")
        progbar['value'] += differential # increments the progress bar by one tick
        root.update()
        test1 = list(sheet[15][13:28]) # grabs all the test results for each test for each probe
        test2 = list(sheet[22][13:28])
        test3 = list(sheet[29][13:28])

        certs = [] # empty list of certs that are going to be printed
    
        for j in range(len(test1)):                             # creates list of passing probes
            test1[j] = test1[j].value
            test2[j] = test2[j].value
            test3[j] = test3[j].value
            if (test1[j]=='Pass' and test2[j]=='Pass' and test3[j]=='Pass'): # finds fully passing probes
                certs.append('Cert' + str((j+1))) # adds these to list
        #wb.save(filename=target)

        xlApp = client.Dispatch("Excel.Application")  
        xlApp.ScreenUpdating = False    # makes
        xlApp.DisplayAlerts = False     # processes
        xlApp.EnableEvents = False      # run
        xlApp.Visible = False           # in 
        xlApp.Interactive = False       # background
        books = xlApp.Workbooks.Open(target)    # opens the new Excel file
        progbar['value'] += differential    # increments the progress bar again
        root.update()
        if (not converting_in_progress): # this checks to see if stop has been pressed before continuing
            converting_in_progress = False
            convert_text.set('Convert')
            convert_btn.configure(bg=gray)
            progbar.lower()
            progbar['value'] = 0
            root.update()
            insert("Creation of new files interrupted")
            books.Close(True)
            return 
        books.Worksheets(certs).Select()    #selects only the passing sheets
        global folder2
        name = folder2 + '/' + os.path.basename(target)[:-5] + '.pdf'
        name = name.replace('/', "\\") 
        xlApp.ActiveSheet.ExportAsFixedFormat(0, name)  # creates a PDF with all the passing certs
        insert(os.path.basename(name) + ' created')
        insert('')
        books.Close(True)
        progbar['value'] += differential    #closes excel and increments bar again
        root.update()
        # end for loop going through files
    converting_in_progress = False
    convert_text.set('Convert')
    convert_btn.configure(bg=gray)
    progbar.lower()     # all excels have been finished, so the process is over
    progbar['value'] = 0
    root.update()
# end of convert_click()

#Colors used
green = '#016952'
white = 'white'
gray = '#444745'

#start of window GUI
root = tk.Tk(className = " Rees Scientific Excelerated")
root.resizable(False, False)

#Fonts
title = tkFont.Font(family="Times New Roman", size = 34, weight='bold')
big = tkFont.Font(family="Times New Roman", size = 15, weight='bold')
reg = tkFont.Font(family="Times New Roman", size = 12)

#Window Creation
canvas = tk.Canvas(root, width=940, height=500)
canvas.grid(columnspan=10, rowspan=10)
canvas.configure(background=green)

#quit button
quit = tk.Button(root, text="Quit", command=root.destroy, font=reg, width=8,
    bg=gray, fg=white)
quit.place(x=840, y=450)

#title 
title_label = tk.Label(text="Excelerated", background=green, 
    foreground=white, font=title)
title_label.place(x=375, y=30)

# integer used to move everything at once without changing a million numbers
p=100

#date text boxes
current_month = tk.StringVar()
month = ttk.Entry(root, textvariable=current_month, width = 3, font=reg)
month.place(x=p, y=125)
current_month.set('')
current_day = tk.StringVar()
day = ttk.Entry(root, textvariable=current_day, width=3, font=reg)
day.place(x=p+50, y=125)
current_day.set('')
current_year = tk.StringVar()
year = ttk.Entry(root, textvariable=current_year, width=5, font=reg)
year.place(x=p+100, y=125)
current_year.set('')

date_label = tk.Label(text = "Date:", background=green, 
    foreground=white, font=big)
date_label.place(x=p-55, y=125)
slash1 = tk.Label(text='/', bg=green, fg=white, font=big)
slash2 = tk.Label(text='/', bg=green, fg=white, font=big)
slash1.place(x=p+33, y=125)
slash2.place(x=p+83, y=(125))

#location text boxes
current_pdf = tk.StringVar()
pdf = ttk.Entry(root, textvariable=current_pdf, width=68, font=reg)
pdf.place(x=p+240, y=150)
pdf['state']='readonly'
current_pdf.set('')
pdf_label = tk.Label(text="PDF:", bg=green, fg=white, font=big)
pdf_label.place(x=p+185, y=149)

current_excel = tk.StringVar()
excel_dir = ttk.Entry(root, textvariable=current_excel, width=68, font=reg)
excel_dir.place(x=p+240, y=105)
excel_dir['state'] = 'readonly'
current_excel.set('')
excel_label = tk.Label(text='Excel:', bg=green, fg=white, font=big)
excel_label.place(x=p+175, y=104)


#browsing button
browse_text = tk.StringVar()
browse_btn = tk.Button(root, textvariable=browse_text, font=reg, width=8,
     fg=white, bg=gray, command=lambda:browse_click())
browse_text.set("Browse")
browse_btn.place(x=25, y= 450)

#convert button
convert_text= tk.StringVar()
convert_btn = tk.Button(root, textvariable=convert_text, font=reg, width=8,
    fg = white, bg = gray, command=lambda:convert_click())
convert_text.set('Convert')
convert_btn.place(x = 125, y=450)

#instructions
words = tk.StringVar()
words.set('Instructions for use:\n' + 
    '1) Click Browse button\n' + 
    '2) Select excel files to use\n' +
    '3) Select location for new excel files to be created\n' +
    '4) Select location for pdf certs to be created\n' +
    '5) Insert date, formated ## / ## / ####\n' +
    '6) Press convert\n' + 
    '7) Scroll in progress box to view progress\n' +
    '8) Wait until process finishes')
instructions = tk.Label(root, textvariable = words, font=reg, width = 35, height=9, justify=LEFT)
instructions.place(x=30, y=225)
instruct_lbl = tk.Label(text="Instructions", bg=green, fg=white, font=big)
instruct_lbl.place(x=138, y=190)

#progress textbox
progress = tk.Text(root, font=reg, width=68, height=9, state='disabled')
progress.place(x=365, y=225)
insert('Press Browse to Begin')
progress.tag_configure('red', foreground='red')
progress_lbl = tk.Label(text="Progress", bg=green, fg=white, font=big)
progress_lbl.place(x=600, y=190)

# style used to make the progress bar the right colors
s= ttk.Style()
s.theme_use('clam')
s.configure("green.Horizontal.TProgressbar", troughcolor = white, bordercolor = white,
    background = gray, lightcolor=gray, darkcolor=gray)

# progress bar used to show conversion progress
progbar = ttk.Progressbar(root, orient='horizontal', mode='determinate', 
    length=300, style = 'green.Horizontal.TProgressbar')
progbar.place(x=225, y=455)
progbar['value'] = 0
progbar.lower()

# needed for GUI
root.mainloop()
