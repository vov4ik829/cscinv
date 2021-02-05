from tkinter import *
from tkinter import filedialog
from tkinter import ttk
from textlogger import textlogger

from invoice import process


root = Tk()
root.title('Generate Invoices')
mainframe = ttk.Frame(root, padding="3 3 12 1")
mainframe.grid(row=0, column=0, sticky=(N, W, E, S))
root.columnconfigure(0, weight=1)
root.rowconfigure(0, weight=1)

s = ttk.Style()
label1 = ttk.Label(mainframe, text='Please pick a report file:')
label1.grid(row=1, column=1, sticky=W,padx=5,pady=5)
filename = StringVar()
filename_entry = ttk.Entry(mainframe, textvariable=filename, width=25)
filename_entry.state(['readonly'])
filename_entry.grid(row=1, column=2, sticky=W, padx=5,pady=5)
filename_button = ttk.Button(mainframe, text='Browse', command=lambda : filename.set(
        filedialog.askopenfilename(
            filetypes=[['Excel',['.xls','.xlsx']]]
            )))
filename_button.grid(row=1,column=3, sticky=W, padx=5,pady=5)
label2 = ttk.Label(mainframe, text='Please pick destination folder:')
label2.grid(row=2, column=1, sticky=W, padx=5,pady=5)
foldername = StringVar()
foldername_entry = ttk.Entry(mainframe, textvariable=foldername,width=25)
foldername_entry.grid(row=2, column=2, sticky=W, padx=5, pady=5)
foldername_entry.state(['readonly'])
foldername_button = ttk.Button(mainframe, text='Browse', command=lambda : foldername.set(
        filedialog.askdirectory()))
foldername_button.grid(row=2,column=3, sticky=W, padx=5, pady=5)

label3 = ttk.Label(mainframe, text='Please enter the date:')
label3.grid(row=3, column=1, sticky=W, padx=5, pady=5)
datestr = StringVar()
datestr_entry = ttk.Entry(mainframe, textvariable=datestr)
datestr_entry.grid(row=3, column=2, sticky=W, pady=5, padx=5)

pr_button = ttk.Button(mainframe, text="Process", command=lambda : process(
    filename.get(), foldername.get(),datestr.get()))
pr_button.grid(row=4, column=2, sticky=(E,W),padx=5,pady=5)

text = Text(mainframe, width=35, height=20)
text.grid(row=5, column=1, columnspan=3, sticky=(E,W),pady=5,padx=5)
textlogger.set_target(text, root)

root.mainloop()