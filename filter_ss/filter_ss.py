import pandas as pd
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
import tkinter.font as tkFont
from subprocess import Popen

class App:
    def __init__(self, root):
        #setting title
        root.title("MGSSHF")
        #setting window size
        width=490
        height=360
        screenwidth = root.winfo_screenwidth()
        screenheight = root.winfo_screenheight()
        alignstr = '%dx%d+%d+%d' % (width, height, (screenwidth - width) / 2,
                                    (screenheight - height) / 2)
        root.geometry(alignstr)
        root.resizable(width=False, height=False)
        self.filetypes = (('Any Spreadsheet', '.csv .xls .xslx .xslxm'),
                          ('Comma separated values', '*.csv'),
                          ('Old excel', '*.xls'),
                          ('Current excel','*.xlsx'))
        self.filetypes = () # This is more trouble than it is worth

        left = 20
        entryW = 350
        buttonW = 70
        H = 30

        titleLabel=tk.Label(root)
        titleLabel["fg"] = "#333333"
        titleLabel["font"] = tkFont.Font(family='Arial',size=16)
        titleLabel["justify"] = "center"
        titleLabel["text"] = "Spreadsheet Heading Filter"
        titleLabel.place(x=left,y=20)

        headerLabel=tk.Label(root)
        headerLabel["fg"] = "#333333"
        headerLabel["justify"] = "left"
        headerLabel["text"] = "1. Use the headers from this spreadsheet"
        headerLabel.place(x=left,y=70)

        self.headerTV = tk.StringVar()
        headerEntry=tk.Entry(root)
        headerEntry["justify"] = "left"
        headerEntry['textvariable'] = self.headerTV
        headerEntry.place(x=left,y=100,width=entryW,height=H)

        headerButton=tk.Button(root)
        headerButton["bg"] = "#f0f0f0"
        headerButton["fg"] = "#000000"
        headerButton["justify"] = "center"
        headerButton["text"] = "Select"
        headerButton.place(x=400,y=100,width=buttonW,height=H)
        headerButton["command"] = self.headerButton_command

        filterLabel=tk.Label(root)
        filterLabel["fg"] = "#333333"
        filterLabel["justify"] = "left"
        filterLabel["text"] = "2. To filter this spreadsheet"
        filterLabel.place(x=left,y=150)

        self.filterTV = tk.StringVar()
        filterEntry=tk.Entry(root)
        filterEntry["borderwidth"] = "1px"
        filterEntry["fg"] = "#333333"
        filterEntry["justify"] = "left"
        filterEntry['textvariable'] = self.filterTV
        filterEntry.place(x=left,y=180,width=entryW,height=H)

        filterButton=tk.Button(root)
        filterButton["bg"] = "#f0f0f0"
        filterButton["fg"] = "#000000"
        filterButton["justify"] = "center"
        filterButton["text"] = "Select"
        filterButton.place(x=400,y=180,width=buttonW,height=H)
        filterButton["command"] = self.filterButton_command

        saveLabel=tk.Label(root)
        saveLabel["fg"] = "#333333"
        saveLabel["justify"] = "left"
        saveLabel["text"] = "3. Save as new spreadsheet"
        saveLabel.place(x=left,y=230)

        self.saveTV = tk.StringVar()
        saveEntry=tk.Entry(root)
        saveEntry["borderwidth"] = "1px"
        saveEntry["fg"] = "#333333"
        saveEntry["justify"] = "left"
        saveEntry['textvariable'] = self.saveTV
        saveEntry.place(x=left,y=260,width=entryW,height=H)

        saveButton=tk.Button(root)
        saveButton["bg"] = "#f0f0f0"
        saveButton["fg"] = "#000000"
        saveButton["justify"] = "center"
        saveButton["text"] = "Save as"
        saveButton.place(x=400,y=260,width=66,height=H)
        saveButton["command"] = self.saveButton_command

        submitButton=tk.Button(root)
        submitButton["bg"] = "#f0f0f0"
        submitButton["fg"] = "#000000"
        submitButton["justify"] = "center"
        submitButton["text"] = "Filter"
        submitButton.place(x=left,y=310,width=450,height=H)
        submitButton["command"] = self.submitButton_command

        helpButton=tk.Button(root)
        helpButton["bg"] = "#f0f0f0"
        helpButton["fg"] = "#000000"
        helpButton["justify"] = "center"
        helpButton["text"] = "Help"
        helpButton.place(x=400,y=20,width=buttonW,height=H)
        helpButton["command"] = self.helpButton_command

    def headerButton_command(self):
        self.headerTV.set(self.choose_file('open',
                             title='Choose file to get the headers from',
                             filetypes=self.filetypes))

    def filterButton_command(self):
        self.filterTV.set(self.choose_file('open',
                             title='Choose file to filter',
                             filetypes=self.filetypes))

    def saveButton_command(self):
        self.saveTV.set(self.choose_file('save',
                           title='Save the new filtered file',
                           filetypes=self.filetypes))

    def submitButton_command(self):
        header_file = self.headerTV.get()
        work_file = self.filterTV.get()
        op_file = self.saveTV.get()
        headers = self.read_file(header_file).columns
        work = self.read_file(work_file)
        self.write_ss(work, op_file, headers)
        mgsep = '\\'
        command = f'explorer /select,"{op_file.replace("/",mgsep)}"'
        Popen(command)

    def helpButton_command(self):
        messagebox.showinfo(title=f'{root.title} Help',message = """
The purpose of this program is to automatically filter a large spreadsheet, so that it only contains the same headers as another spreadsheet.

Imagine you have a spreadsheet with the columns A, B, C, D, E, F, G, H, I, J, K, L, M, N, O, P, Q, R, S, T, U, V, W, X, Y, Z
But you only want columns S, I, C, E
You will have to manually delete the extra columns in excel the first time, but from now on, you can use your manually filtered spreadsheet to automatically filter new spreadsheets with this program.                    

The columns in the output file will also match the order of the headers file

The first field is where you put the previously manually curated spreadsheet (the headers file).
The second field is where you put the spreadsheet you want to filter
The third field is where you want this program to save the filtered spreadsheet.

Version 1.0
Made by Michael Gilmore 2023
        """)

    def choose_file(self, action, title=None, initialdir=None, filetypes=None):
        if filetypes is None:
            filetypes = self.filetypes
        if action == 'open':
            path = filedialog.askopenfilename(#initialdir="/", 
                                                 title=title,
                                                 filetypes=filetypes)
        elif action == 'save':
            path = filedialog.asksaveasfilename(#initialdir="/", 
                                                 title=title,
                                                 filetypes=filetypes)
        else:
            raise AssertionError(f'{action} not in ("open","save")')
        return path

    def read_file(self, file):
        ext = file.split('.')[-1].lower()
        if ext == 'csv':
            df = pd.read_csv(file)
        elif 'xl' in ext:
            df = pd.read_excel(file)
        else:
            messagebox.showerror('Unrecognised file type')
            raise BaseException('Unrecognised file type')
        return df

    def write_ss(self, df, opfile, columns):
        ext = opfile.split('.')[-1].lower()
        if ext == 'csv':
            df.to_csv(opfile, columns=columns, index=False)
        elif 'xl' in ext:
            df.to_excel(opfile, columns=columns, index=False)
        else:
            messagebox.showerror('Unrecognised file type')
            raise BaseException('Unrecognised file type')

if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()
