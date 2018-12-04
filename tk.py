# from tkinter import *
# from tkinter import ttk
# from tkinter import filedialog

# interface = Tk()

# def openfile():
#     return filedialog.askopenfilename()

# button = ttk.Button(interface, text="Open", command=openfile)  # <------
# button.grid(column=1, row=1)

# interface.mainloop()


from tkinter import *
from tkinter.filedialog import askopenfilename
from tkinter.messagebox import showerror
import pandas as pd

class MyFrame(Frame):
    def __init__(self):
        Frame.__init__(self)
        self.master.title("Example")
        self.master.rowconfigure(5, weight=1)
        self.master.columnconfigure(5, weight=1)
        self.grid(sticky=W+E+N+S)

        self.button = Button(self, text="Browse", command=self.load_file, width=10)
        self.button.grid(row=1, column=0, sticky=W)
        self.button2 = Button(self, text = "Process", command = self.calc_wioa, width = 10)
        self.button2.grid(row=1, column=1, sticky=W)
        self.df = None
        


    def load_file(self):
        fname = askopenfilename(filetypes=(("Excel files", "*.xlsx"),
                                           ("Excel files", "*.xls"),
                                           ("All files", "*.*") ))
        if fname:
            try:
                print("""here it comes: self.settings["template"].set(fname)""")
                # print(fname)
                self.df = pd.read_excel(fname, skiprows = 4)

            except:                     # <- naked except is a bad idea
                showerror("Open Source File", "Failed to read file\n'%s'" % fname)
            return self.df

    def calc_wioa(self):
                print(self.df)
                # df = pd.read_excel(fname, skiprows = 4)
                self.df = self.df[:-3]
                self.df = self.df.drop(self.df.columns[0], axis=1)
                #dfg = df.groupby("Office")["Service"].count()
                self.df["Group"] = self.df["Service"].str[0] + "00s"
                self.df["Num of Activities"] = self.df['Service']
                self.df['Create Date'] = pd.to_datetime(self.df['Create Date']).dt.date
                dfd = self.df[["Staff Created", "Create Date", "Group", "Num of Activities",'Region / LWIA']]
                dfd = dfd.groupby(['Region / LWIA',"Staff Created","Create Date","Group"]).count()
                # #df = df[['UserName', 'State ID', 'Region / LWIA', 'Office', ' Office of Responsibility', 'First Name', 'Last Name', 'City, State, Country', 'Service', 'Completion Status', 'Provider', 'Program', 'Staff Created', 'Create Date', 'Actual Begin Date', 'Projected Begin Date', 'Actual End Date', 'Projected End Date', 'Staff Edited']]
                # #df = df.drop_duplicates(['State ID','Service','Create Date'])
                # ## =============================================================================
                # ## drop duplicates if on same date.
                # ## =============================================================================
                # df = df.drop_duplicates(['State ID','Create Date'])

                # df = df[['UserName', 'State ID', 'Region / LWIA', 'Office', ' Office of Responsibility', 'Service', 'Completion Status', 'Staff Created', 'Create Date', 'Staff Edited']]
                # df = df[df['Completion Status']!= "* Void *"]
                # df = df[~df['Service'].astype(str).str.startswith('F')]

                dfc = self.df.groupby('Service').count()
                print(dfd)
                dfd.to_excel('/Users/Spears/Downloads/wioalist.xlsx')
        

if __name__ == "__main__":
    MyFrame().mainloop()
    
