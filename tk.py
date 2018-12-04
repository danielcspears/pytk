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
from tkinter import ttk  
from tkinter.filedialog import asksaveasfilename
from tkinter.filedialog import askopenfilename
from tkinter.messagebox import showerror
import pandas as pd



class MyFrame(Frame):
    def __init__(self):
        Frame.__init__(self)
        self.master.title("Process Files")
        self.master.rowconfigure(7, weight=1)
        self.master.columnconfigure(7, weight=1)
        self.grid(sticky=W+E+N+S)

        self.button = Button(self, text="Open File", command=self.load_wfile, width=10)
        self.button.grid(row=0, column=2, sticky=W)
        self.label = Label(self, text = "WIOA", width = 10)
        self.label.grid(row=3, column = 0, sticky = W)
        self.label1 = Label(self, text = "RESEA", width = 10)
        self.label1.grid(row=5, column = 0, sticky = W)
        self.button2 = Button(self, text = "Process", command = self.calc_wioa, width = 10)
        self.button2.grid(row=3, column=2, sticky=W)
        self.button3 = Button(self, text = "Process", command = self.calc_resea, width = 10)
        self.button3.grid(row=5, column=2, sticky=W)
        self.button4 = Button(self, text = "Save", command = self.save_wfile, width = 10, background = 'red')
        self.button4.grid(row=7, column=3, sticky=W)
        # self.button5 = Button(self, text = "Save", command = self.save_rfile, width = 10, bg = 'red')
        # self.button5.grid(row=4, column=3, sticky=W)
        self.sep = ttk.Separator(self, orient=HORIZONTAL)
        self.sep.grid(row = 2, column=0, columnspan=7, sticky=(E,W))
        # self.sep1 = ttk.Separator(self, orient=HORIZONTAL)
        # self.sep1.grid(row = 4, column=1, columnspan=1, sticky=(E,W))
        self.df = None
        self.data = None
        self.dfd = None


    def load_wfile(self):
        fname = askopenfilename(filetypes=(("Excel files", "*.xlsx"),
                                           ("Excel files", "*.xls"),
                                           ("All files", "*.*")))
        if fname:
            try:
                print("""here it comes: self.settings["template"].set(fname)""")
                # print(fname)
                self.df = pd.read_excel(fname, skiprows = 4)
            

            except:                     # <- naked except is a bad idea
                showerror("Open Source File", "Failed to read file\n'%s'" % fname)
            return self.df
    def load_rfile(self):
        fname = askopenfilename(filetypes=(("Excel files", "*.xlsx"),
                                           ("Excel files", "*.xls"),
                                           ("All files", "*.*")))
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
        self.dfd.to_excel('/Users/Spears/Downloads/wioalist.xlsx')
        return self.dfd

    def calc_resea(self):    
        keep_col = ['Completion Status','State ID','Office','First Name','Last Name', 'Actual Date','Service', 'Staff Edited']
        self.data = self.df[keep_col]   
        self.data = self.data.loc[self.data['Completion Status'] == 'Successful Completion']
        self.data['Actual Date'] = pd.to_datetime(self.data['Actual Date']).dt.date
        self.data = self.data.drop_duplicates(['State ID','Actual Date'])
        self.data['Minutes'] = 0.0
        self.data['Minutes'].loc[(self.data['Service']=='138 - Single Visit Completion of Initial RESEA') | \
        (self.data['Service']=='038 - Late Compliance of Initial RESEA')| \
        (self.data['Service']=='037 - Continued UI Re-Employment Workshop/Orientation')] = 90.0     
        self.data['Minutes'].loc[(self.data['Service']=='021 - Late Compliance of RESEA SP2') | \
        (self.data['Service']=='121 - REA/RESEA Subsequent Call In (WP)')] = 65.0
        self.data = self.data.groupby(['Staff Edited' , 'Office', 'Actual Date'])['Minutes'].sum().reset_index()

        # =============================================================================
        # round
        # =============================================================================
        def roundx(x):
            return round(x*4)/4.0

        self.data['Hours to Charge'] = (self.data['Minutes']/60).apply(roundx)
        # self.data.to_excel('/Users/Spears/Downloads/reseatime.xls')
        return self.data

    # def save_rfile(self):
    #     file_path = asksaveasfilename(filetypes=(("Excel files", "*.xlsx"),("Excel files", "*.xls"),("All files", "*.*")))
    #     if file_path:
    #         self.data.to_excel(file_path)

    def save_wfile(self):
        file_path = asksaveasfilename(filetypes=(("Excel files", "*.xlsx"),("Excel files", "*.xls"),("All files", "*.*")))
        if file_path:
            self.data.to_excel(file_path)


if __name__ == "__main__":
    MyFrame().mainloop()
    
