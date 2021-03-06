# danielcspears - open, process, and save 
from tkinter import *
from tkinter import ttk  
from tkinter.filedialog import asksaveasfilename
from tkinter.filedialog import askopenfilename
from tkinter.messagebox import showerror
import pandas as pd



class MyFrame(Frame):
    def __init__(self):
        Frame.__init__(self)
        self.master.title("Processing App")
        self.master.rowconfigure(14, weight=1)
        self.master.columnconfigure(14, weight=1)
        self.grid(sticky=W+E+N+S)
        self.grid_rowconfigure(1, minsize=5)
        self.grid_rowconfigure(4, minsize=10) 
        self.grid_rowconfigure(7, minsize=10)
        self.grid_rowconfigure(10, minsize=10)
        self.grid_rowconfigure(13, minsize=10)
        self.grid_columnconfigure(14, minsize=5)


        self.label = Label(self, text = "WIOA", width = 10)
        self.label.grid(row=3, column = 0, sticky = E)
        self.label1 = Label(self, text = "RESEA", width = 10)
        self.label1.grid(row=6, column = 0, sticky = E)
        self.label2 = Label(self, text = "Business", width = 10)
        self.label2.grid(row=9, column = 0, sticky = E)
        self.label3 = Label(self, text = "Wag Pey", width = 10)
        self.label3.grid(row=11, column = 0, sticky = E)

        self.buttonw = Button(self, text="Open File", command=self.load_wfile, width=10, fg="blue")
        self.buttonw.grid(row=3, column=3, sticky=W)
        self.buttonr = Button(self, text="Open File", command=self.load_rfile, width=10, fg="blue")
        self.buttonr.grid(row=6, column=3, sticky=W)
        self.buttonb = Button(self, text="Open File", command=self.load_bfile, width=10, fg="blue")
        self.buttonb.grid(row=9, column=3, sticky=W)
        self.buttonwpl = Button(self, text="Open File", command=self.load_wpfile, width=10, fg="blue")
        self.buttonwpl.grid(row=11, column=3, sticky=W)

        self.buttonwp = Button(self, text = "Process", command = self.calc_wioa, width = 10)
        self.buttonwp.grid(row=3, column=5, sticky=W)
        self.buttonrp = Button(self, text = "Process", command = self.calc_resea, width = 10)
        self.buttonrp.grid(row=6, column=5, sticky=W)
        self.buttonbp = Button(self, text = "Process", command = self.calc_business, width = 10)
        self.buttonbp.grid(row=9, column=5, sticky=W)
        self.buttonwpb = Button(self, text = "Process", command = self.calc_wp, width = 10)
        self.buttonwpb.grid(row=11, column=5, sticky=W)

        self.button4 = Button(self, text = "Save", command = self.save_wfile, width = 10, fg = 'red')
        self.button4.grid(row=3, column=7, sticky=W)
        self.button5 = Button(self, text = "Save", command = self.save_rfile, width = 10, fg = 'red')
        self.button5.grid(row=6, column=7, sticky=W)
        self.button6 = Button(self, text = "Save", command = self.save_bfile, width = 10, fg = 'red')
        self.button6.grid(row=9, column=7, sticky=W)
        self.button7 = Button(self, text = "Save", command = self.save_wpfile, width = 10, fg = 'red')
        self.button7.grid(row=11, column=7, sticky=W)

        # self.sep = ttk.Separator(self, orient=HORIZONTAL)
        # self.sep.grid(row = 4, column=2, columnspan=12, sticky=(E,W))
        self.sep1 = ttk.Separator(self, orient=VERTICAL)
        self.sep1.grid(row = 0, column=1, rowspan=12, sticky=(N,S))
        self.df = None
        self.data = None
        self.dfd = None
        self.df1c = None
        self.df1 = None
        self.df2 = None
        self.df2g = None

    def load_wfile(self):
        fname = askopenfilename(filetypes=(("Excel files", "*.xlsx"),("Excel files", "*.xls"),("All files", "*.*")))
        if fname:
            try:
                print("""here it comes: self.settings["template"].set(fname)""")
                # print(fname)
                self.df = pd.read_excel(fname, skiprows = 4)
            except:                     # <- naked except is a bad idea
                showerror("Open Source File", "Failed to read file\n'%s'" % fname)
            self.buttonw.config(state=DISABLED)
            return self.df


    def load_rfile(self):
        fname = askopenfilename(filetypes=(("Excel files", "*.xlsx"),("Excel files", "*.xls"),("All files", "*.*")))
        if fname:
            try:
                print("""here it comes: self.settings["template"].set(fname)""")
                # print(fname)
                self.data = pd.read_excel(fname, skiprows = 4)
            except:                     # <- naked except is a bad idea
                showerror("Open Source File", "Failed to read file\n'%s'" % fname)
            self.buttonr.config(state=DISABLED)
            return self.data


    def load_bfile(self):
        fname = askopenfilename(filetypes=(("Excel files", "*.xlsx"),("Excel files", "*.xls"),("All files", "*.*")))
        if fname:
            try:
                print("""here it comes: self.settings["template"].set(fname)""")
                # print(fname)
                self.df1 = pd.read_excel(fname, skiprows = 6)
            except:                     # <- naked except is a bad idea
                showerror("Open Source File", "Failed to read file\n'%s'" % fname)
            self.buttonb.config(state=DISABLED)
            return self.df1

    def load_wpfile(self):
        fname = askopenfilename(filetypes=(("Excel files", "*.xlsx"),("Excel files", "*.xls"),("All files", "*.*")))
        if fname:
            try:
                print("""here it comes: self.settings["template"].set(fname)""")
                # print(fname)
                self.df2 = pd.read_excel(fname, skiprows = 4)
            except:                     # <- naked except is a bad idea
                showerror("Open Source File", "Failed to read file\n'%s'" % fname)
            self.buttonwpl.config(state=DISABLED)
            return self.df2

    def calc_wioa(self):
        self.df = self.df[:-3]
        self.df = self.df.drop(self.df.columns[0], axis=1)
        self.df["Group"] = self.df["Service"].str[0] + "00s"
        self.df["Num of Activities"] = self.df['Service']
        self.df['Create Date'] = pd.to_datetime(self.df['Create Date']).dt.date
        self.df = self.df.drop_duplicates(['State ID','Create Date'])
        self.df = self.df[self.df['Completion Status']!= "* Void *"]
        self.df = self.df[~self.df['Service'].astype(str).str.startswith('F')]
        self.df = self.df[~self.df['Service'].astype(str).str.startswith('L')]
        self.df = self.df[["Staff Created", "Create Date", "Group", "Num of Activities",'Region / LWIA']]
        self.dfd = self.df.groupby(['Region / LWIA',"Staff Created","Create Date","Group"]).count()
        self.buttonwp.config(state=DISABLED)
        return self.dfd

    def calc_resea(self):    
        keep_col = ['Completion Status','State ID','Office','First Name','Last Name', 'Actual Date','Service', 'Staff Edited']
        self.data = self.data[keep_col]   
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
        self.data = self.data.loc[self.data['Minutes']!= 0].sort_values(by=["Staff Edited","Actual Date"])
        self.buttonrp.config(state=DISABLED)
        return self.data

    def calc_business(self):
        self.df1 =self.df1[:-3]
        self.df1 = self.df1.drop(self.df1.columns[0], axis=1)
        self.df1 = self.df1[['Emp. ID', 'Company Name', 'Service Code', 'Staff Reported', 'Actual\nDate']]
        self.df1 = self.df1.rename(index=str, columns={"Actual\nDate":"Actual Date"})
        self.df1['Actual Date'] = pd.to_datetime(self.df1['Actual Date']).dt.date
        self.df1 = self.df1.drop_duplicates(['Emp. ID','Actual Date'])
        #self.df1["Group"] = self.df1["Service Code"].str[:3]
        self.df1 = self.df1[self.df1['Staff Reported']!="System Set"]
        self.df1["Num of Activities"] = self.df1['Service Code']
        self.df1 = self.df1[['Staff Reported',"Actual Date",'Service Code','Num of Activities']]
        self.df1c = self.df1.groupby(['Staff Reported','Actual Date','Service Code',]).count()
        self.buttonbp.config(state=DISABLED)
        return self.df1c

    def calc_wp(self):
        disallowed = ['021 - Late Compliance of RESEA SP2', \
        'A00 - WIOP Attendance (ABAWD)', 'A20 - Adult/DW WIOP (ABAWD)', \
        '038 - Late Compliance of Initial RESEA',  'A10 - Youth WIOP (ABAWD)', \
        '037 - Continued UI Re-Employment Workshop/Orientation', 'A22 - Adult/DW WIOP (ABAWD) Re-referral', 
        'A12 - Youth WIOP (ABAWD) Re-referral', '022 - Late Compliance of RESEA SP3',\
        '121 - REA/RESEA Subsequent Call In (WP)','138 - Single Visit Completion of Initial RESEA', '650 - Enrolled in STEP Job Readiness Program']

        self.df2 = self.df2[:-3]
        self.df2 = self.df2.drop(self.df2.columns[0], axis=1)
        self.df2['Actual Date'] = pd.to_datetime(self.df2['Actual Date']).dt.date
        self.df2['Group'] = self.df2["Service"].str[0]
        self.df2 = self.df2.drop_duplicates(["State ID","Group", "Actual Date"])
        self.df2 = self.df2[~((self.df2[['State ID','Actual Date']].duplicated(keep=False)) & (self.df2['Service'].isin(disallowed).any()))]
        self.df2 = self.df2[self.df2["Service"]!= '133 - Self-Directed Job Search through VOS (WP)']
        self.df2 = self.df2[~self.df2['Service'].astype(str).str.startswith('A')]
        self.df2 = self.df2[~self.df2['Service'].astype(str).str.startswith('0')]
        self.df2 = self.df2[~self.df2['Service'].astype(str).str.startswith('L')]
        self.df2 = self.df2[~self.df2['Service'].astype(str).str.startswith('F')]
        self.df2 = self.df2[~self.df2['Service'].astype(str).str.startswith('E')]
        self.df2 = self.df2[~self.df2['Service'].astype(str).str.startswith('6')]
        self.df2 = self.df2[~self.df2['Service'].astype(str).str.startswith('5')]
        self.df2 = self.df2[~self.df2['Service'].astype(str).str.startswith('7')]
        self.df2 = self.df2[['UserName', 'State ID','Region / LWIA', 'Service', 'Group','Completion Status', 'Staff Created', 'Actual Date']]
        self.df2 = self.df2[self.df2['Completion Status'] == "Successful Completion"]
        self.df2 = self.df2[self.df2['Staff Created'] != "Process, GUS Batch "]
        self.df2["Num of Activities"] = self.df2['Service']
        self.df2 = self.df2[['Region / LWIA',"Staff Created", "Actual Date", "Service","Num of Activities"]]
        self.df2g = self.df2.groupby(['Region / LWIA','Staff Created','Actual Date','Service',]).count()
        self.buttonwpb.config(state=DISABLED)
        return self.df2g

    def save_wfile(self):
        file_wpath = asksaveasfilename(filetypes=(("Excel files", "*.xlsx"),("Excel files", "*.xls"),("All files", "*.*")))
        if file_wpath:
            self.button4.config(state=DISABLED)
            self.dfd.to_excel(file_wpath)

    def save_rfile(self):
        file_path = asksaveasfilename(filetypes=(("Excel files", "*.xlsx"),("Excel files", "*.xls"),("All files", "*.*")))
        if file_path:
            self.button5.config(state=DISABLED)
            self.data.to_excel(file_path)

    def save_bfile(self):
        file_path = asksaveasfilename(filetypes=(("Excel files", "*.xlsx"),("Excel files", "*.xls"),("All files", "*.*")))
        if file_path:
            self.button6.config(state=DISABLED)
            self.df1c.to_excel(file_path)

    def save_wpfile(self):
        file_path = asksaveasfilename(filetypes=(("Excel files", "*.xlsx"),("Excel files", "*.xls"),("All files", "*.*")))
        if file_path:
            self.button7.config(state=DISABLED)
            self.df2g.to_excel(file_path)


if __name__ == "__main__":
    MyFrame().mainloop()
    
