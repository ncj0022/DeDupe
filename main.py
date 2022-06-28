import pandas as pd
import tkinter as tk
from tkinter import *
from tkinter import ttk
from tkinter import filedialog
from tkinter import messagebox as msg
from pandastable import Table
from tkintertable import TableCanvas

# Class for DeDupe application
class  DeDupe:
    def __init__(self, root):
        self.root = root
        self.file_name = ' '
        self.f = Frame(self.root, height=200, width=300, bg = "#00853E")

        self.f.pack()

        # Buttons
        #help_email_button = Button(tab1, text = 'Help', fg ='blue', command=lambda: get_email())
        self.convert_button = Button(self.f,
                                text = 'Convert',
                                font = ('Arial', 14),
                                bg = 'Orange',
                                fg = 'Black',
                                command = self.convert_csv_to_xls)
        self.display_button = Button(self.f,
                                text = 'Display',
                                font = ('Arial', 14),
                                bg = 'Green',
                                fg = 'Black',
                                command = self.display_xls_file)
        self.exit_button = Button(self.f,
                                text = 'Exit',
                                font = ('Arial', 14),
                                bg = 'Red',
                                fg = 'Black',
                                command = root.destroy)
        self.remove_button = Button(self.f,
                                text = 'Remove',
                                font = ('Arial', 14),
                                bg = 'Blue',
                                fg = 'Black',
                                command = self.remove_duplicate_account)

        #Create Entry Widget for user input.
        self.e1 = tk.Entry(self.f)
        self.e2 = tk.Entry(self.f)
        self.e3 = tk.Entry(self.f)

        # Place the widgets using grid manager on Dedupe Page
        self.convert_button.grid(row = 3, column = 0,padx = 0, pady = 15)
        self.display_button.grid(row = 4, column = 0,padx = 10, pady = 15)
        self.remove_button.grid(row = 5, column = 0,padx = 10, pady = 15)

        # Add Input fields to grid
        self.e1.grid(row=5, column= 1)
        self.e2.grid(row=5, column=2)
        self.e3.grid(row=5, column=3)

        self.exit_button.grid(row = 8, column = 0,padx = 10, pady = 15)


    def choose_fields(self):
        print("Field 1: %s\nField 2: %s\n Field 3: %s" % (self.e1.get(), self.e2.get(), self.e3.get()))

    # Convert csv file to excel file
    def convert_csv_to_xls(self):
        try:
            self.file_name = filedialog.askopenfilename(initialdir = '/Desktop',
                                                        title = 'Select a CSV file',
                                                        filetypes = (('csv file','*.csv'),
                                                                    ('csv file','*.csv')))
            
            df = pd.read_csv(self.file_name)
            
            # Next - Pandas DF to Excel file on disk
            if(len(df) == 0):	
                msg.showinfo('No Rows Selected', 'CSV has no rows')
            else:
                
                # saves in the current directory
                with pd.ExcelWriter('DeDupe.xls') as writer:
                        df.to_excel(writer,'GFGSheet')
                        writer.save()
                        msg.showinfo('Excel file created', 'Excel File created')	
        
        except FileNotFoundError as e:
                msg.showerror('Error in opening file', e)

    # Display chosen csv or excel file in Frame
    def display_xls_file(self):
        try:
            self.file_name = filedialog.askopenfilename(initialdir = '/Desktop',
                                                        title = 'Select a excel file',
                                                        filetypes = (('excel file','*.xls'),
                                                                    ('excel file','*.xls')))
            df = pd.read_excel(self.file_name)
            
            if (len(df)== 0):
                msg.showinfo('No records', 'No records')
            else:
                pass
                
            # Now display the DF in 'Table' object
            # under'pandastable' module
            self.f2 = Frame(self.root, height=200, width=300)
            self.f2.pack(fill=BOTH,expand=1)
            self.table = Table(self.f2, dataframe=df,read_only=True)
            self.table.show()

        except FileNotFoundError as e:
            print(e)
            msg.showerror('Error in opening file',e)

    # Remove duplicate records in chosen excel file
    # Base on field name given by the user
    def remove_duplicate_account(self):
        try:
            self.file_name = filedialog.askopenfilename(initialdir = '/Desktop', title = 'Remove Duplicate Accounts', filetypes = (( 'excel file', '*.xls'), ('excel file', '*.xls')))
            rf = pd.read_excel(self.file_name)

            # Remove duplicates from field based on first field
            rf.drop_duplicates(subset=[self.e1.get()])

            if(len(rf)==0):
                msg.showinfo('No records', 'No records')
            else:
                pass

            # Display table object
            self.f2 = Frame(root, height=200, width=300)
            self.f2.pack(fill=BOTH, expand=1)
            self.table = Table(self.f2, dataframe=rf,read_only=True)
            self.table.show()
        except FileNotFoundError as e:
            print(e)
            msg.showerror('Error in opening file', e)

# Driver Code
root = Tk()
root.title('DeDupe')

# Create DeDupe object
obj = DeDupe(root)
root.geometry('800x600')

# Run main
root.mainloop()