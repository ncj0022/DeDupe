import pandas as pd
import tkinter as tk
from tkinter import *
from tkinter import ttk
from tkinter import filedialog
from tkinter import messagebox as msg
from pandastable import Table
from tkintertable import TableCanvas


def convert_csv_to_xls():
    try:
        file_name = filedialog.askopenfilename(initialdir = '/Desktop',
                                                    title = 'Select a CSV file',
                                                    filetypes = (('csv file','*.csv'),
                                                                ('csv file','*.csv')))
        
        df = pd.read_csv(file_name)
        
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

def display_xls_file():
    try:
        file_name = filedialog.askopenfilename(initialdir = '/Desktop',
                                                    title = 'Select a excel file',
                                                    filetypes = (('excel file','*.xls'),
                                                                ('excel file','*.xls')))
        df = pd.read_excel(file_name)
        
        if (len(df)== 0):
            msg.showinfo('No records', 'No records')
        else:
            pass
            
        # Now display the DF in 'Table' object
        # under'pandastable' module
        f2 = Frame(root, height=200, width=300)
        f2.pack(fill=BOTH,expand=1)
        table = Table(f2, dataframe=df,read_only=True)
        table.show()

    except FileNotFoundError as e:
        print(e)
        msg.showerror('Error in opening file',e)

def remove_duplicate_account():
    try:
        file_name = filedialog.askopenfilename(initialdir = '/Desktop', title = 'Remove Duplicate Accounts', filetypes = (( 'excel file', '*.xls'), ('excel file', '*.xls')))
        rf = pd.read_excel(file_name)

        rf.drop_duplicates(subset=['Students Username'])

        if(len(rf)==0):
            msg.showinfo('No records', 'No records')
        else:
            pass

        f2 = Frame(root, height=200, width=300)
        f2.pack(fill=BOTH, expand=1)
        table = Table(f2, dataframe=rf,read_only=True)
        table.show()
    except FileNotFoundError as e:
        print(e)
        msg.showerror('Error in opening file', e)

# Driver Code
root = tk.Tk()
root.title('DeDupe')
root.geometry('500x500')

my_notebook = ttk.Notebook(root)
my_notebook.pack()

# Tabs
tab1 = Frame(my_notebook, width=500, height=500, bg="blue")
tab2 = Frame(my_notebook, width=500, height=500, bg="red")

tab1.pack(fill="both", expand=1)
tab2.pack(fill="both", expand=1)

my_notebook.add(tab1, text = "Dupe")
my_notebook.add(tab2, text = "CRM")

# Buttons
#help_email_button = Button(tab1, text = 'Help', fg ='blue', command=lambda: get_email())
convert_button = Button(tab1,
                        text = 'Convert',
                        font = ('Arial', 14),
                        bg = 'Orange',
                        fg = 'Black',
                        command = convert_csv_to_xls)
display_button = Button(tab1,
                        text = 'Display',
                        font = ('Arial', 14),
                        bg = 'Green',
                        fg = 'Black',
                        command = display_xls_file)
exit_button = Button(tab1,
                        text = 'Exit',
                        font = ('Arial', 14),
                        bg = 'Red',
                        fg = 'Black',
                        command = root.destroy)
remove_button = Button(tab1,
                        text = 'Remove',
                        font = ('Arial', 14),
                        bg = 'Blue',
                        fg = 'Black',
                        command = remove_duplicate_account)

# Place the widgets using grid manager on Dedupe tab
convert_button.grid(row = 3, column = 0,padx = 0, pady = 15)
display_button.grid(row = 3, column = 1,padx = 10, pady = 15)
exit_button.grid(row = 3, column = 2,padx = 10, pady = 15)
remove_button.grid(row = 4, column = 1,padx = 10, pady = 15)

root.mainloop()