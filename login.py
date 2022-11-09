from tkinter import Tk, Toplevel, Button
from tkinter import *
from tkinter import messagebox
from tkcalendar import Calendar
import datetime


class MainWindow(Tk):
    def __init__(self):
        Tk.__init__(self)
        self.title('Bulck Upload')
        self.geometry('500x300')
        #check databse credentials are true
        # messagebox.showinfo("Login"," You login Successfully")
        # open main window for run reports
        self.iconbitmap('')
        d = datetime.date.today()
        reg_menu= StringVar()
        fac_menu= StringVar()
        Creteria_fram = LabelFrame(self,text="Creteria Selection",padx=5,pady=20)
        Creteria_fram.pack(padx=10,pady=10,fill="both",expand="yes")
        region_lbl = Label(Creteria_fram,text="Select Region:").grid(row=0,column=0,pady=20)
        region_menu = OptionMenu(Creteria_fram, reg_menu, "Aswan", "Luxor", "Sharm", "PortSaid","Ismailia","SouthS ini").grid(row=0,column=1,pady=20)
        facility_lbl = Label(Creteria_fram,text="Select Facility:").grid(row=2,column=0,pady=20)
        facility_menu = OptionMenu(Creteria_fram, fac_menu, "Facility1", "Facility2", "Facility3", "Facility4","Facility5",).grid(row=2,column=1,pady=20)
        from_date_lbl = Label(Creteria_fram,text="From Date:").grid(row=3,column=0,pady=20)
        from_date = Calendar(Creteria_fram,text="From Date:",selectmode="day",year=d.year,month=d.month,day=d.day).grid(row=3,column=1,pady=20)
        #from_date_lbl.get_date()
        to_date_lbl = Label(Creteria_fram,text="To Date:").grid(row=4,column=0,pady=20)
        to_date = Calendar(Creteria_fram,text="To Date:",selectmode="day",year=d.year,month=d.month,day=d.day).grid(row=4,column=1,pady=20)
        #to_date_lbl.get_date()


class LoginWindow(Toplevel):
    def __init__(self, parent):
        Toplevel.__init__(self, parent)
        self.parent = parent
        self.title('Login Window')
        self.geometry('500x300')
        self.iconbitmap('')
        self.protocol('WM_DELETE_WINDOW', root.quit)
        login_fram = LabelFrame(self,text="Login Data",padx=5,pady=5)
        login_fram.pack(padx=10,pady=10,fill="both",expand="yes")

        username_lbl = Label(login_fram,text="User Name:").grid(row=0,column=0,pady=5)
        pass_lbl = Label(login_fram,text="Password:").grid(row=1,column=0,pady=5)

        global username_entry
        global pass_entry

        username_entry = Entry(login_fram)
        username_entry.grid(row=0,column=1,pady=5)

        pass_entry = Entry(login_fram,show="*")
        pass_entry.grid(row=1,column=1,pady=5)

        login_btn = Button(login_fram,text="Login",command=self.login).grid(row=2,column=0,pady=5)

    def login(self):
        user_name = username_entry.get()
        password = pass_entry.get()
        #check empty username or password
        if user_name == "":
            messagebox.showerror("Input Required","Please Enter Your Username")
        elif password =="":
            messagebox.showerror("Input Required","Please Enter Your Password")
        else:
            self.destroy()
            self.parent.deiconify()
            #check databse credentials are true
            # messagebox.showinfo("Login"," You login Successfully")
            # open main window for run reports



if __name__ == '__main__':
    root = MainWindow()
    root.withdraw()

    login = LoginWindow(root)

    root.mainloop()
