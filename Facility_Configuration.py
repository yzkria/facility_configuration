from tkinter import *
from tkinter import messagebox
from tkcalendar import Calendar
import datetime

root = Tk()
root.title("Login Screen")
root.iconbitmap('')
# The different message boxes you can use are:
# showinfo ,showwarning ,showerror ,askquestion ,askokcancel ,askyesno
def login_fun():
    user_name = username_entry.get()
    password = pass_entry.get()
    #check empty username or password
    if user_name == "":
        messagebox.showerror("Input Required","Please Enter Your Username")
    elif password =="":
        messagebox.showerror("Input Required","Please Enter Your Password")
    else:
        #check databse credentials are true
        # messagebox.showinfo("Login"," You login Successfully")
        # open main window for run reports
        main = Toplevel()
        main.title("Bulck Upload")
        main.iconbitmap('')
        d = datetime.date.today()
        reg_menu= StringVar()
        fac_menu= StringVar()
        Creteria_fram = LabelFrame(main,text="Creteria Selection",padx=5,pady=20)
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
        login_fram.destroy()

login_fram = LabelFrame(root,text="Login Data",padx=5,pady=5)
login_fram.pack(padx=10,pady=10,fill="both",expand="yes")

username_lbl = Label(login_fram,text="User Name:").grid(row=0,column=0,pady=5)
pass_lbl = Label(login_fram,text="Password:").grid(row=1,column=0,pady=5)

username_entry = Entry(login_fram)
username_entry.grid(row=0,column=1,pady=5)

pass_entry = Entry(login_fram,show="*")
pass_entry.grid(row=1,column=1,pady=5)

login_btn = Button(login_fram,text="Login",command=login_fun).grid(row=2,column=0,pady=5)


mainloop()
