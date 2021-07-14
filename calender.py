import tkinter
from tkinter import *
from tkinter import *
from tkcalendar import Calendar



def main(UNIQUEDATES, DATES):

    calRoot = Tk()

    # Set geometry
    calRoot.geometry("400x600")

    dateDir = tkinter.Label(calRoot, text="Choose and add the dates the program was sponsored in the email")
    dateDir.pack(pady=10)

    # Add Calendar
    cal = Calendar(calRoot, selectmode = 'day',
                year = 2020, month = 5,
                day = 22)

    cal.pack(pady = 20)

    x=15
    y=270


    # Add Button and Label
    Button(calRoot, text = "Add Date",
        command =lambda: add_date(x,y,False)).place(x=90, y=260)
    # Add Button and Label
    Button(calRoot, text = "Add Unqiue Email Date",
        command =lambda: add_date(x,y,True)).place(x=175, y=260)


    date = Label(calRoot, text = "")
    date.pack(pady = 20)

    def add_date(x,y,unique):
        if(unique == True):
            UNIQUEDATES.append(cal.get_date())
            DATES.append(cal.get_date() + " Unique Email")
        else:
            DATES.append(cal.get_date())
        for i in DATES:
            y = y + 20
            if(y == 550):
                y = 290
                x = x + 80
            Label(calRoot, text="{}".format(i)).place(x=x, y=y)
        print(DATES)

    calRoot.mainloop()