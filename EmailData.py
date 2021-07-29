import tkinter as tk
from tkinter import *

MImpressions = []
MClicks = []
CImpressions = []
CClicks = []
EMAILIMPRESSIONS = []
EMAILCLICKS = []

# Calender Tk window to choose dates for email blasts and unique emails
def main(DATES):
    global MImpressions, MClicks, CImpressions, CClicks
    

    Mailroot = Tk()
    # Set geometry
    Mailroot.geometry("500x500")

    ## Final data per date in array slot

    for i in DATES:
        MImpressions.append(tk.IntVar())    
        MClicks.append(tk.IntVar()) 
        CImpressions.append(tk.IntVar()) 
        CClicks.append(tk.IntVar()) 



    Label(Mailroot, text="MailChimp", font= ('Century 12 underline')).place(x=45,y=10)
    Label(Mailroot, text="Constant Contact", font= ('Century 12 underline')).place(x=335,y=10)

    Label(Mailroot, text="Impressions",).place(x=8,y=40)
    Label(Mailroot, text="Clicks",).place(x=120,y=40)

    Label(Mailroot, text="Impressions",).place(x=318,y=40)
    Label(Mailroot, text="Clicks",).place(x=440,y=40)
    
    y=40

    for i, val in enumerate(DATES):
        y = y + 40
        if("Unique" in val):
            Label(Mailroot, text="{}".format(val)).place(x=185, y=y)
        else:
            Label(Mailroot, text="{}".format(val)).place(x=220, y=y) 
        Entry(Mailroot, textvariable=MImpressions[i], width=10).place(x=10, y=y)
        Entry(Mailroot, textvariable=MClicks[i], width=10).place(x=105, y=y)
        Entry(Mailroot, textvariable=CImpressions[i], width=10).place(x=320, y=y)
        Entry(Mailroot, textvariable=CClicks[i], width=10).place(x=425, y=y)

    Button(Mailroot, text="Continue", command= lambda: calculate(DATES), width=15, height=1).place(x=355,y=450)


    Mailroot.mainloop()

def calculate(DATES):
    global MImpressions, MClicks, CImpressions, CClicks

    for i, val in enumerate(DATES):
        print(MImpressions[i].get() + CImpressions[i].get())
        EMAILIMPRESSIONS.append(MImpressions[i].get() + CImpressions[i].get())
        EMAILCLICKS.append(MClicks[i].get() + CClicks[i].get())
    print(EMAILIMPRESSIONS)
    print(EMAILCLICKS)
    return EMAILIMPRESSIONS, EMAILCLICKS



if __name__ == "__main__":
    DATES=['5/22/20', '5/23/20', '5/24/20']
    main(DATES)