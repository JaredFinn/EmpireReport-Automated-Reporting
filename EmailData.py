import tkinter as tk
from tkinter import *




# Calender Tk window to choose dates for email blasts and unique emails
def main(DATES):

    root = Tk()
    # Set geometry
    root.geometry("500x500")

    MImpressions= []
    MClicks = []
    CImpressions = []
    CClicks = []
    ## Final data per date in array slot


    Label(root, text="MailChimp", font= ('Century 12 underline')).place(x=45,y=10)
    Label(root, text="Constant Contact", font= ('Century 12 underline')).place(x=335,y=10)

    Label(root, text="Impressions",).place(x=8,y=40)
    Label(root, text="Clicks",).place(x=120,y=40)

    Label(root, text="Impressions",).place(x=318,y=40)
    Label(root, text="Clicks",).place(x=440,y=40)
    
    y=40

    for i, val in enumerate(DATES):
        y = y + 40
        if("Unique" in val):
            Label(root, text="{}".format(val)).place(x=185, y=y)
        else:
            Label(root, text="{}".format(val)).place(x=220, y=y)
        MImpressions.append(tk.IntVar())    
        Entry(root, textvariable=MImpressions[i], width=10).place(x=10, y=y)
        MClicks.append(tk.IntVar()) 
        Entry(root, textvariable=MClicks[i], width=10).place(x=105, y=y)
        CImpressions.append(tk.IntVar()) 
        Entry(root, textvariable=CImpressions[i], width=10).place(x=320, y=y)
        CClicks.append(tk.IntVar()) 
        Entry(root, textvariable=CClicks[i], width=10).place(x=425, y=y)


    def calculate():
        for i, val in enumerate(DATES):
            EMAILIMPRESSIONS.append(MImpressions[i].get() + CImpressions[i].get())
            EMAILCLICKS.append(MClicks[i].get() + CClicks[i].get())
        print(EMAILIMPRESSIONS[i].get())
        print(EMAILCLICKS[i].get())
        root.destroy
        return EMAILIMPRESSIONS, EMAILCLICKS


    Button(root, text="Continue", command= lambda: calculate(), width=15, height=1).place(x=355,y=450)


    root.mainloop()




if __name__ == "__main__":
    DATES=['5/22/20', '5/23/20', '5/24/20']
    main(DATES)