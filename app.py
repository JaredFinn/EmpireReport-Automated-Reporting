import tkinter as tk
from tkinter import *
from tkinter.ttk import Combobox

root = tk.Tk()
root.title("Report Automation")

canvas = tk.Canvas(root, height=600, width= 600, bg="#4a98f0")
canvas.pack()

frame = tk.Frame(root, bg="white")
frame.place(relwidth=0.9, relheight=0.4, relx=0.05, rely=0.05)

emailFrame = tk.Frame(root, bg="white")
emailFrame.place(relwidth=0.9, relheight=0.45, relx=0.05, rely=0.5)
emailLabel=Label(emailFrame, text="Drafted Email", bg="white")
emailLabel.place(x=10, y=5)

email = Text(emailFrame, bg="grey")
email.pack(padx=40, pady=30)

data=("Current Story", "Past Story", "Ad Report")
cb=Combobox(root, values=data)
cb.place(x=250, y=100)

reportLabel=Label(root, text="Select Report Type", bg="white")
reportLabel.place(x=135, y=100)
storyLabel= Label(root, text="Enter Story", bg="white")
storyLabel.place(x=135, y=150)

storyInput = Entry(root, width=30)
storyInput.place(x=210, y=150)

root.mainloop()