from email.mime import text
import tkinter as tk
from tkinter import ttk

root = tk.Tk()
root.geometry("400x300")
style = ttk.Style(root)
root.tk.call('source', 'azure/azure.tcl')
style.theme_use('azure')

button = ttk.Button(root, text="hello", style="Accentbutton")
button.pack()

root.mainloop()