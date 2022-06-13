import tkinter as tk
from TkinterDnD2 import *

def gen_stl(event):
    print(event.data)

root = TkinterDnD.Tk()
root.drop_target_register(DND_FILES)
root.dnd_bind('<<Drop>>', gen_stl)
label = tk.Label(root)
label["text"] = "Drag & Drop Excel files here"
label["width"] = 30
label["height"] = 10
label.pack()
root.mainloop()
