from gettext import npgettext
import tkinter as tk
from tkinter import messagebox
from tkinter import filedialog
from TkinterDnD2 import *
import pandas as pd
import numpy as np

def main():
    root = TkinterDnD.Tk()
    root.drop_target_register(DND_FILES)
    root.dnd_bind('<<Drop>>', gen_stl)
    label = tk.Label(root)
    label["text"] = "Drag & Drop Your Excel file here"
    label["width"] = 30
    label["height"] = 10
    label.pack()
    root.mainloop()

def gen_stl(event):
    file = event.data
    if not file.endswith('xlsx'):
        messagebox.showwarning('Error', 'This is not an Excel file')
        return
    print('Reading ' + file)
    xls = pd.ExcelFile(file)   
    dfV = xls.parse('Sheet1')   # Vertices (x, y, z)
    dfI = xls.parse('Sheet2')   # Indices (triangles), 1-based
    vtx = dfV.to_numpy()
    idx = dfI.to_numpy()
    print('Vertices')
    print(vtx)
    print('Indices')
    print(idx)
    stl = make_stl_string(vtx, idx)
    out = filedialog.asksaveasfile(defaultextension='stl', title='Save as ...', filetypes=[('stl', '*.stl')])
    out.write(stl)
    out.close()
    messagebox.showinfo('Completed', 'STL file has been saved.\n' + out.name)

def make_stl_string(vtx: np.ndarray, idx: np.ndarray) -> str:
    stl = 'solid excelstl\n'
    nIdx = idx.shape[0]
    for i in range(nIdx):
        tri = idx[i]
        stl += 'facet normal 0 0 1\n'
        stl += 'outer loop\n'
        for j in range(3):
            # Convert indices from 1-based to 0-based
            stl += 'vertex ' + str(vtx[tri[j]-1][0]) + ' '
            stl += str(vtx[tri[j]-1][1]) + ' '
            stl += str(vtx[tri[j]-1][2]) + '\n' 
        stl += 'endloop\n'
        stl += 'endfacet\n'
    stl += 'endsolid excelstl\n'
    return stl


if __name__ == '__main__':
    main()