from gettext import npgettext
import tkinter as tk
from tkinter import messagebox
from tkinter import filedialog
from TkinterDnD2 import *
import pandas as pd
import numpy as np

dataUnit = ''
thickness = 5 # [mm]

class Application(tk.Frame):
    def __init__(self, master):
        super().__init__(master)
        self.pack()
        self.master.title('ExcelSTL')
        self.drop_target_register(DND_FILES)
        self.dnd_bind('<<Drop>>', gen_stl)
        # Unit radio
        f1 = tk.LabelFrame(self, text='Unit')
        f1.pack()
        var = tk.StringVar()
        def setDataUnit():
            global dataUnit
            dataUnit = var.get()
        cm = tk.Radiobutton(f1, value='cm', variable=var, text='[cm] to [mm]', command=setDataUnit)
        cm.pack(anchor=tk.W)
        mm = tk.Radiobutton(f1, value='mm', variable=var, text='[mm] to [mm] (no change)', command=setDataUnit)
        mm.pack(anchor=tk.W)
        cm.invoke()
        # D&D label
        label = tk.Label(self)
        label["text"] = "Drag & Drop Your Excel file here"
        label["width"] = 30
        label["height"] = 10
        label.pack()

def main():
    root = TkinterDnD.Tk()
    app = Application(master=root)
    app.mainloop()

def gen_stl(event):
    file = event.data
    if not file.endswith('xlsx'):
        messagebox.showerror('Error', 'This is not an Excel file.')
        return
    print('Reading ' + file)
    xls = pd.ExcelFile(file)   
    dfV = xls.parse('Sheet1', header=None)   # Vertices (x, y, z)
    dfI = xls.parse('Sheet2', header=None)   # Indices (triangles), 1-based
    vtx = dfV.to_numpy()
    idx = dfI.to_numpy()
    print('Vertices')
    print(vtx)
    print('Indices')
    print(idx)
    # 2D mode
    if vtx.shape[1] == 2:
        stl = make_stl_string_2d(vtx, idx)
    # 3D mode
    elif vtx.shape[1] == 3:
        stl = make_stl_string(vtx, idx)
    else:
        messagebox.showerror('Error', 'Size of array is wrong.')
        return
    out = filedialog.asksaveasfile(defaultextension='stl', title='Save as ...', filetypes=[('stl', '*.stl')])
    out.write(stl)
    out.close()


def make_stl_string_2d(vtx: np.ndarray, idx: np.ndarray) -> str:
    if dataUnit == 'cm':
        vtx *= 10.0
    nVtx0 = vtx.shape[0]
    nIdx0 = idx.shape[0]
    # Append Z-axis values (0)
    vtx = np.append(vtx, np.zeros(nVtx0).reshape(-1, 1), axis=1) 
    # Append bottom vertices
    vtx3D = np.append(vtx, vtx, axis=0) 
    for i in range(nVtx0):
        vtx3D[i][2] += thickness
    # Append bottom triangles
    idx3D = np.append(idx, idx, axis=0)
    for i in range(nIdx0):
        idx3D[i+nIdx0][0] += nVtx0
        idx3D[i+nIdx0][1] += nVtx0
        idx3D[i+nIdx0][2] += nVtx0
    # Append side triangles
    for i in range(nIdx0):
        for j in range(3):
            e11 = idx3D[i][j]
            e12 = idx3D[i][(j+1)%3]
            e21 = idx3D[i+nIdx0][j]
            e22 = idx3D[i+nIdx0][(j+1)%3]
            if is_edge_unique(idx, e11, e12):
                idx3D = np.append(idx3D, np.array([e11, e21, e12]).reshape(1, -1), axis=0)
                idx3D = np.append(idx3D, np.array([e12, e21, e22]).reshape(1, -1), axis=0)
    # Invert bottom triangles
    for i in range(nIdx0):
        idx3D[i+nIdx0][1], idx3D[i+nIdx0][2] = idx3D[i+nIdx0][2], idx3D[i+nIdx0][1] # swap
    return make_stl_string(vtx3D, idx3D)

def is_edge_unique(idx: np.ndarray, v1: int, v2: int):
    n = idx.shape[0]
    duplicateCount = 0
    # Search all edges
    for i in range(n):
        for j in range(3):
            duplicateCount += (v1 == idx[i][j] and v2 == idx[i][(j+1)%3])
            duplicateCount += (v2 == idx[i][j] and v1 == idx[i][(j+1)%3])
    return duplicateCount == 1

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