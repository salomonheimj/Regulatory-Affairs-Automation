# Command for executable in pyinstaller: pyinstaller main_program.py --onefile --add-binary "chromedriver.exe;."
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
import pandas as pd


root= tk.Tk()
canvas1 = tk.Canvas(root, width = 300, height = 450, bg = 'lightsteelblue2', relief = 'raised')
canvas1.pack()

label1 = tk.Label(root, text='Seguimiento radicados', bg = 'lightsteelblue2')
label1.config(font=('helvetica', 20))
canvas1.create_window(150, 60, window=label1)

def getRegistros_nuevos ():
    root.quit()
    import registros_nuevos

browseButton_CSV = tk.Button(root, text="      Registros nuevos     ", command=getRegistros_nuevos, bg='green', fg='white', font=('helvetica', 12, 'bold'))
canvas1.create_window(150, 130, window=browseButton_CSV)


def getRenovaciones ():
    root.quit()
    import renovaciones

browseButton_CSV = tk.Button(root, text="      Renovaciones     ", command=getRenovaciones, bg='green', fg='white', font=('helvetica', 12, 'bold'))
canvas1.create_window(150, 180, window=browseButton_CSV)


def getModificaciones ():
    root.quit()
    import modificaciones

browseButton_CSV = tk.Button(root, text="      modificaciones     ", command=getModificaciones, bg='green', fg='white', font=('helvetica', 12, 'bold'))
canvas1.create_window(150, 230, window=browseButton_CSV)



def exitApplication():
    MsgBox = tk.messagebox.askquestion ('Cerrar aplicación','¿Seguro que deseas salir de la aplicación?',icon = 'warning')
    if MsgBox == 'yes':
       root.destroy()
       
exitButton = tk.Button (root, text='       Cerrar     ',command=exitApplication, bg='brown', fg='white', font=('helvetica', 12, 'bold'))
canvas1.create_window(150, 310, window=exitButton)

root.mainloop()