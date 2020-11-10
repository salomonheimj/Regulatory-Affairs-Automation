# Command for executable in pyinstaller: pyinstaller Pantalla_principal.py --onefile --add-binary "chromedriver.exe;."
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
import pandas as pd


root= tk.Tk()
canvas1 = tk.Canvas(root, width = 300, height = 450, bg = 'lightsteelblue2', relief = 'raised')
canvas1.pack()

label1 = tk.Label(root, text='Procesos automaticos', bg = 'lightsteelblue2')
label1.config(font=('helvetica', 20))
canvas1.create_window(150, 60, window=label1)

def getRadicadosSeguimientos ():
    root.quit()
    import main_program

browseButton_CSV = tk.Button(root, text="Ir a seguimiento de radicados", command=getRadicadosSeguimientos, bg='green', fg='white', font=('helvetica', 12, 'bold'))
canvas1.create_window(150, 130, window=browseButton_CSV)


def getSometimientos ():
    root.quit()
    import sometimiento_medicamentos

browseButton_CSV = tk.Button(root, text="Ver sometimiento de medicamentos", command=getSometimientos, bg='green', fg='white', font=('helvetica', 12, 'bold'))
canvas1.create_window(150, 180, window=browseButton_CSV)



def exitApplication():
    MsgBox = tk.messagebox.askquestion ('Cerrar aplicación','¿Seguro que deseas salir de la aplicación?',icon = 'warning')
    if MsgBox == 'yes':
       root.destroy()
     
exitButton = tk.Button (root, text='       Cerrar     ',command=exitApplication, bg='brown', fg='white', font=('helvetica', 12, 'bold'))
canvas1.create_window(150, 250, window=exitButton)

root.mainloop()