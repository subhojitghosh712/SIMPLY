import tkinter as tk
import threading as th
from tkinter import filedialog
import subprocess
import sys
import os

def spawn_program_and_die(program, exit_code=0):
    subprocess.Popen(program)
    sys.exit(exit_code)

def change():
    pg = filedialog.askopenfilename()
    os.system('"%s"' % pg)
    #spawn_program_and_die(['python', 'mainprogram.py'])

s = th.Timer(5.0, change)
s.start()

root = tk.Tk()
root.title("SIMPLY - Order Management System")
root.geometry("800x500")
root.resizable(0,0)
root.config(background="#003459")
icon = tk.PhotoImage(file='./Images/logo.png')
root.iconphoto(True, icon)

main_frame = tk.Frame(root, bg="#003459")
main_frame.propagate(False)
main_frame.pack()
main_frame.configure(width=1600,height=950)

sb = icon.subsample(3)
logo = tk.Label(main_frame, image=sb, bg="#003459")
logo.pack(pady=(30,20))

sbn = tk.PhotoImage(file='./Images/logoname.png').subsample(4)
logoname = tk.Label(main_frame, image=sbn, bg="#003459")
logoname.pack()

sbt1 = tk.PhotoImage(file='./Images/tg1.png').subsample(3)
tagline1 = tk.Label(main_frame, image=sbt1, bg="#003459")
tagline1.pack(padx=(30,0), pady=20)

sbt2 = tk.PhotoImage(file='./Images/tg2.png').subsample(3)
tagline2 = tk.Label(main_frame, image=sbt2, bg="#003459")
tagline2.pack(padx=(30,0), pady=40)

root.mainloop()

