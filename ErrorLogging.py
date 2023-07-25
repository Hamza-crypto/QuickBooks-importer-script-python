import tkinter as tk
import pandas as pd 
import os
from PathManager import locationManager as lm

def error_popup(msg):
    """Super simple pop-up to indicate an error has occured."""
    popup = tk.Tk()
    popup.wm_title("!")
    label = tk.Label(popup, text=msg)
    label.pack(side="top", fill="x", pady=10)
    B1 = tk.Button(popup, text="Okay", command = popup.destroy)
    B1.pack()
    popup.mainloop()


class textLog:
    def __init__(self):
        self.log = pd.DataFrame(columns=[''])
        self.lm = lm()
        self.current_loc = self.lm.get_current_loc()

    def append(self, description, location):
        """Adds row to log. Not commited until WRITE is run"""
        self.log = self.log.append(pd.DataFrame({'Description':description, 'Location':location}))

    def WRITE(self):
        self.log.to_csv(os.path.join(self.current_location, 'REFERENCE_ERROR.csv'), index=False)
