import logging
import tk
from tkinter import Text, Tk
import tkinter

class TextLogger():
    def __init__(self):
        self.textobj=None
        self.root_to_refresh = None

    def log(self, record):
        logging.info(record)
        if self.textobj is None:
            return
        msg = record
        self.textobj.config(state='normal')
        self.textobj.insert('end', msg+'\n')
        self.textobj.yview('end')
        self.textobj.config(state='disabled')
        # EPXLAIN TO VAXO
        self.root_to_refresh.update_idletasks()


    def set_target(self, target:Text, root:Tk):
        self.textobj = target
        self.textobj.config(state='disabled')
        self.root_to_refresh = root

logger = logging.getLogger()
logger.addHandler(logging.FileHandler(filename='log.txt', encoding='utf-8'))

logging.basicConfig(level=logging.INFO, force=True)
textlogger = TextLogger()