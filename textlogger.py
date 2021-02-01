import logging
import tk
from tkinter import Text

class TextLogger():
    def __init__(self):
        self.textobj=None

    def log(self, record):
        logging.info(record)
        if self.textobj is None:
            return
        msg = record
        self.textobj.config(state='normal')
        self.textobj.insert('end', msg+'\n')
        self.textobj.yview('end')
        self.textobj.config(state='disabled')

    def set_target(self, target:Text):
        self.textobj = target
        self.textobj.config(state='disabled')
        
logging.basicConfig(filename='log.txt', level=logging.INFO)
textlogger = TextLogger()

