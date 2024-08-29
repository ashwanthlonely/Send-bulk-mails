from src.gui import EmailSenderGUI
from tkinter import Tk

if __name__ == "__main__":
    root = Tk()
    app = EmailSenderGUI(root)
    root.mainloop()
