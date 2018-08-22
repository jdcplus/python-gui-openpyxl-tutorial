from tkinter import *

class GUI(Frame):

    def __init__(self,master=None):
        Frame.__init__(self, master)
        self.grid()

        self.fnameLabel = Label(master, text="First Name")
        self.fnameLabel.grid()

        self.fnameEntry = Entry(master)
        self.fnameEntry.grid()

        self.lnameLabel = Label(master, text="Last Name")
        self.lnameLabel.grid()

        self.lnameEntry = Entry(master)
        self.lnameEntry.grid()

        self.submitButton = Button(master, command=self.buttonClick, text="Submit")
        self.submitButton.grid()


    def buttonClick(self):
        """ handle button click event and output text from entry area"""
        print('hello') 


if __name__ == "__main__":
    guiFrame = GUI()
    guiFrame.mainloop()