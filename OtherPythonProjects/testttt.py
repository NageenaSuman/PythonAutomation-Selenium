from tkinter import *
root = Tk()
root.geometry("300x100")
root.title("Input for DPI")

def printSomething():
    label.config(text=entry.get())


ok=Label(root, text="Enter the value for DPI:")
ok.grid(row=2,column=1)
entryvalue = StringVar()

entry= Entry(root, textvariable=entryvalue)
entry.grid(row=2, column=2)


button = Button(root, text="OK", command=printSomething)
button.grid(row=3, column=2)

#Create a Label to print the Name
label= Label(root, text="", font= ('Helvetica 14 bold'), foreground= "red3")
label.grid(row=1, column=2)

root.mainloop()
