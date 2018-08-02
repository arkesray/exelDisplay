from tkinter import *
import xlrd
import datetime


"""Defining the main class for the display"""
class Application(Frame):

    def __init__(self, master):
        Frame.__init__(self, master)
        self.grid()
        self.create_widgets()

    def create_widgets(self):
        for i in range(0,len(data)):
            count = 0
            for j in range(0,len(data[0])):
                if i == 0:
                    if len(str(data[i][j])) > 20:
                        label = Label(self, font = "none 12 italic underline",text= str(data[i][j])[0:16]+"...")
                    else:
                        label = Label(self, font = "none 12 italic underline",text= str(data[i][j]))
                else:
                    if len(str(data[i][j])) > 20:
                        label = Label(self,text= str(data[i][j])[0:16]+"...")
                    else:
                        label = Label(self,text= str(data[i][j]))
                label.grid(row = i + 30,column = colStarts[count],columnspan = 10,sticky = W)
                count = count + 1

def getName():
    global file_location 
    file_location = entry.get()
    global root1
    root1.destroy()

file_location = ""
entry = ""

root1 = Tk()
root1.title("Selection Page")
root1.geometry("300x150")
frame = Frame(root1, bg = "black")
label = Label(frame, text = "Enter the path to Excel File", bg = "black", fg = "white")
entry = Entry(frame, width = 30)
button = Button(frame,text = "Submit", command = getName, bg = "black", fg = "white")
label.pack()
entry.pack()
button.pack(expand = False)
frame.pack(fill = BOTH, expand = True)

def update_labelTime():
    global labelTime
    labelTime.config(text = datetime.datetime.now().time().strftime("%I:%M %p"))
    labelTime.after(1000, update_labelTime)

def update_labelDate():
    global labelDate
    labelTime.config(text = datetime.datetime.now().date())
    labelTime.after(1000, update_labelDate)

eventNames = ["Free-Time", "Wake Up", "Eat Breakfst", "Take Shower", "Eat Lunch", " Take Nap", "Eat Snacks", "Complete Work", "Take Dinner", "Go to Sleep"]
eventTimings = {1:("6:00","6:05"), 2:("7:00","7:15"), 3:("12:00","12:10"), 4:("12:30","12:45"), 5:("14:30","17:30"),
                6:("18:30","18:40"), 7:("19:00","20:30"), 8:("21:30","22:00"), 9:("23:30","6:00")}

def findEvent(time):
    
    for eventNumber in eventTimings:
        if datetime.datetime.strptime(eventTimings[eventNumber][0], "%H:%M").time() < time and datetime.datetime.strptime(eventTimings[eventNumber][1], "%H:%M").time() > time:
            return eventNumber
    return 0


def update_labelEvent():
    global labelEvent
    global currentEvent

    currentEvent = eventNames[findEvent(datetime.datetime.now().time())]
    labelEvent.config(text = currentEvent)
    labelEvent.after(1000, update_labelEvent)

frame2 = Frame(root1, bg = "black")
labelDate = Label(frame2, text = datetime.datetime.now().date(), bg = "black", fg = "white")
labelDate.pack(side = LEFT)
labelDate.after(1000, update_labelDate)

currentEvent = "Free-Time"
labelEvent = Label(frame2, text = currentEvent, bg = "green", fg = "white", width = 25)
labelEvent.pack(side = LEFT, fill = X)
labelEvent.after(1000, update_labelEvent)

labelTime = Label(frame2, text = datetime.datetime.now().time().strftime("%I:%M %p"), bg = "black", fg = "white")
labelTime.pack(side = RIGHT)
labelTime.after(1000, update_labelTime)

frame2.pack(side = BOTTOM, fill = X)
root1.mainloop()  

# sencond window
root = Tk()
root.title("Sheet")
#file_location = "E:\\Book1.xlsx"
workbook = xlrd.open_workbook(file_location)
sheet = workbook.sheet_by_index(0)
data = [[sheet.cell_value(r,c) for c in range(sheet.ncols)] for r in range(sheet.nrows)]

for i in range(sheet.nrows):
    for j in range(sheet.ncols):
        if type(data[i][j]) == float:
            if int(data[i][j]) == data[i][j]:
                data[i][j] = int(data[i][j])

for i in range(len(data)):
    if i == 0:
        data[i] = ["No  "] + data[i]
    else:
        data[i] = [str(i) + "        "] + data[i]

colStarts = [v for v in range(0,len(data[0])*50,50)]

x = len(data[0])*100
y = len(data)*40
#root.geometry(str(x) + "x" + str(y))
app = Application(root)
root.mainloop()