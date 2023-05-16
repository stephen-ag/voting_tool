###------- this macro works for the files inloacation C:\Users\stephen.arput\Documents\RESULTS\LOC222\LOC2--
###-------file path customization needs to be done...in later version...
####------- Error handling has to be taken in the later version.....

from tkinter import filedialog
from PIL import ImageTk, Image
import PIL.Image
import tkinter as tk
import os, time
import io
import xlsxwriter
from xlsxwriter.utility import xl_rowcol_to_cell
import pandas as pd
from tkinter import *
from tkinter import ttk
#import page2
#from execute_macro import execute

root = Tk()
root.geometry("1600x1600+20+40")
# root.attributes('-fullscreen', True)
root['bg'] = '#CDCDBA'
root['bd'] = 3
# image resizing
global cnt1_sel1

def toplevel2():
    #execfile("page2.py")
    import page2
    page2.top()



def toplevel():
    global cnt1_sel1

    def close():
        top.destroy()

    def popupmsg(msg):
        popup = tk.Tk()
        popup.wm_title("Submit window")
        popup.geometry("200x60+20+40")
        label = ttk.Label(popup, text=msg, font=LARGE_FONT)
        label.pack(side="bottom", fill="x", pady=5)
        B1 = ttk.Button(popup, text="Okay", command=popup.destroy)
        B1.pack()
        popup.mainloop()

    cnt_sel1, cnt_sel2, cnt_sel3 = [0, 0, 0]
    result1 = []
    result2 = []
    result3 = []

    def select1():
        global cnt_sel1
        cnt_sel1 = cnt_sel1 + 1
        print(" total count of select 1 page2 = ", cnt_sel1)
        result1.append(cnt_sel1)
        print(result1)
        popupmsg('   !!! SUBMITTED !!!   ')
        return (cnt_sel1)

    def select2():
        global cnt_sel2
        cnt_sel2 = cnt_sel2 + 1
        print(" total count of select 2 page2  = ", cnt_sel2)
        result2.append(cnt_sel2)
        print(result2)
        popupmsg('   !!! SUBMITTED !!!   ')

    def select3():
        global cnt_sel3
        cnt_sel3 = cnt_sel3 + 1
        print(" total count of select 3 page2 = ", cnt_sel3)
        result3.append(cnt_sel3)
        print(result3)
        popupmsg('   !!! SUBMITTED !!!   ')

    print("entering button controls")

    global new_img1
    top = Toplevel()
    top.geometry("1200x1200+20+40")
    top.attributes('-fullscreen', True)
    top.title(" Second page of vote tool ")

    img = PIL.Image.open("C:/Users/arpuste/PycharmProjects/pythonProject1/vote1.jpg")
    resize = img.resize((400, 300))

    new_img1 = ImageTk.PhotoImage(resize)

    # panel = Label(root, image = new_img, height = 400, width = 300)
    panel = Label(top, image=new_img1)
    panel.place(x=1000, y=350)
    # oldlace'azure'

    Label(top, text="DMCS Banaswadi, Election Tool v1.0 ", bg="#8B8B25", height="2", \
          width="800", fg="white",
          font=("Calibri", 40)).pack()
    Label(top, text="Note: This tool is specific to project requirement, read the requirement\
        of this tool for process and input data. Files with binary excel format not supported",
          height="3",
          width="400",
          font=("Calibri", 12)).pack()
    print("entering post processing module")

    button7 = Button(top, text=" Exit ", height="1", width="20", \
                     font=("Calibri", 13), bg="gray21", fg="white", command=close)
    button7.place(x=1100, y=660)
    # ---------------------------------

    button1 = Button(top, text="  Select 1", height="2", width="25", \
                     font=("Calibri", 13), bg="light grey", fg="black", command=select1)
    button1.place(x=60, y=370)

    button11 = Button(top, text="  Select 11", height="2", width="25", \
                      font=("Calibri", 13), bg="light grey", fg="black", command=select2)
    button11.place(x=360, y=370)

    button111 = Button(top, text="  Select 111", height="2", width="25", \
                       font=("Calibri", 13), bg="light grey", fg="black", command=select3)
    button111.place(x=660, y=370)



##img = Image.open("C:/Users/arpuste/Downloads/Post-processing-Tool-main/path.png")
img = PIL.Image.open("C:/Users/arpuste/PycharmProjects/pythonProject1/vote.jpg")
resize = img.resize((400, 300))

new_img = ImageTk.PhotoImage(resize)

# panel = Label(root, image = new_img, height = 400, width = 300)
panel = Label(root, image=new_img)
panel.place(x=1000, y=350)
# oldlace'azure'
root.title('DMCS election Tool  ')
Label(root, text="DMCS Banaswadi, Election Tool v1.0 ", bg="#8B8B25", height="2", \
      width="800", fg="white",
      font=("Calibri", 40)).pack()
#Label(root, text="DMCS Voting Tool",
#      height="3",
#      width="400",
#      font=("Calibri", 12)).pack()

print("entering post processing module")
#time.sleep(1)
#Label.move()

LARGE_FONT= ("Verdana", 12)
NORM_FONT = ("Helvetica", 10)
SMALL_FONT = ("Helvetica", 8)
result1 = []
result2 = []
result3 = []

def popupmsg(msg):
    popup = tk.Tk()
    popup.wm_title("Submit window")
    popup.geometry("200x60+20+40")
    label = ttk.Label(popup, text=msg, font=LARGE_FONT)
    label.pack(side="bottom", fill="x", pady=5)
    B1 = ttk.Button(popup, text="Okay", command = popup.destroy)
    B1.pack()
    popup.mainloop()
cnt_sel1,cnt_sel2,cnt_sel3 =[0,0,0]
def select1():
    global cnt_sel1
    cnt_sel1= cnt_sel1 + 1
    print(" total count of select 1 = ",cnt_sel1)
    result1.append(cnt_sel1)
    print(result1)
    popupmsg('   !!! SUBMITTED !!!   ')
    return (cnt_sel1)

def select2():
    global cnt_sel2
    cnt_sel2 = cnt_sel2 + 1
    print(" total count of select 2 = ", cnt_sel2)
    result2.append(cnt_sel2)
    print(result2)
    popupmsg('   !!! SUBMITTED !!!   ')

def select3():
    global cnt_sel3
    cnt_sel3 = cnt_sel3 + 1
    print(" total count of select 3= ", cnt_sel3)
    result3.append(cnt_sel3)
    print(result3)
    popupmsg('   !!! SUBMITTED !!!   ')


def execute():
      pass


def close1():

    root.destroy()
print("entering button controls")
# ---------------------------------

button1 = Button(root, text="  Select 1", height="2", width="25", \
                 font=("Calibri", 13), bg="light grey", fg="black", command=select1)
button1.place(x=60, y=370)

button11 = Button(root, text="  Select 11", height="2", width="25", \
                 font=("Calibri", 13), bg="light grey", fg="black", command=select2)
button11.place(x=360, y=370)

button111 = Button(root, text="  Select 111", height="2", width="25", \
                 font=("Calibri", 13), bg="light grey", fg="black", command=select3)
button111.place(x=660, y=370)



# ---------------------------------
button2 = Button(root, text="  Select 2 ", height="2", width="25", \
                 font=("Calibri", 13), bg="dodgerblue3", fg="white", command=execute)
button2.place(x=60, y=600)
# --------------------------------
# --------------------------------
button22 = Button(root, text="  Select 22 ", height="2", width="25", \
                 font=("Calibri", 13), bg="dodgerblue3", fg="white", command=execute)
button22.place(x=360, y=600)
# --------------------------------
button222 = Button(root, text="  Select 222 ", height="2", width="25", \
                 font=("Calibri", 13), bg="dodgerblue3", fg="white", command=execute)
button222.place(x=660, y=600)
# ---------------------------------
# --------------------------------
page1= Button(root, text=" Page1 ", height="1", width="10", \
                 font=("Calibri", 13), bg="gray21", fg="white", command=toplevel2)
page1.place(x=1000, y=660)
# --------------------------------
# --------------------------------
page2 = Button(root, text=" Page2 ", height="1", width="10", \
                 font=("Calibri", 13), bg="gray21", fg="white", command=toplevel2)
page2.place(x=1100, y=660)

# --------------------------------
page3 = Button(root, text=" Page3 ", height="1", width="10", \
                 font=("Calibri", 13), bg="gray21", fg="white", command=toplevel2)
page3.place(x=1200, y=660)

Exit = Button(root, text=" Exit ", height="1", width="20", \
                 font=("Calibri", 13), bg="gray21", fg="white", command=close1)
Exit.place(x=1100, y=740)


page4 = Button(root, text=" Page4 ", height="1", width="10", \
                 font=("Calibri", 13), bg="gray21", fg="white", command=toplevel)
page4.place(x=1300, y= 660)
# --------------------------------
page5 = Button(root, text=" Page5 ", height="1", width="10", \
                 font=("Calibri", 13), bg="gray21", fg="white", command=toplevel)
page5.place(x=1000, y= 700)
# --------------------------------
# --------------------------------
page6 = Button(root, text=" Page6 ", height="1", width="10", \
                 font=("Calibri", 13), bg="gray21", fg="white", command=toplevel2)
page6.place(x=1100, y=700)

# --------------------------------
page7 = Button(root, text=" Page7 ", height="1", width="10", \
                 font=("Calibri", 13), bg="gray21", fg="white", command=toplevel2)
page7.place(x=1200, y=700)


page8 = Button(root, text=" Page8 ", height="1", width="10", \
                 font=("Calibri", 13), bg="gray21", fg="white", command=toplevel)
page8.place(x=1300, y= 700)
# --------------------------------
#label = Label(root)



print("completed button controls")
root.state("zoomed")
root.mainloop()

#!!!!!!!!!!!!!!!! writing it to file
with open('test2.txt', 'a') as file2:
    # write contents to the test2.txt file
    file2.write('{} select1\n'.format(cnt_sel1))
    file2.write('{} select2\n'.format(cnt_sel2))
    file2.write('{} select3\n'.format(cnt_sel3))
    #file2.write(str(cnt1_sel1))
    file2.write('Results of Iteration\n')
    file2.close()

#!!!!!!!!!!!!!!!! writing it to excel file
name=['select1 count','select2 count','select3 count']
lis=[cnt_sel1,cnt_sel2,cnt_sel3]
d = dict(zip(name,lis))
print(d)
df=pd.DataFrame(list(d.items()),columns = ['Category','count'])
print(df)

df.to_excel('Results_to_excel.xlsx', sheet_name='new_sheet_name')
