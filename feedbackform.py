from tkinter import *
import csv
import datetime as dt
import tkinter as tk
import xlwt
import xlrd
import xlsxwriter
from xlwt import Workbook,Formula
import openpyxl
workbook = xlwt.Workbook() #creating workbook

sheet1 = workbook.add_sheet("Feedbackform", cell_overwrite_ok=True) #creating sheet name as feedbackfrom and we can overwrite over the cell
style = xlwt.easyxf('font: bold 1') #to make data as bold
sheet1.write(0, 0, "Training Conducted By:",style) # writing data to particular cell
sheet1.write(0, 2, "Subject:",style)
sheet1.write(1,2,"Review",style)
sheet1.write(0, 4, "Date:",style)
sheet1.write(0, 6, "External/Internal:",style)
sheet1.write(1,0, "Course Content:",style)
sheet1.write(2,0, "1. The course schedule was laid out in a precise manner.")
sheet1.write(3,0,"2. The course length was sufficient to deliver the content.") 
sheet1.write(4,0,"3. The handouts were relevant and useful.")
sheet1.write(5,0, "Instructor:",style)
sheet1.write(6,0, "1. The instructor was clear and understandable.")
sheet1.write(7,0,"2. The instructor invited participation from the class.") 
sheet1.write(8,0,"3. The instructor was responsive to questions and comments.")
sheet1.write(9,0,"4. The instructor was well prepared for the class.")
sheet1.write(10,0,"5. The instructor has good knowledge about the subject.")
sheet1.write(11,0, "Exercises:",style)
sheet1.write(12,0, "1. The exercises reinforced the course content.")
sheet1.write(13,0,"2. Sufficient time was provided to complete the exercises.") 
sheet1.write(14,0,"3. Questions pertaining to the exercises were answered in a timely manner.")
sheet1.write(15,0, "Facilities:",style)
sheet1.write(16,0, "1. Adequate facilities were provided to facilitate learning.")
sheet1.write(17,0,"2. Machines were responsive and adequate for the exercises provided.")
sheet1.write(18,0, "Training Benefit:",style)
sheet1.write(19,0, "1. The training met my objectives.")
sheet1.write(20,0,"2. I will be able to apply the knowledge that I learned. ") 
sheet1.write(21,0,"3. I will recommend this training to my colleagues.")
sheet1.write(22,0,"Any Suggestions / Comments:",style)
sheet1.col(0).width=17000 # modifying width of the cell
sheet1.col(1).width=7000
sheet1.col(3).width=7000
sheet1.col(6).width=5000
sheet1.col(5).width=4000
sheet1.col(7).width=3000

w=Tk()

frame=Frame(w,bd=1,relief=SUNKEN) #creating frame
frame.grid(row=0,column=0)

Label(frame,text="Training Feedback form\n",fg='white',bg='blue',font=("Verdana 14 bold")).grid(row=0,column=0)

frame1=Frame(w,bd=1,relief=SUNKEN)
frame1.grid(row=1,column=0)

frame2=Frame(w,bd=1,relief=SUNKEN)
frame2.grid(row=2,column=0)

frame3=Frame(w,bd=1,relief=SUNKEN)
frame3.grid(row=3,column=0)

Label(frame1,text="Training conducted by\n",fg='black',bg='white',height=1,width=18).grid(sticky=W,row=1)
list1=("Sharathkumar","Rajan","Shivaraj")
var1=StringVar(frame1)#assigining variable
var1.set("Select Trainer")#setting variable name before selecting 

def sel(event):
    value=var1.get() #for getting selected value
    print(value) #printing selected value
    sheet1.write(0,1, value)#writing to the particular cell
    workbook.save('feedbackform.xls')#saving the sheet
      
opt1=OptionMenu(frame1,var1,*list1,command=sel)
opt1.grid(row=1,column=1)

l2=Label(frame1,text="Subject\n",fg='black',bg='white',height=1,width=14)
l2.grid(sticky=W,row=2)

list2=("Android","Python","Java")

var2=StringVar(frame1)
var2.set("Select Subject")

def sel(event):
    value1=var2.get()
    print(value1)
    sheet1.write(0,3,value1)
    
opt2=OptionMenu(frame1,var2,*list2,command = sel)
opt2.grid(row=2,column=1)

Label(frame1,text="Date",fg='black',bg='white',height=1,width=5).grid(row=1,column=5)

e2 = Entry(frame1, width=10)
e2.grid(row=1,column=6)

def calling(event):
    value=e2.get()
    print(value)
    sheet1.write(0,5, value)
    workbook.save('feedbackform.xls')

    '''def date(event):
        #value = e1.get()
        if value == '':
           messagebox.showwarning("Date should not be empty")
        elif value.isstr():
            messagebox.showwarning("Enter valid Date")'''
e2.bind("<Return>", calling)

x=IntVar()

def pass1():
    var=x.get()
    print(var)
    sheet1.write(0,7, var)
    workbook.save('feedbackform.xls')
    
Radiobutton(frame1,text="External",variable=x,value=1,command=pass1).grid(row=2,column=5)
Radiobutton(frame1,text="Internal",variable=x,value=2,command=pass1).grid(row=2,column=6)


Label(frame2,text="Please rate the training on a scale of 1 - 5 (one being the lowest and five being the highest). You do not need to put your name on this form – your responses are anonymous.\n",fg='black',bg='white').grid(column=0,row=3)

Label(frame3,text="Course content\n",fg='black',bg='white',font="Verdana 14 bold").grid(sticky=W,row=4)
Label(frame3,text="1. The course schedule was laid out in a precise manner.\n ",fg='black',bg='white').grid(sticky=W,row=5)

a = IntVar()

def pass1():
    var1=a.get()
    print(var1)
    sheet1.write(2,2, var1)
    workbook.save('feedbackform.xls')

Radiobutton(frame3, text="1", variable=a, value=1,command=pass1).grid(row=5,column=1)
Radiobutton(frame3, text="2", variable=a, value=2,command=pass1).grid(row=5,column=2)
Radiobutton(frame3, text="3", variable=a, value=3,command=pass1).grid(row=5,column=3)
Radiobutton(frame3, text="4", variable=a, value=4,command=pass1).grid(row=5,column=4)
Radiobutton(frame3, text="5", variable=a, value=5,command=pass1).grid(row=5,column=5)
        
Label(frame3,text="2. The course length was sufficient to deliver the content.\n",fg='black',bg='white').grid(sticky=W,row=6) 

b=IntVar()

def pass1():
    var2=b.get()
    print(var2)
    sheet1.write(3,2, var2)
    workbook.save('feedbackform.xls')
    
Radiobutton(frame3, text="1", variable=b, value=1,command=pass1).grid(row=6,column=1)
Radiobutton(frame3, text="2", variable=b, value=2,command=pass1).grid(row=6,column=2)
Radiobutton(frame3, text="3", variable=b, value=3,command=pass1).grid(row=6,column=3)
Radiobutton(frame3, text="4", variable=b, value=4,command=pass1).grid(row=6,column=4)
Radiobutton(frame3, text="5", variable=b, value=5,command=pass1).grid(row=6,column=5)

Label(frame3,text="3. The handouts were relevant and useful.\n",fg='black',bg='white').grid(sticky=W,row=7)

c=IntVar()
def pass1():
    var3=c.get()
    print(var3)
    sheet1.write(4,2, var3)
    workbook.save('feedbackform.xls')
    
Radiobutton(frame3, text="1", variable=c, value=1,command=pass1).grid(row=7,column=1)
Radiobutton(frame3, text="2", variable=c, value=2,command=pass1).grid(row=7,column=2)
Radiobutton(frame3, text="3", variable=c, value=3,command=pass1).grid(row=7,column=3)
Radiobutton(frame3, text="4", variable=c, value=4,command=pass1).grid(row=7,column=4)
Radiobutton(frame3, text="5", variable=c, value=5,command=pass1).grid(row=7,column=5)

Label(frame3,text="Instructor\n",fg='black',bg='white',font="Verdana 14 bold").grid(row=8,sticky=W)

Label(frame3,text="1.The instructor was clear and understandable.\n",fg='black',bg='white').grid(row=9,sticky=W)

d=IntVar()
def pass1():
    var4=d.get()
    print(var4)
    sheet1.write(6,2, var4)
    workbook.save('feedbackform.xls')
Radiobutton(frame3, text="1", variable=d, value=1,command=pass1).grid(row=9,column=1)
Radiobutton(frame3, text="2", variable=d, value=2,command=pass1).grid(row=9,column=2)
Radiobutton(frame3, text="3", variable=d, value=3,command=pass1).grid(row=9,column=3)
Radiobutton(frame3, text="4", variable=d, value=4,command=pass1).grid(row=9,column=4)
Radiobutton(frame3, text="5", variable=d, value=5,command=pass1).grid(row=9,column=5)

Label(frame3,text="2. The instructor invited participation from the class.\n",fg='black',bg='white').grid(row=10,sticky=W)

e=IntVar()
def pass1():
    var5=e.get()
    print(var5)
    sheet1.write(7,2, var5)
    workbook.save('feedbackform.xls')
Radiobutton(frame3, text="1", variable=e, value=1,command=pass1).grid(row=10,column=1)
Radiobutton(frame3, text="2", variable=e, value=2,command=pass1).grid(row=10,column=2)
Radiobutton(frame3, text="3", variable=e, value=3,command=pass1).grid(row=10,column=3)
Radiobutton(frame3, text="4", variable=e, value=4,command=pass1).grid(row=10,column=4)
Radiobutton(frame3, text="5", variable=e, value=5,command=pass1).grid(row=10,column=5)

Label(frame3,text="3. The instructor was responsive to questions and comments.\n",fg='black',bg='white').grid(row=11,sticky=W)

f=IntVar()
def pass1():
    var6=f.get()
    print(var6)
    sheet1.write(8,2, var6)
    workbook.save('feedbackform.xls')
Radiobutton(frame3, text="1", variable=f, value=1,command=pass1).grid(row=11,column=1)
Radiobutton(frame3, text="2", variable=f, value=2,command=pass1).grid(row=11,column=2)
Radiobutton(frame3, text="3", variable=f, value=3,command=pass1).grid(row=11,column=3)
Radiobutton(frame3, text="4", variable=f, value=4,command=pass1).grid(row=11,column=4)
Radiobutton(frame3, text="5", variable=f, value=5,command=pass1).grid(row=11,column=5)

Label(frame3,text="4. The instructor was well prepared for the class.\n",fg='black',bg='white').grid(row=12,sticky=W)

g=IntVar()
def pass1():
    var7=g.get()
    print(var7)
    sheet1.write(9,2, var7)
    workbook.save('feedbackform.xls')
Radiobutton(frame3, text="1", variable=g, value=1,command=pass1).grid(row=12,column=1)
Radiobutton(frame3, text="2", variable=g, value=2,command=pass1).grid(row=12,column=2)
Radiobutton(frame3, text="3", variable=g, value=3,command=pass1).grid(row=12,column=3)
Radiobutton(frame3, text="4", variable=g, value=4,command=pass1).grid(row=12,column=4)
Radiobutton(frame3, text="5", variable=g, value=5,command=pass1).grid(row=12,column=5)

Label(frame3,text="5. The instructor has good knowledge about the subject.\n",fg='black',bg='white').grid(row=13,sticky=W)

h=IntVar()
def pass1():
    var8=h.get()
    print(var8)
    sheet1.write(10,2, var8)
    workbook.save('feedbackform.xls')
Radiobutton(frame3, text="1", variable=h, value=1,command=pass1).grid(row=13,column=1)
Radiobutton(frame3, text="2", variable=h, value=2,command=pass1).grid(row=13,column=2)
Radiobutton(frame3, text="3", variable=h, value=3,command=pass1).grid(row=13,column=3)
Radiobutton(frame3, text="4", variable=h, value=4,command=pass1).grid(row=13,column=4)
Radiobutton(frame3, text="5", variable=h, value=5,command=pass1).grid(row=13,column=5)

Label(frame3,text="Exercises.\n",fg='black',bg='white',font="Verdana 14 bold").grid(row=14,sticky=W)

Label(frame3,text="1. The exercises reinforced the course content.\n",fg='black',bg='white').grid(row=15,sticky=W)

i=IntVar()
def pass1():
    var9=i.get()
    print(var9)
    sheet1.write(12,2, var9)
    workbook.save('feedbackform.xls')
Radiobutton(frame3, text="1", variable=i, value=1,command=pass1).grid(row=15,column=1)
Radiobutton(frame3, text="2", variable=i, value=2,command=pass1).grid(row=15,column=2)
Radiobutton(frame3, text="3", variable=i, value=3,command=pass1).grid(row=15,column=3)
Radiobutton(frame3, text="4", variable=i, value=4,command=pass1).grid(row=15,column=4)
Radiobutton(frame3, text="5", variable=i, value=5,command=pass1).grid(row=15,column=5)

Label(frame3,text="2. Sufficient time was provided to complete the exercises.\n",fg='black',bg='white').grid(row=16,sticky=W)

j=IntVar()
def pass1():
    var10=j.get()
    print(var10)
    sheet1.write(13,2, var10)
    workbook.save('feedbackform.xls')
Radiobutton(frame3, text="1", variable=j, value=1,command=pass1).grid(row=16,column=1)
Radiobutton(frame3, text="2", variable=j, value=2,command=pass1).grid(row=16,column=2)
Radiobutton(frame3, text="3", variable=j, value=3,command=pass1).grid(row=16,column=3)
Radiobutton(frame3, text="4", variable=j, value=4,command=pass1).grid(row=16,column=4)
Radiobutton(frame3, text="5", variable=j, value=5,command=pass1).grid(row=16,column=5)

Label(frame3,text="3. Questions pertaining to the exercises were answered in a timely manner.\n",fg='black',bg='white').grid(row=17,sticky=W)

k=IntVar()
def pass1():
    var11=k.get()
    print(var11)
    sheet1.write(14,2, var11)
    workbook.save('feedbackform.xls')
Radiobutton(frame3, text="1", variable=k, value=1,command=pass1).grid(row=17,column=1)
Radiobutton(frame3, text="2", variable=k, value=2,command=pass1).grid(row=17,column=2)
Radiobutton(frame3, text="3", variable=k, value=3,command=pass1).grid(row=17,column=3)
Radiobutton(frame3, text="4", variable=k, value=4,command=pass1).grid(row=17,column=4)
Radiobutton(frame3, text="5", variable=k, value=5,command=pass1).grid(row=17,column=5)

Label(frame3,text="Facilities.\n",fg='black',bg='white',font="Verdana 14 bold").grid(row=18,sticky=W)
            
Label(frame3,text="1. Adequate facilities were provided to facilitate learning.\n",fg='black',bg='white').grid(row=19,sticky=W)

l=IntVar()
def pass1():
    var12=l.get()
    print(var12)
    sheet1.write(16,2, var12)
    workbook.save('feedbackform.xls')
Radiobutton(frame3, text="1", variable=l, value=1,command=pass1).grid(row=19,column=1)
Radiobutton(frame3, text="2", variable=l, value=2,command=pass1).grid(row=19,column=2)
Radiobutton(frame3, text="3", variable=l, value=3,command=pass1).grid(row=19,column=3)
Radiobutton(frame3, text="4", variable=l, value=4,command=pass1).grid(row=19,column=4)
Radiobutton(frame3, text="5", variable=l, value=5,command=pass1).grid(row=19,column=5)

Label(frame3,text="2. Machines were responsive and adequate for the exercises provided.\n",fg='black',bg='white').grid(row=20,sticky=W)

m=IntVar()
def pass1():
    var13=m.get()
    print(var13)
    sheet1.write(17,2, var13)
    workbook.save('feedbackform.xls')
Radiobutton( frame3, text="1", variable=m, value=1,command=pass1).grid(row=20,column=1)
Radiobutton( frame3, text="2", variable=m, value=2,command=pass1).grid(row=20,column=2)
Radiobutton( frame3, text="3", variable=m, value=3,command=pass1).grid(row=20,column=3)
Radiobutton(frame3, text="4", variable=m, value=4,command=pass1).grid(row=20,column=4)
Radiobutton( frame3, text="5", variable=m, value=5,command=pass1).grid(row=20,column=5)

Label( frame3,text="Training Benefits\n",fg='black',bg='white',font="Verdana 14 bold").grid(row=21,sticky=W)

Label( frame3,text="1. The training met my objectives.\n",fg='black',bg='white').grid(row=22,sticky=W)

n=IntVar()
def pass1():
    var14=n.get()
    print(var14)
    sheet1.write(19,2, var14)
    workbook.save('feedbackform.xls')
Radiobutton(frame3, text="1", variable=n, value=1,command=pass1).grid(row=22,column=1)
Radiobutton(frame3, text="2", variable=n, value=2,command=pass1).grid(row=22,column=2)
Radiobutton(frame3, text="3", variable=n, value=3,command=pass1).grid(row=22,column=3)
Radiobutton(frame3, text="4", variable=n, value=4,command=pass1).grid(row=22,column=4)
Radiobutton(frame3, text="5", variable=n, value=5,command=pass1).grid(row=22,column=5)

Label(frame3,text="2. I will be able to apply the knowledge that I learned.\n",fg='black',bg='white').grid(row=23,sticky=W)

o=IntVar()
def pass1():
    var15=o.get()
    print(var15)
    sheet1.write(20,2, var15)
    workbook.save('feedbackform.xls')
Radiobutton(frame3, text="1", variable=o, value=1,command=pass1).grid(row=23,column=1)
Radiobutton(frame3, text="2", variable=o, value=2,command=pass1).grid(row=23,column=2)
Radiobutton(frame3, text="3", variable=o, value=3,command=pass1).grid(row=23,column=3)
Radiobutton(frame3, text="4", variable=o, value=4,command=pass1).grid(row=23,column=4)
Radiobutton(frame3, text="5", variable=o, value=5,command=pass1).grid(row=23,column=5)

Label(frame3,text="3. I will recommend this training to my colleagues.\n",fg='black',bg='white').grid(row=24,sticky=W)

p=IntVar()
def pass1():
    var16=p.get()
    print(var16)
    sheet1.write(21,2, var16)
    workbook.save('feedbackform.xls')
Radiobutton(frame3, text="1", variable=p, value=1,command=pass1).grid(row=24,column=1)
Radiobutton(frame3, text="2", variable=p, value=2,command=pass1).grid(row=24,column=2)
Radiobutton(frame3, text="3", variable=p, value=3,command=pass1).grid(row=24,column=3)
Radiobutton(frame3, text="4", variable=p, value=4,command=pass1).grid(row=24,column=4)
Radiobutton(frame3, text="5", variable=p, value=5,command=pass1).grid(row=24,column=5)

Label(frame3,text="Any Suggestions\n",fg='black',bg='white',font="Verdana 14 bold").grid(row=25,sticky=W)

e1 = Entry(frame3, width=100)
e1.grid(row=26,sticky=W)
def calling(event):
    value=e1.get()
    sheet1.write(22,1, value)
    workbook.save('feedbackform.xls')
e1.bind("<Return>", calling)#To access keyboard enter
#e.bind("<Tab>",passw)
'''def passw(event):
       value=e.get()
       workbook = xlwt.Workbook()
       sheet1 = workbook.add_sheet('Feedbackform', formatting_info=True)
       sheet1.write(0,1, value)
       workbook.save('feedbackform.xls')'''   


w.mainloop()

workbook.save("feedbackform.xls")



