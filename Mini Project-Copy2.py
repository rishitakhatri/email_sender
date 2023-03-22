#!/usr/bin/env python
# coding: utf-8

# In[24]:


#!/usr/bin/env python
# coding: utf-8

# In[ ]:


from tkinter import *
from tkinter import messagebox,filedialog
from pygame import mixer
from email.message import EmailMessage
import smtplib
import os
import imghdr
import pandas
import pandas.util
from PIL import ImageTk
import PIL

check=False

class Student:
    def __init__(self, roll,first_name,name, email,attendance):
        self.roll=roll
        self.first_name=first_name
        self.name = name
        self.email=email
        self.attendance = attendance
        
    def isDefault(self):
        if self.attendance < 75 :
            return True
        else:
            return False

def browse():
    global final_emails
    global student_list
    path=filedialog.askopenfilename(initialdir='c:/',title='Select Excel File')
    if path=='':
        messagebox.showerror('Error','Please select an Excel File')

    else:
        data=pandas.read_excel(path)
        #preparing student list
        student_list = []
        for i in data['roll']:
            student_list.append(Student(i,data['first_name'][i-1],data['Name'][i-1],data['Email'][i-1],data['attendance'][i-1]))
            
        if 'Email' in data.columns:
            
            final_emails=[]
            

            for s in student_list:
                final_emails.append(s.email)   
            if len(final_emails)==0:
                    
                messagebox.showerror('Error','File does not contain any email addresses')

            else:
                toEntryField.config(state=NORMAL)
                toEntryField.insert(0,os.path.basename(path))
                toEntryField.config(state='readonly')
                totalLabel.config(text='Total: '+str(len(final_emails)))
                sentLabel.config(text='Sent:')
                leftLabel.config(text='Left:')
                failedLabel.config(text='Failed:')







def button_check():
    if choice.get()=='multiple':
        browseButton.config(state=NORMAL)
        toEntryField.config(state='readonly')

    if choice.get()=='single':
        browseButton.config(state=DISABLED)
        toEntryField.config(state=NORMAL)



def attachment():
    global filename,filetype,filepath,check
    check=True

    filepath=filedialog.askopenfilename(initialdir='c:/',title='Select File')
    filetype=filepath.split('.')
    filetype=filetype[1]
    filename=os.path.basename(filepath)
    textarea.insert(END,f'\n{filename}\n')



def sendingEmail(toAddress,subject,body):
    f=open('credentials.txt','r')
    for i in f:
        credentials=i.split(',')

    message=EmailMessage()
    message['subject']=subject
    message['to']=toAddress
    message['from']=credentials[0]
    message.set_content(body)
    if check:
        if filetype=='png' or filetype=='jpg' or filetype=='jpeg':
            f=open(filepath,'rb')
            file_data=f.read()
            subtype=imghdr.what(filepath)


            message.add_attachment(file_data,maintype='image',subtype=subtype,filename=filename)

        else:
            f = open(filepath, 'rb')
            file_data = f.read()
            message.add_attachment(file_data,maintype='application',subtype='octet-stream',filename=filename)


    s=smtplib.SMTP('smtp.gmail.com',587)
    s.starttls()
    s.login(credentials[0],credentials[1])
    s.send_message(message)
    x=s.ehlo()
    if x[0]==250:
        return 'sent'
    else:
        return 'failed'






def send_email():

            sent=0
            failed=0
            for x in student_list:
                if x.isDefault()==True:
                    result=sendingEmail(x.email,'Hello ' + str(x.first_name) + ' , Kindly Check your Attendance status','Dear '+ x.name+',' +'\n        This is to inform you that your attendance is ' + str(x.attendance) +'%'+ ' which is less than compared to the requirement. Kindly approach your class teacher for further procedure else you would not able to sit for semester exams. All the best for your exams. '+'\nRoll Number: '+str(x.roll)+'\nStudent Name: '+x.name+'\nEmail ID: '+x.email+'\nAverage Attendance: '+str(x.attendance)+'\nAttendance status: Defaulter' + '\n\n\nThank you and Regards.')
                else:
                    
                    result=sendingEmail(x.email,'Hello ' + str(x.first_name) + ' , Kindly Check your Attendance status','Dear ' + x.name +','+'\n       This is to inform you that your attendance is ' + str(x.attendance) +'%'+ ' which satisfies the requirement. You are not a defaulter. All The best for your exams.'+'\nRoll Number: '+str(x.roll)+'\nStudent Name: '+x.name+'\nEmail ID: '+x.email+'\nAverage Attendance: '+str(x.attendance)+'\nAttendance status: Not Defaulter' + '\n\nThank you and Regards.')
                if result=='sent':
                    sent+=1
                if result=='failed':
                    failed+=1

                totalLabel.config(text='')
                sentLabel.config(text='Sent:' + str(sent))
                leftLabel.config(text='Left:' + str(len(final_emails) - (sent + failed)))
                failedLabel.config(text='Failed:' + str(failed))

                totalLabel.update()
                sentLabel.update()
                leftLabel.update()
                failedLabel.update()

            messagebox.showinfo('Success','Emails are sent successfully')




def settings():
    def clear1():
        fromEntryField.delete(0,END)
        passwordEntryField.delete(0,END)

    def save():
        if fromEntryField.get()=='' or passwordEntryField.get()=='':
            messagebox.showerror('Error','All Fields Are Required',parent=root1)

        else:
            f=open('credentials.txt','w')
            f.write(fromEntryField.get()+','+passwordEntryField.get())
            f.close()
            messagebox.showinfo('Information','CREDENTIALS SAVED SUCCESSFULLY',parent=root1)

    root1=Toplevel()
    root1.title('Setting')
    root1.geometry('650x340+350+90')

    root1.config(bg='dodger blue2')

    Label(root1,text='Credential Settings',image=logoImage,compound=LEFT,font=('goudy old style',40,'bold'),
          fg='white',bg='gray20').grid(padx=60)

    fromLabelFrame = LabelFrame(root1, text='From (Email Address)', font=('times new roman', 16, 'bold'), bd=5, fg='white',
                              bg='dodger blue2')
    fromLabelFrame.grid(row=1, column=0,pady=20)

    fromEntryField = Entry(fromLabelFrame, font=('times new roman', 18, 'bold'), width=30)
    fromEntryField.grid(row=0, column=0)

    passwordLabelFrame = LabelFrame(root1, text='Password', font=('times new roman', 16, 'bold'), bd=5,
                                fg='white',
                                bg='dodger blue2')
    passwordLabelFrame.grid(row=2, column=0, pady=20)

    passwordEntryField = Entry(passwordLabelFrame, font=('times new roman', 18, 'bold'), width=30,show='*')
    passwordEntryField.grid(row=0, column=0)

    Button(root1,text='SAVE',font=('times new roman',18,'bold'),cursor='hand2',bg='gold2',fg='black'
           ,command=save).place(x=210,y=280)
    Button(root1,text='CLEAR',font=('times new roman',18,'bold'),cursor='hand2',bg='gold2',fg='black'
           ,command=clear1).place(x=340,y=280)

    f=open('credentials.txt','r')
    for i in f:
        credentials=i.split(',')

    fromEntryField.insert(0,credentials[0])
    passwordEntryField.insert(0,credentials[1])







    root1.mainloop()


def iexit():
    result=messagebox.askyesno('Notification','Do you want to exit?')
    if result:
        root.destroy()
    else:
        pass

def clear():
    toEntryField.delete(0,END)
    
    
    
    subjectEntryField.delete(0,END)
    textarea.delete(1.0,END)




root=Tk()
root.title('Email sender app')
root.geometry('780x520+100+50')
root.resizable(0,0)
root.config(bg='dodger blue2')
img = PhotoImage(file="bg.png")
label = Label(
root,
image=img
)
label.place(x=0, y=0)
titleFrame=Frame(root,bg='white')
titleFrame.grid(row=0,column=0)
logoImage=PhotoImage(file='email.png')
titleLabel=Label(titleFrame,text='  Email Sender',image=logoImage,compound=LEFT,font=('Goudy Old Style',28,'bold'),
                 bg='white',fg='dodger blue2')
titleLabel.grid(row=0,column=0)
settingImage=PhotoImage(file='setting2.png')

Button(titleFrame,image=settingImage,bd=0,bg='white',cursor='hand2',activebackground='white'
       ,command=settings).grid(row=0,column=1,padx=20)

chooseFrame=Frame(root,bg='dodger blue2')
chooseFrame.grid(row=1,column=0,pady=10)
choice=StringVar()

choice.set('single')

toLabelFrame=LabelFrame(root,text='Select Excel File',font=('times new roman',16,'bold'),bd=5,fg='white',bg='dodger blue')
toLabelFrame.grid(row=2,column=0,padx=100)

toEntryField=Entry(toLabelFrame,font=('times new roman',18,'bold'),width=30)
toEntryField.grid(row=0,column=0)

browseImage=PhotoImage(file='browse.png')

browseButton=Button(toLabelFrame,text=' Browse',image=browseImage,compound=LEFT,font=('arial',12,'bold'),
       cursor='hand2',bd=0,fg='white' , bg='dodger blue',activebackground='blue',state=NORMAL,command=browse)
browseButton.grid(row=0,column=1,padx=20)

subjectLabelFrame=LabelFrame(root,text='---------------',font=('times new roman',16,'bold'),bd=5,fg='white',bg='dodger blue')
subjectLabelFrame.grid(row=3,column=100,pady=10)


emailLabelFrame=LabelFrame(root,text='Select Attachments',font=('times new roman',16,'bold'),bd=5,fg='white',bg='dodger blue')
emailLabelFrame.grid(row=4,column=0,padx=20)

attachImage=PhotoImage(file='browse.png')

Button(emailLabelFrame,text='Add Attachments',image=attachImage,compound=LEFT,font=('arial',12,'bold'),
       cursor='hand2',bd=0,fg='white',bg='dodger blue',activebackground='blue',command=attachment).grid(row=0,column=1)

textarea=Text(emailLabelFrame,font=('times new roman',14,),height=8)
textarea.grid(row=1,column=0,columnspan=2)

sendImage=PhotoImage(file='sendemail.png')
Button(root,image=sendImage,bd=0,bg="#1c86ee",cursor='hand2',activebackground='blue'
       ,command=send_email).place(x=490,y=450)

clearImage=PhotoImage(file='clear.png')
Button(root,image=clearImage,bd=0,bg="#1c86ee",cursor='hand2',activebackground='blue'
       ,command=clear).place(x=590,y=450)

exitImage=PhotoImage(file='exit.png')
Button(root,image=exitImage,bd=0,bg="#1c86ee",cursor='hand2',activebackground='blue'
       ,command=iexit).place(x=690,y=450)

totalLabel=Label(root,font=('times new roman',18,'bold'),bg="#1c86ee",fg='black')
totalLabel.place(x=10,y=450)

sentLabel=Label(root,font=('times new roman',18,'bold'),bg="#1c86ee",fg='black')
sentLabel.place(x=100,y=450)

leftLabel=Label(root,font=('times new roman',18,'bold'),bg="#1c86ee",fg='black')
leftLabel.place(x=190,y=450)

failedLabel=Label(root,font=('times new roman',18,'bold'),bg="#1c86ee",fg='black')
failedLabel.place(x=280,y=450)

root.mainloop()


# In[ ]:





# In[ ]:





# In[ ]:





# In[ ]:




