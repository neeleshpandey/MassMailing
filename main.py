from tkinter import *
from tkinter import ttk
from tkinter.filedialog import askopenfile, askopenfiles
import smtplib
import csv
from email.message import EmailMessage
from bs4 import BeautifulSoup


attachments = []

def sendemail():    #This Function Sends the Mail
    try:
        loginMail = entryLoginMail.get()
        loginPass = entryLoginPass.get()
        subject = subjectEntry.get()
        content = messageBody.get('1.0','end')
        contentType = 'plain'
        if bool(BeautifulSoup(content, "html.parser").find()):
            contentType = 'html'
        
        j = i = 0
        report = open('./report.txt',"w")
        report.write("-"*10+"\tMail\t"+"-"*10+"\n\n")
        report.write("From:\t"+loginMail+"\n\n"+"Subject:\t"+subject+'\n\n'+"Message :\n"+content+"\n\n")


        for file in attachments:
            report.write("Attachment:\t"+file+"\n\n")

        report.write("-"*10+"\tMail\t"+"-"*10+"\n\n")
        report.write("-"*10+"\tSent Report\t"+"-"*10+"\n\n")

        with open(addressBook, 'r') as csvfile:
            csvreader = csv.reader(csvfile)
            for address in csvreader:
                try:
                    email = EmailMessage()
                    email['Subject'] = subject
                    email['From'] = loginMail
                    email['To'] = address
                    # email.set_content(content)
                    email.add_alternative(content,subtype=contentType)

                    for file in attachments:
                        with open(file,'rb') as fh:
                            file_data = fh.read()
                            file_name = fh.name.split('/')[-1]

                        email.add_attachment(file_data, maintype='application', subtype='octet-stream',filename=file_name)

                    with smtplib.SMTP_SSL('smtp.gmail.com',465) as smtp:

                        smtp.login(loginMail,loginPass)

                        smtp.send_message(email)
                    i += 1
                    report.write(address[0]+"   Sent Successfully\n")
                except Exception as e:
                    j += 1
                    report.write(address[0]+"   Unsuccessful\n")
        Label(mainframe, text=f"Yup!,{i} Email Sent Successfully,{j} Unsuccessful").grid(column=4,row=10,sticky=W)
        report.close()

    except Exception as e:
        Label(mainframe, text=str(e)).grid(column=4,row=9,sticky=W)
        report("-"*10+"\tError in Complete,in Complete Data\t"+"-"*10)
        report.close()

    
    
def selectCSV():    #Gets the position of CSV file
    file = askopenfile(mode ='r', filetypes =[('CSV Files', '*.csv')])
    global addressBook
    addressBook = file.name

def selectAttachment(): #Gets location of all Attachments to send in Mail
    global attachments
    files = askopenfiles(mode ='r')
    selectedFiles = ''
    attachments = []
    for file in files:
        attachments.append(file.name)
        selectedFiles+=file.name+'   '
    Label(mainframe,text=selectedFiles).grid(column=4, row=6, sticky=(W, E))


        
root = Tk() #GUI to Send Mail
root.title('Mailing System')
root.resizable(0,0)

mainframe = ttk.Frame(root, padding="3 3 3 3")
mainframe.grid(column=0, row=0, sticky=(N, W, E, S))
mainframe.columnconfigure(0, weight=1)
mainframe.rowconfigure(0, weight=1)

loginMail = StringVar()
loginPass = StringVar()
subject = StringVar()
msgbody = StringVar()

Label(mainframe, text="Sender's Email Address: ").grid(column=0, row=1, sticky=W)
entryLoginMail = Entry(mainframe, width=50, textvariable=loginMail)
entryLoginMail.grid(column=4, row=1, sticky=(W, E))

Label(mainframe, text="Password : ").grid(column=0, row=2, sticky=W)
entryLoginPass = Entry(mainframe, width=50, textvariable=loginPass)
entryLoginPass.grid(column=4, row=2, sticky=(W, E))

Label(mainframe, text="Recepient's Email Address: ").grid(column=0, row=3, sticky=W)
Button(mainframe,bg="brown",fg = "white",text = "Select a CSV file",command=selectCSV,relief=GROOVE).grid(column=4, row=3, sticky=(W, E))

Label(mainframe, text="Subject: ").grid(column=0, row=4, sticky=W)
subjectEntry = Entry(mainframe, width=50, textvariable=subject)
subjectEntry.grid(column=4, row=4, sticky=(W, E))

Label(mainframe, text="Message Body: ").grid(column=0, row=5, sticky=W)
messageBody = Text(mainframe, width=50, height=10)
messageBody.grid(column=4, row=5, sticky=(W, E))

Button(mainframe,bg="green",fg = "white",text = "Add Attachment",command=selectAttachment,relief=GROOVE).grid(column=0, row=6, sticky=(W, E))

Button(mainframe, bg="grey", text="Send Email", command=sendemail).grid(column=4,row=7,sticky=E)

for child in mainframe.winfo_children(): child.grid_configure(padx=5, pady=5)

root.mainloop()
