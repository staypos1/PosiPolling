__author__ = 'Eric William Rowe'
__copyright__ = "Copyright (C) 2015 Eric William Rowe"
__version__ = "1.0"
import os, time, glob, datetime,sys,mysql.connector,smtplib,tkFileDialog,fileinput,ttk
from Tkinter import *
from PIL import ImageTk
from mysql.connector import MySQLConnection,Error
from mysql.connector.constants import ClientFlag
from python_mysql_dbconfig import read_db_config
dbconfig = read_db_config()
mylist = []
now = time.strftime("%H:%M:%S")
mylist.append(now)
today = datetime.datetime.today()
wkdate = today.strftime("%U")
compName = os.getenv('USERNAME')
path_to_watch = "C:/Users/%s/Google Drive/scanpet/" %compName

connectionText = "ON"
#fill dictionary from mysql info
mI = {}
with open("options.txt") as m:
	for line in m:
		(key,val) = line.split()
		mI[str(key)] = val

toEmail = mI['toEmail']
fromEmail = mI['fromEmail']
emailPW = mI['emailPW']
msgE = 'Subject: Production availability\n\n The production availability for week %s is complete. Check the shared ProductionAvailability folder.' % wkdate
global emailSending
emailSending = 0
def toggleEmail(): 
	global emailSending
	if emailSending==1:
		emailSending=0
		logWindow.insert(END,'%s Email is disabed. \n' % mylist[0])
	else:
		emailSending=1
		logWindow.insert(END,'%s Email is enabled. \n' % mylist[0])
def aboutMe():
	textA = 'Eric William Rowe\nCopyright (C) 2015 Eric William Rowe\nversion = 1.0'
	top = Toplevel()
	top.title("About...")
	msg = Message(top,text=textA,width=400)
	msg.pack()

#GUI ROOT
root = Tk()
root.geometry("1000x500+0+0")
root.maxsize(1000,500)
root.resizable(0,0)
root.iconbitmap(r'favicon.ico')
root.configure(bg="grey",padx=5,pady=5)
root.title("MGC Availability")



#dir
dirname = tkFileDialog.askdirectory(parent=root,initialdir="/",title='Please select a directory to watch')
#main window
w = Frame(root,bg='grey')
w.pack(side="top", fill="both", expand=True,padx=5, pady=5)
#title
title = Frame(w, height=100,background="grey",relief=GROOVE,bd=2)
title.grid(row=0,column=0,columnspan=3)
title.pack(fill=X)
mgctitle = Label(title,text="MGC Availability",background="grey",fg="green",font=("Helvetica", 16))
mgctitle.pack()
#menubar
menubar = Menu(root)
# create a pulldown menu, and add it to the menu bar
filemenu = Menu(menubar, tearoff=0)
filemenu.add_command(label="Exit", command=root.quit)
menubar.add_cascade(label="File", menu=filemenu)
helpmenu = Menu(menubar, tearoff=0)
helpmenu.add_command(label="About", command=aboutMe)
menubar.add_cascade(label="Help", menu=helpmenu)
# display the menu
root.config(menu=menubar)
#logging window
logW = Frame(w,width=100, height=100,background="green",relief=GROOVE,bd=2)
logW.grid(row=1,column=0)
logW.pack(side=LEFT,fill=X, expand=True)
logTitle = Label(logW,text="Logging Window",bd=2,bg='grey')
logTitle.pack(fill=X)
logWindow = Text(logW,background="green",wrap=WORD)
logWindow.pack(fill=BOTH, expand=True)

#information window
infoC = Frame(w,width=250, height=100,relief=RIDGE,bd=2,bg='grey')
infoC.pack(side=TOP,anchor=W,fill=X)
infoTitle = Label(infoC,text="Information",relief=RAISED,bd=2,bg='grey')
infoTitle.pack(side=TOP,anchor=W,fill=X, expand=True)
connectionStatus = Label(infoC,fg="green",text="Connection is : %s" % connectionText,bg='grey',relief=RIDGE,bd=2)
connectionStatus.pack(fill=X)
currentDirL = Label(infoC,text='Watching: %s'%str(dirname),bg='grey',relief=RIDGE,bd=2)
currentDirL.pack(side=TOP,anchor=W,fill=X)
mysqlH = Label(infoC,text="Current MySQL Host : \t %s" %mI['host'],anchor=W,bg='grey')
mysqlH.grid(row=0)
mysqlH.pack(fill=X)
mysqlU = Label(infoC,text="Current MySQL User : \t %s" %mI['user'] ,anchor=W,bg='grey')
mysqlU.grid(row=1)
mysqlU.pack(fill=X)
mysqlD = Label(infoC,text="Current MySQL Database : \t %s" %mI['database'],anchor=W,bg='grey')
mysqlD.grid(row=2)
mysqlD.pack(fill=X)
#options frame
optF = Frame(w,width=250, height=100,relief=RIDGE,bd=2,bg='grey')
optF.grid(row=1,column=1)
optF.pack(side=TOP,anchor=W,fill=X)
opt = Label(optF,text="Options",relief=RAISED,bd=2,bg='grey')
opt.pack(side=TOP,anchor=W,fill=X, expand=True)
emailButton = Checkbutton(optF,text="Send Email on completion?",command=toggleEmail,bg='grey')
emailButton.pack(anchor=W)


global before
before = dict ([(f, None) for f in os.listdir (dirname)])

try:
	cnx = MySQLConnection(**dbconfig)
	cursor = cnx.cursor()
	logWindow.insert(END,'%s Connected to MySQL database. \n' % mylist[0])
except Error as e:
	logWindow.insert(END,(e))
	logWindow.insert(END,('\n'))
	connectionPicture.configure(connectionText = "OFF")

	
def sendEmail(fromE,passwordE,toE,msg):
	# The actual mail send
	server = smtplib.SMTP('smtp.gmail.com:587')
	server.starttls()
	server.login(fromE,passwordE)
	server.sendmail(fromE, toE, msg)
	server.quit()
	
def task():
	mylist=[]
	now = time.strftime("%H:%M:%S")
	mylist.append(now)
	global before
	global emailSending
	after = dict ([(f, None) for f in os.listdir (dirname)])
	added = [f for f in after if not f in before]
	removed = [f for f in before if not f in after]
	if added:
		logWindow.insert(END, '\n')
		logWindow.insert(END, '%s: Added: %s \n' % (mylist[0], str(added)))
		if "availability_" in str(added):
			if(len(added) == 1):
				if ".csv" in str(added):
					cursor.execute("""truncate table prav;""")
					cursor.execute("""LOAD DATA LOCAL INFILE 'C:/Users/%s/Google Drive/scanpet/%s'
				into table prav
				FIELDS terminated by ',' lines terminated by '\n'
				(SKU, ghp, mum, br1, br2, bl1, bl2, gleft, gright, bghp, bonsai, h1, h1_2, h2, h2_3, h3, h3_4, h4, h4_5, h5, h5_6, h6, h6_7, h7, h7_8, h8, h8_9, h9, h10, backfield);""" % (compName,''.join(added)))
					cursor.execute("""update prav
				set prav.SKU = trim(prav.SKU)
				where sku REGEXP '^[^A-Za-z]';""")
					cursor.execute("""call findmissingsku;""")
					cursor.execute("""Call AppendData;""")
					try: 
						cursor.execute("""Call AFB""")
						logWindow.insert(END,'%s: Created outfile for BILL. \n'% mylist[0])
					except:logWindow.insert(END,'%s: Unable to export for Bill. \n'% mylist[0])
					try: 
						cursor.execute("""Call LISA""")
						logWindow.insert(END,'%s: Created outfile for LISA. \n' % mylist[0])
					except:logWindow.insert(END,"%s: Unable to export for Lisa."% mylist[0])
					try: 
						cursor.execute("""Call ProdPullSheet""")
						logWindow.insert(END,'%s: Created outfile for ProdPullSheet. \n' % mylist[0])
					except:logWindow.insert(END,"%s: Unable to export for ProdPullSheet. \n" % mylist[0])
					try:
						if(emailSending):
							sendEmail(fromEmail,emailPW,toEmail,msgE)
							logWindow.insert(END,"%s: Email was dispached to %s \n" % (mylist[0],toEmail))
						else:
							logWindow.insert(END,"%s: Email was disabled to one was not sent. \n" % (mylist[0],toEmail))
					except:logWindow.insert(END,'%s: email failed. is email sending enabled?\n'% mylist[0])
					logWindow.insert(END,'%s: Production availability complete. \n'% mylist[0])
					cnx.commit()
				else:logWindow.insert(END,"%s: new file found but does not start with 'availability_'" % mylist[0])
			else:logWindow.insert(END,'%s: too many files added at once! \n' %mylist[0] )
		else:logWindow.insert(END,"%s: new file found but does not start with 'availability_'"%mylist[0])
			
	if removed: 
		logWindow.insert(END, '\n')
		logWindow.insert(END, '%s Removed: %s \n' % (mylist[0], str(removed)))
	before = after
	root.after(10000, task)
	
if __name__ == "__main__":
	try:
		logWindow.insert(END, 'Watching directory: %s\n' %dirname)
		root.after(2000, task())
		root.mainloop()
	except Error as e:
		logWindow.insert(END,'%s %s \n' %(mylist[0],e))
		print e