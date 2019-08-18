from tkinter import Label, Button, Entry
import tkinter as tk
from tkinter import messagebox, filedialog, PhotoImage, StringVar, SUNKEN, W, X, BOTTOM
from os import path
import TestResultAnalysis
import threading, sys

class GUI_COntroller:
    '''
	   This class initialize the required controls for TkInter GUI
	'''
    def __init__(self,TkObject):
 
 
	    #Load company image
        Imageloc=tk.PhotoImage(file='alstom_logo.gif')		
        label3=Label(image=Imageloc,)
        label3.image = Imageloc		
        label3.place(x=200,y=10)
		
	
	    #SAEntry = Entry(root,takefocus=False,justify=tk.CENTER,font=50,)

        global TkObject_ref, validResDirSelected, selectedDir_seq, validSeqDirSelected
        TkObject_ref =  TkObject
        validSeqDirSelected = False		
        validResDirSelected = False		

		
        #label
        global label1		
        label1 = Label(TkObject,bd=7, text="Test tracking sheet preparation for selected build", bg="green", fg="black",width=60, font=200)	
        label1.place(x=100,y=80)
        label1.config(font=('helvetica',12,'bold'))		

        global label_select_seq_loc		
        label_select_seq_loc = Label(TkObject,bd=7, text="1. Select Test sequence files directory :: ", bg="orange", fg="black",width=30, font=200)	
        label_select_seq_loc.place(x=30,y=140)
        label_select_seq_loc.config(font=('helvetica',12,'bold'))			
		
        #select sequence files directory
        global 	button1_select_seq_loc	
        button1_select_seq_loc=Button(TkObject,activebackground='green',borderwidth=5, text='Click here to select path',width=25, command=GUI_COntroller.selectSeqDirectory)
        button1_select_seq_loc.place(x=430,y=140)
        button1_select_seq_loc.config(font=('helvetica',12,'bold'))

        global label_select_result_loc		
        label_select_result_loc = Label(TkObject,bd=7, text="2. Select Test results files directory :: ", bg="orange", fg="black",width=30, font=200)	
        label_select_result_loc.place(x=30,y=210)
        label_select_result_loc.config(font=('helvetica',12,'bold'))			
		
        #select sequence files directory
        global 	button1_select_result_loc	
        button1_select_result_loc=Button(TkObject,activebackground='green',borderwidth=5, text='Click here to select path',width=25, command=GUI_COntroller.selectResDirectory)
        button1_select_result_loc.place(x=430,y=210)
        button1_select_result_loc.config(font=('helvetica',12,'bold'))		

        global label_buildVersion		
        label_buildVersion = Label(TkObject,bd=7, text="3. What is software build version :: ", bg="orange", fg="black",width=30, font=200)	
        label_buildVersion.place(x=30,y=290)
        label_buildVersion.config(font=('helvetica',12,'bold'))
		
        global EntryObj_buildVersion
        EntryObj_buildVersion = Entry(TkObject,font=20,bd=2)
        EntryObj_buildVersion.place(x=430,y=300)		

        global label_labelVersion		
        label_labelVersion = Label(TkObject,bd=7, text="4. What is test labelled version :: ", bg="orange", fg="black",width=30, font=200)	
        label_labelVersion.place(x=30,y=350)
        label_labelVersion.config(font=('helvetica',12,'bold'))

		
        global EntryObj_labelVersion
        EntryObj_labelVersion = Entry(TkObject,font=20,bd=2)
        EntryObj_labelVersion.place(x=430,y=360)
		
        #Exit Window
        global button2_close		
        button2_close=Button(TkObject,activebackground='green',borderwidth=5, text='Close Window', command=GUI_COntroller.exitWindow)
        button2_close.place(x=600,y=430)	
        button2_close.config(font=('helvetica',12,'bold'))	

				
        #select sequence files directory
        global 	button1_executeTest	
        button1_executeTest=Button(TkObject,activebackground='green',borderwidth=5, text='Run Test Tracking Analyse',width=25, command=GUI_COntroller.RunTest)
        button1_executeTest.place(x=230,y=430)
        button1_executeTest.config(font=('helvetica',12,'bold'))	

    def exitWindow():
            TkObject_ref.destroy()

    def RunTest():

        runTest = True
        global validSeqDirSelected, validResDirSelected, EntryObj_buildVersion, EntryObj_labelVersion
	
        if validSeqDirSelected == False:
            messagebox.showerror('Error','Please select a test sequence directory!')
            runTest = False
            TkObject_ref.destroy()
            sys.exit()
	
        if validResDirSelected == False:
            messagebox.showerror('Error','Please select a test results directory!')
            runTest = False			
            TkObject_ref.destroy()
            sys.exit()			
			
        if len(EntryObj_buildVersion.get()) == 0:
            messagebox.showerror('Error','Please enter software build version!')
            runTest = False			
            TkObject_ref.destroy()		
            sys.exit()			
			
        if len(EntryObj_labelVersion.get()) == 0:
            messagebox.showerror('Error','Please enter test label build version!')
            runTest = False			
            TkObject_ref.destroy()
            sys.exit()			
			
        if runTest:
            TestTracking.RunTest()
			
    def selectSeqDirectory():
            global selectedDir_seq, validSeqDirSelected
            selectedDir_seq = filedialog.askdirectory(initialdir = "/")
            if not path.isdir(selectedDir_seq):
                messagebox.showerror('Error','Please select a valid directory!')				
            else:
                button1_select_seq_loc.destroy()

                label4= Label(TkObject_ref,bg='white',text=str(selectedDir_seq),font=40)
                label4.place(x=430,y=140)
                validSeqDirSelected = True

    def selectResDirectory():
            global selectedDir_res, validResDirSelected
            selectedDir_res = filedialog.askdirectory(initialdir = "/")
            if not path.isdir(selectedDir_res):
                messagebox.showerror('Error','Please select a valid directory!')				
            else:
                button1_select_result_loc.destroy()

                label5= Label(TkObject_ref,bg='white',text=str(selectedDir_res),font=40)
                label5.place(x=430,y=210)
                validResDirSelected = True				
	
class TestTracking:
    def RunTest(): 

        global thread,statusBarText, button1_executeTest

        button1_executeTest.config(state="disabled")
		
        statusBarText = StringVar()		
        StatusLabel = Label(TkObject_ref, textvariable=statusBarText, fg="green", bd=1,relief=SUNKEN,anchor=W) 
        StatusLabel.config(font=('helvetica',11,'bold'))
        StatusLabel.pack(side=BOTTOM, fill=X)
        statusBarText.set("Sequence analyis in progress...")
		
        thread = threading.Thread(target=TestResultAnalysis.script_exe, args = (selectedDir_seq,selectedDir_res,EntryObj_buildVersion.get(), EntryObj_labelVersion.get(), TkObject_ref,statusBarText))
        thread.start()

if __name__ == '__main__':	
	
       root = tk.Tk()
       
       #Change the background window color
       root.configure(background='gray')     
       
       #Set window parameters
       root.geometry('850x600')
       root.title('Test tracking sheet')
       
       #Removes the maximizing option
       root.resizable(0,0)
       
       ObjController = GUI_COntroller(root)
       
       #keep the main window is running
       root.mainloop()
       sys.exit()
