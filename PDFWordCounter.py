#!/usr/bin/env python3

from tkinter import *
from tkinter.filedialog import askopenfilename
import PyPDF2
import time
from collections import Counter

class PDFWordCounter:
    def __init__(self,root):
        self.root = root

        self.root.withdraw() # hides Tk() window
        self.openPDFFile = askopenfilename(title = "Select PDF file",filetypes = (("PDF files","*.pdf"),("all files","*.*")))
        self.data = PyPDF2.PdfFileReader(open(self.openPDFFile, 'rb'))
        self.root.deiconify() # shows Tk() window

        self.root.title('PDF Word Counter')
        self.Lab = Label(root,text='Total page number of selected PDF file:').grid(row=0,sticky=W)
        self.Lab = Label(root,text='\nInput PDF pages range:').grid(row=1,sticky=E)
        self.Lab = Label(root,text='From page : ').grid(row=2,sticky=E)
        self.Lab = Label(root,text='To page : ').grid(row=3,sticky=E)
        self.Lab = Label(root,text='Error output').grid(row=4,sticky=E)
        self.showNumPage=Text(root,height=1,width=23)
        self.showNumPage.grid(row=0,column=1)
        self.from_page = Entry(root)
        self.to_page = Entry(root)
        self.T=Text(root,height=5,width=23)

        self.from_page.grid(row=2,column=1)
        self.to_page.grid(row=3,column=1)
        self.T.grid(row=4,column=1)
        self.from_page.focus_set()

        self.numPage=self.data.numPages
        self.showNumPage.insert(END,self.numPage)

        self.Var1=IntVar()
        self.ChkBtn=Checkbutton(root,text="Include PDF text", variable=self.Var1)
        self.ChkBtn.grid(row=7,column=0)

        self.root.grid_rowconfigure(8,minsize=30)
        self.button_choose = Button(root,text='Choose another PDF file...', command = self.choose_another_file)
        self.button_choose.grid(row=8, sticky=W+E, padx=(10,10), pady=(10,0))

        self.root.grid_rowconfigure(9,minsize=50)
        self.button_run = Button(root,text='RUN COUNTER', command = self.printtext)
        self.button_run.grid(row=9,column=0,sticky=W+E, padx=(10,10))
        self.button_about = Button(root,text='About', command = self.about)
        self.button_about.grid(row=9,column=1,sticky=W+E, padx=(10,10))
        self.button_exit = Button(root,text='Exit', fg='red', command = self.destroy_root)
        self.button_exit.grid(row=9,column=2,padx=(10,10))
        self.root.mainloop()

    def printtext(self):
        try:
                    
            self.sumPDFpages=''
            self.x=int(self.from_page.get())
            self.y=int(self.to_page.get())
            self.ChkValue=self.Var1.get()

            if (self.x < 1) or (self.x > self.y):
                    raise IndexError
                    
            for i in range(self.x-1,self.y):
                    self.PDFpage=self.data.getPage(i)
                    self.PDFpage=self.PDFpage.extractText()
                    self.sumPDFpages=self.sumPDFpages+self.PDFpage
               
                
            self.output = Toplevel(self.root)
            self.output.title('PDF Word Counter')
            self.scrollbar = Scrollbar(self.output)
            self.scrollbar.grid(row=1, column=2, sticky=N+S+W)
            self.T1 = Text(self.output,height=40,width=80)
            self.T1.grid(row=1,column=1)
            self.T1.config(yscrollcommand=self.scrollbar.set)
            self.scrollbar.config(command=self.T1.yview)
            
            self.closeBttn=Button(self.output,text='Close', command=self.destroy_output)
            self.closeBttn.grid(row=2,column=1)
            
            if self.ChkValue == True:
                self.T1.insert(END,'This is text from PDF file:\n\n' + self.sumPDFpages + '\n')
                
            self.PDFpageList = self.sumPDFpages.split()
            self.time1=time.time()
            self.wordCount = Counter(self.PDFpageList).most_common(10)
            self.wordCount=str(self.wordCount)
            self.time2=time.time()
            self.T1.insert(END,'\n\nThis is a list of 10 most common words in choosen PDF file:\n\n'+self.wordCount)
            self.time_exec=str(self.time2-self.time1)
            self.T1.insert(END,'\n\nExecution Counter time:\n'+self.time_exec+' sec')
            self.output.focus_set()                                                        
            self.output.grab_set()

                
        except ValueError:
            self.T.insert(END,'Input integer values\n')

        except IndexError:
            self.T.insert(END,'Pages out of range\n')

    def choose_another_file(self):
        self.openPDFFile = askopenfilename(title = "Select PDF file",filetypes = (("PDF files","*.pdf"),("all files","*.*")))
        self.data = PyPDF2.PdfFileReader(open(self.openPDFFile, 'rb'))
        self.numPage=self.data.numPages
        self.showNumPage.delete('1.0',END)
        self.from_page.delete(0, END)
        self.to_page.delete(0, END)
        self.from_page.focus_set()
        self.showNumPage.insert(END,self.numPage)
        

    def about(self):
        self.about = Toplevel(self.root)
        self.about.title('About')
        self.lab_name = Label(self.about,text='Author: Aleksandar Kurjakov', font='Ariel 12 bold').grid(row=1,column=1,padx=20,pady=15)
        self.lab_contact = Label(self.about,text='Contact: kurjak021@gmail.com', font='Ariel 12 bold').grid(row=2,column=1,padx=20,pady=15)
        self.lab_version = Label(self.about,text='PDF WordCounter V1.0', font='Ariel 10 bold').grid(row=3,column=1,padx=20,pady=15)
        self.about.focus_set()                                                        
        self.about.grab_set()

    def destroy_root(self):
        self.root.destroy()

    def destroy_output(self):
        self.output.destroy()
    
root = Tk()
App = PDFWordCounter(root)
