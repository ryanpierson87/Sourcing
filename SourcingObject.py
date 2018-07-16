# -*- coding: utf-8 -*-
"""
Created on Thu Jul 12 12:06:24 2018

@author: 502755426
"""

#The backup
import zipfile as zp
import datetime as dt
from tkinter import *
from tkinter.filedialog import askopenfilename
import pandas as pd
import os
import shutil as shtl
from tkinter import messagebox

class Program:
###
######## Creates the interface|| links the reference .csv file|| populates data based on the .csv file
###
    def __init__(self, master):
            path = os.getcwd()
            fair = path.split("\\")
            for i in fair:
                if i.isdigit():
                    link = i
            os.chdir('C:/Users/'+ link)
            self.data = pd.read_csv("SourcingPython\\ProgramData.csv")
            self.data.Input[0] = ""
                        

            self.input = Button(text="input",bg="brown", fg="white", relief=RIDGE,width=10, command=self.inputSelect).grid(row=1,column=0)
            self.inLabel = Label(text=self.data.Input[0], relief=SUNKEN,width=50).grid(row=1,column=1)
            self.output = Button(text="output",bg="brown", fg="white", relief=RIDGE,width=10, command=self.outputSelect).grid(row=2,column=0)
            self.outLabel = Label(text=self.data.Output[0], relief=SUNKEN,width=50).grid(row=2,column=1)
            self.zipFile = Button(text="Start Automation", fg="green", relief=RIDGE, width=40, command=self.zippo).grid(row=5, column=0, columnspan=2)

###############################

###
#####       methods to select the input zip file and output directory
###
    def inputSelect(self):
        print(self.data["Input"][0])
        self.data.Input.at[0]  = askopenfilename()
        print(self.data.Input[0])
        self.data.to_csv("SourcingPython\\ProgramData.csv", index=False)
        self.inLabel = Label(text=self.data.Input[0][self.data.Input[0].rfind("/")+ 1:], relief=SUNKEN,width=50).grid(row=1,column=1) 
        root.update()
    
    def outputSelect(self):
        print(self.data["Output"][0])
        self.data.Output.at[0]  = filedialog.askdirectory()
        print(self.data.Output[0])
        self.data.to_csv("SourcingPython\\ProgramData.csv", index=False)
        self.outLabel = Label(text=self.data.Output[0], relief=SUNKEN,width=50).grid(row=2,column=1) 
        root.update()
#################################

        
    def zippo(self):
#####
##  Extracts a file into a specific location and deleting original 
#####This portion works within the object correctly
####
        if len(self.data.Input[0]) > 3:
            print("success")
            now = dt.datetime.now()
            current ="Sourcing-" + str(now.month) + "_" +str( now.day) + "-"+str(now.year)
            zip_file = self.data.Input[0]
            self.zip_dest = self.data.Input[0][:self.data.Input[0].rfind("/")+ 1] + current
            print("success2")
            file = zp.ZipFile(zip_file)
            os.mkdir(self.zip_dest)
            file.extractall(path=self.zip_dest)
            print("success3")
            file.close()
            os.remove(zip_file)
            self.combine(self.zip_dest)
            messagebox.showinfo("Complete", "Sourcing Templates are available")
        
        #The dictionary to divvy up the Buyer Review documents
    def combine(self, zip_dest):
        self.reports_to_collect = {}
        
        ### Loops through the files in the directory of the extracted zip file
        for roots, dirs, files in os.walk(zip_dest):
            for file in files:
                report = file
                initial = report.find(" ") + 1
                if initial > 0:
                    try:
                        org = report[initial: report.find("-", initial)]
                    except:
                        pass
                    if org in self.reports_to_collect:
                        if report not in self.reports_to_collect:
                            self.reports_to_collect[org].append(file)
                    else:
                        self.reports_to_collect[org] = []
                        self.reports_to_collect[org].append(file)

#Moves the data from the files to duplicates of the empty tempalte documents
        self.columns = ['ORDER UNIT OF MEASURE', 'Order UOM Price','Supplier', \
           'Supplier Site\n(POI preferred)', 'Supplier Item Number','New/Existing Part Number (entered by Loading Team/Code)', 'Ship To']

        self.template = pd.read_excel("SourcingPython\SourcingTemplate.xlsx")
        self.template.dropna(inplace=True)
        final = pd.DataFrame()
        #
        #Loops through the above dictionary and and creates the copy of the template for each org
        for key in self.reports_to_collect:
            template1 = self.template.copy()
            for i in self.reports_to_collect[key]:
                df = pd.read_excel(self.zip_dest+"\\" + i, "New Item Entry")     
                df.drop_duplicates(keep=False, inplace= True)   
                final = pd.concat([final, df])
            #Moves copies the data for each sheet per org into the created template duplicate
            ##
            #CODE FOR REST OF THE PROGRAM WITHIN THIS LOOP
            ##
            for j in range(len(self.columns)): 
                template1[self.columns[j]] = final[self.columns[j]]
            final.to_excel(key + ".xlsx", index=False)
            template1.to_excel(self.data.Output[0]+"//Sourcing "+key + ".xlsx", index=False)
            
            #Save the concatenation as a variable to be used in the first SQL Query
            self.sql1 = "'" + template1["New/Existing Part Number (entered by Loading Team/Code)"]+"'"
            self.sql1 =','.join(self.sql1)
            #####
            #####     NEXT STEP IN THE PROCESS
            #####
            
        shtl.rmtree(zip_dest) 
            
            
        
root = Tk()
app = Program(root)
root.mainloop()
