# -*- coding: utf-8 -*-
"""
Created on Wed Jul  5 18:38:37 2017

@author: chou
"""

import os
import tkinter as tkinter 
import time
import re
from tkinter import ttk
import shutil
import zipfile
import datetime as dt
import csv
import sqlite3
from selenium import webdriver
from selenium.webdriver.support.ui import Select


#Creat and Connect to the databas#
conn = sqlite3.connect('Life_Time_Database.db')
c = conn.cursor()
c.execute('CREATE TABLE IF NOT EXISTS Life_Time_Database(Batch REAL, Substrate REAL, Pixel REAL,Machine TEXT, Date TEXT, Measured REAL, Saved REAL, Uploaded REAL, Emailed REAL)')



class DIABLO_AutoUpload(tkinter.Frame):
    
    def __init__(self, parent):
        '''
        Constructor
        '''
        tkinter.Frame.__init__(self, parent)
        self.parent=parent
        self.DIABLO_UploadInterface()
    
    def DIABLO_UploadInterface(self):
        
        """Draw a user interface allowing the user to type
        items and insert them into the treeview
        """
        self.parent.title("DIABLO")       
        self.parent.grid_rowconfigure(0,weight=1)
        self.parent.grid_columnconfigure(0,weight=1)
        self.parent.geometry('580x420')
        self.Time = time.strftime("%d/%m/%Y")

        #Enter the number of batch#
        self.Batch_Label = tkinter.Label(self.parent, text = ' Batch # ', font = ("Arial", 12), width = 5,  height = 1).place (x = 20, y = 20)
        self.var_Batch_Number = tkinter.IntVar()
        self.Batch_Label_Entry = tkinter.Entry(self.parent, textvariable = self.var_Batch_Number, width = 10,show = None).place(x = 20, y = 50)

        #Enter the number of substrate#
        self.Substrate_Label = tkinter.Label(self.parent, text = ' Sub. # ', font = ("Arial", 12), width = 5,  height = 1).place (x = 110, y = 20)
        self.var_Substrate_Number = tkinter.IntVar()
        self.Substrate_Label_Entry = tkinter.Entry(self.parent, textvariable = self.var_Substrate_Number, width = 10,show = None).place(x = 110, y = 50)

        #A empty table for Data Input#
        self.Ininfo_tree = ttk.Treeview( columns=('Substrate','Pixel', 'Machine','Date','Mea.','Sav.','Upl.','Em'))
        self.Ininfo_tree.heading('#0', text='B.')
        self.Ininfo_tree.heading('#1', text='Sub.')
        self.Ininfo_tree.heading('#2', text='Pix.')
        self.Ininfo_tree.heading('#3', text='Mac.')
        self.Ininfo_tree.heading('#4', text='Date')
        self.Ininfo_tree.heading('#5', text='Mea.')
        self.Ininfo_tree.heading('#6', text='Sav.')
        self.Ininfo_tree.heading('#7', text='Upl.')
        self.Ininfo_tree.heading('#8', text='Em.')
    

        self.Ininfo_tree.column('#0', width = 50)
        self.Ininfo_tree.column('#1', width = 70)
        self.Ininfo_tree.column('#2', width = 50)
        self.Ininfo_tree.column('#3', width = 50)
        self.Ininfo_tree.column('#4', width = 100)
        self.Ininfo_tree.column('#5', width = 60)
        self.Ininfo_tree.column('#6', width = 60)
        self.Ininfo_tree.column('#7', width = 60)
        self.Ininfo_tree.column('#8', width = 60)

        self.Ininfo_tree.place(x = 10, y = 100)

        #Scrollbar should be added.#

        self.Intreeview = self.Ininfo_tree
        
        
        #To search data in the database#
        
        
        def Search_Data():
            
             os.chdir(r"C:\Users\Administrator.ELS-109\Desktop\LT_Data_Managemet") 
             #Clear the data of the treeview#
             for i in self.Intreeview.get_children():
                self.Intreeview.delete(i)
                
             if self.var_Batch_Number.get() != 0 and self.var_Substrate_Number.get() == 0:
                conn = sqlite3.connect('Life_Time_Database.db')
                c = conn.cursor()
                c.execute('SELECT * FROM Life_Time_Database WHERE Batch = ? ', (self.var_Batch_Number.get() ,))
                self.found = c.fetchall()
                if self.found:
                    self.i = 0
                    #Check whether the new input is different from the saved data#
                    for row in self.found:            
                        self.Intreeview.insert("" , self.i,  text= (int(row[0])), values=(int(row[1]), '%s' %(int(row[2])),"%s" %row[3],row[4],"%s"  %((row[5])),
                                       '%s' %((row[6])),'%s' %((row[7])),'%s' %((row[8]))))                        
                        self.i = self.i + 1
                        
             if self.var_Batch_Number.get()  == 0 and self.var_Substrate_Number.get() != 0:
                conn = sqlite3.connect('Life_Time_Database.db')
                c = conn.cursor()
                c.execute('SELECT * FROM Life_Time_Database WHERE Substrate = ? ', (self.var_Substrate_Number.get(),))
                self.found = c.fetchall()
                if self.found:
                    self.i = 0
                    #Check whether the new input is different from the saved data#
                    for row in self.found:
                         self.Intreeview.insert("" , self.i,  text= (int(row[0])), values=(int(row[1]), '%s' %(int(row[2])),"%s" %row[3],row[4],"%s"  %((row[5])),
                                       '%s' %((row[6])),'%s' %((row[7])),'%s' %((row[8]))))     
                         self.i = self.i + 1     
        
             if self.var_Batch_Number.get()  != 0 and self.var_Substrate_Number.get() != 0:
                conn = sqlite3.connect('Life_Time_Database.db')
                c = conn.cursor()
                c.execute('SELECT * FROM Life_Time_Database WHERE Batch=? AND Substrate = ? ', (self.var_Batch_Number.get() ,self.var_Substrate_Number.get(),))
                self.found = c.fetchall()
                if self.found:
                    self.i = 0
                    #Check whether the new input is different from the saved data#
                    for row in self.found:
                         self.Intreeview.insert("" , self.i,  text= (int(row[0])), values=(int(row[1]), '%s' %(int(row[2])),"%s" %row[3],row[4],"%s"  %((row[5])),
                                       '%s' %((row[6])),'%s' %((row[7])),'%s' %((row[8]))))     
                         self.i = self.i + 1     
                 
        
        self.Search_Button = tkinter.Button(self.parent, text = " Search ", width = 10, command = Search_Data).place(x = 180, y = 30)
        
        
        
        
        
                #Clear the data of the treeview and all input values on the interface#
        def Clear_Data():
            #Clear the data of the treeview#
            for i in self.Intreeview.get_children():
                self.Intreeview.delete(i)
                
            #Clear the values on the interface#    
            self.var_Batch_Number.set(0)
            self.var_Substrate_Number.set(0)
            
        
        self.Clear_Button = tkinter.Button(self.parent, text = " Clear ", width = 10, command = Clear_Data)
        self.Clear_Button.place(x = 180, y = 55)
        


        def VG_Data_Page():
            os.chdir(r"C:\Users\Administrator.ELS-109\Desktop\LT_Data_Managemet") 
            #Clear the window first#
            for i in self.Intreeview.get_children():
                self.Intreeview.delete(i)
    
            #Counter#
            self.i = 0
            
            conn = sqlite3.connect('Life_Time_Database.db')
            c = conn.cursor()
            VG_Data = c.execute("SELECT * FROM Life_Time_Database WHERE Machine = 'VG2' ORDER BY Batch")
            #Fetch one by one with for-loop
            VG_Data = c.fetchall()
            for row in VG_Data:
                #print row
                self.Intreeview.insert("" , self.i,  text= (int(row[0])), values=(int(row[1]), '%s' %(int(row[2])),"%s" %row[3],row[4],"%s"  %((row[5])),
                                       '%s' %((row[6])),'%s' %((row[7])),'%s' %((row[8]))))
                self.i = self.i + 1
            c.close()
            conn.close()
      
        def Lesker_Data_Page():
            os.chdir(r"C:\Users\Administrator.ELS-109\Desktop\LT_Data_Managemet") 
            #Clear the window first#
                        #Clear the window first#
            for i in self.Intreeview.get_children():
                self.Intreeview.delete(i)
            #Counter#
            self.i = 0
    
            conn = sqlite3.connect('Life_Time_Database.db')
            c = conn.cursor()
            Lesker_Data = c.execute("SELECT * FROM Life_Time_Database WHERE Machine = 'Lesker' ORDER BY Batch")
            #Fetch one by one with for-loop
            Lesker_Data = c.fetchall()
            for row in Lesker_Data:
                #print row
                self.Intreeview.insert("" , self.i,  text= (int(row[0])), values=(int(row[1]), '%s' %(int(row[2])),"%s" %row[3],row[4],"%s"  %((row[5])),
                                       '%s' %((row[6])),'%s' %((row[7])),'%s' %((row[8]))))               
                self.i = self.i + 1
            c.close()
            conn.close()
    
    
        self.VG_Data_Button = tkinter.Button(self.parent, text = " VG ", width = 15, command = VG_Data_Page).place(x = 300, y = 40)
        self.Lesker_Data_Button = tkinter.Button(self.parent, text = " Lesker ", width = 15, command = Lesker_Data_Page).place(x = 450, y = 40)
    

        def Summon():
        
            """Locate the folder"""
            os.chdir(r"C:\Users\Administrator.ELS-109\Desktop\LT Results") 
            self.LT_Result_Folder = r"C:\Users\Administrator.ELS-109\Desktop\LT Results" #Set path#
    
            """Export data"""
            for root,dirs,files in os.walk(r'C:\Users\Administrator.ELS-109\Desktop\LT Results'):  #Export all the data saved in past 12 hours#
                for dirs_name in dirs:
                    self.now = dt.datetime.now()
                    self.before = self.now - dt.timedelta(hours=0)
                    try:
                        self.st = os.stat(dirs_name)
                    except WindowsError:
                        pass
                    self.mod_time = dt.datetime.fromtimestamp(self.st.st_ctime)
                    if  self.before < self.mod_time:
                        shutil.copyfile(os.path.join(root, dirs_name), r'C:\Users\Administrator.ELS-109\Desktop\LT_Data_Managemet\DataManagement')
                        
            """Classify, compress and upload"""
            os.chdir(r'C:\Users\Administrator.ELS-109\Desktop\LT_Data_Managemet\DataManagement') #Move the data management folder#
            self.About_Upload_Folder = os.listdir(r'C:\Users\Administrator.ELS-109\Desktop\LT_Data_Managemet\DataManagement')
            self.Management_Folder = r'C:\Users\Administrator.ELS-109\Desktop\LT_Data_Managemet\DataManagement'
            self.Sorted_folders = sorted(self.About_Upload_Folder, key=lambda x: int(x.split('_')[2]))
            
            """Email info"""
            self.Email_VG = []
            self.Email_Batch_Number_VG = []
            self.Email_Lesker = []
            self.Email_Batch_Number_Lesker = []
            self.i=0
            
            
            for element in self.Sorted_folders:
                
                self.regex = re.compile(r'\d+')
                self.regex.findall(element)
                """Get Info of subtrate"""
                self.Output_filename=[int(x) for x in self.regex.findall(element)]
                
              
                """VG or Lesker"""
                os.chdir(r'C:\Users\Administrator.ELS-109\Desktop\LT_Data_Managemet')
                self.Criteria = 2000  #A criteria for distinguish the substrate from Lesker and VG#
                
                self.Lesker_Batch_Number_list = []
                self.Lesker_Substrate_I_list = []
                self.Lesker_Substrate_F_list = []
                self.VG_Batch_Number_list = []
                self.VG_Substrate_I_list = []
                self.VG_Substrate_F_list = []

                self.Lesker_Batch_Number = open('Lesker_Batch_Number.csv', 'r')
                self.Lesker_Index = 0
                for line in csv.reader(self.Lesker_Batch_Number):
                    for self.Lesker_Batch_element in line:
                       self. Lesker_Batch_Number_list.append(self.Lesker_Batch_element)
                       self.Lesker_Index = self.Lesker_Index + 1
                self.Lesker_Batch_Number .close()
                        
                self.Lesker_Substrate_I = open('Lesker_Substrate_I.csv', 'r')
                for line in csv.reader(self.Lesker_Substrate_I):
                    for self.Lesker_Substrate_I_element in line:
                        self.Lesker_Substrate_I_list.append(self.Lesker_Substrate_I_element)
                self.Lesker_Substrate_I.close()

                self.Lesker_Substrate_F  = open('Lesker_Substrate_F.csv', 'r')
                for line in csv.reader(self.Lesker_Substrate_F):
                    for self.Lesker_Substrate_F_element in line:
                        self.Lesker_Substrate_F_list.append(self.Lesker_Substrate_F_element)
                self.Lesker_Substrate_F .close()

                self.VG_Batch_Number = open('VG_Batch_Number.csv', 'r')
                self.VG_Index = 0
                for line in csv.reader(self.VG_Batch_Number):
                    for self.VG_Batch_element  in line:
                        self.VG_Batch_Number_list.append(self.VG_Batch_element)
                        self.VG_Index = self.VG_Index  + 1
                self.VG_Batch_Number .close()

                self.VG_Substrate_I = open('VG_Substrate_I.csv', 'r')
                for line in csv.reader(self.VG_Substrate_I):
                    for self.VG_Substrate_I_element in line:
                        self.VG_Substrate_I_list.append(self.VG_Substrate_I_element )
                self.VG_Substrate_I.close()

                self.VG_Substrate_F  = open('VG_Substrate_F.csv', 'r')
                for line in csv.reader(self.VG_Substrate_F):
                    for self.VG_Substrate_F_element in line:
                        self.VG_Substrate_F_list.append(self.VG_Substrate_F_element)
                self.VG_Substrate_F .close()


                for w in range (0,self.Lesker_Index):
                    if self.Output_filename[1] >= int(self.Lesker_Substrate_I_list[w]) and self.Output_filename[1] <= int(self.Lesker_Substrate_F_list[w]) :
                        self.Lesker = self.Lesker_Batch_Number_list[w]
                        self.Lesker_I = self.Lesker_Substrate_I_list[w]
                        self.Lesker_F = self.Lesker_Substrate_F_list[w]
                
                for i in range (0,self.VG_Index):
                    if self.Output_filename[1] >= int(self.VG_Substrate_I_list[i]) and self.Output_filename[1] <= int(self.VG_Substrate_F_list[i]) :
                        self.VG = self.VG_Batch_Number_list[i]
                        self.VG_I = self.VG_Substrate_I_list[i]
                        self.VG_F = self.VG_Substrate_F_list[i]
                    
                if self.Output_filename[1] < self.Criteria:
                    self.Upload_Info = ["Lesker", int(self.Lesker), int(self.Lesker_I), int(self.Lesker_F), int(self.Output_filename[1]), int(self.Output_filename[2])]
                    self.Email_Lesker.append(self.Upload_Info) 
                    
                if self.Output_filename[1] > self.Criteria:
                    self.Upload_Info = ["VG2", int(self.VG), int(self.VG_I), int(self.VG_F), int(self.Output_filename[1]), int(self.Output_filename[2])]
                    self.Email_VG.append(self.Upload_Info) 
                
                os.chdir(r'C:\Users\Administrator.ELS-109\Desktop\LT_Data_Managemet\DataManagement')
                
                """Upload info"""
                self.Machine = self.Upload_Info[0]
                self.Batch = self.Upload_Info[1]
                self.Substrate_I = self.Upload_Info[2]
                self.Substrate_F = self.Upload_Info[3]                
                
                
                #Insert data into database#
                os.chdir(r"C:\Users\Administrator.ELS-109\Desktop\LT_Data_Managemet") 
                conn = sqlite3.connect('Life_Time_Database.db')
                c = conn.cursor()
                c.execute('SELECT * FROM Life_Time_Database WHERE Batch = ? AND Substrate = ? AND Pixel = ?', (self.Batch, self.Upload_Info[4],self.Upload_Info[5],))
                self.found = c.fetchall()
                if self.found:
                    pass
                else:
                    c.execute("INSERT INTO Life_Time_Database (Batch, Substrate, Pixel , Machine, Date,  Measured, Saved, Uploaded, Emailed) VALUES (?, ?,?,?, ?, ?, ?, ?, ?)",(self.Batch, self.Upload_Info[4],  self.Upload_Info[5], self.Machine,self.Time,'Y','Y', 'Y', 'Y'))
                    conn.commit()
                        
                os.chdir(r'C:\Users\Administrator.ELS-109\Desktop\LT_Data_Managemet\DataManagement')
                
                """Compress the folder"""
                if zipfile.is_zipfile(str(element)) == False: 
                    
                    shutil.make_archive(str(element),'zip',os.path.join(self.Management_Folder,str(element)))

                    """Connect through selenium"""
                    self.driver = webdriver.Chrome()
                    self.driver.maximize_window()
                    self.driver.get('http://cynora-dd.cynoraad.lan')
                    self.driver.find_element_by_xpath('//*[@id="username"]').send_keys('chou')
                    self.driver.find_element_by_xpath('//*[@id="password"]').send_keys('I1TankOwAk')

                    self.driver.find_element_by_xpath('/html/body/div/div/div/div[2]/div/form/fieldset/div[3]/div/button').click()

                    self.driver.get('http://cynora-dd.cynoraad.lan/upload/')

                    
                    Select(self.driver.find_element_by_xpath('//*[@id="id_location"]')).select_by_value('%s'%self.Machine)
                    
                    self.driver.find_element_by_xpath('//*[@id="id_batch"]').send_keys('%s'%self.Batch)
                    
                    self.driver.find_element_by_xpath('//*[@id="id_raw_lt"]').click()
                    self.driver.find_element_by_xpath('//*[@id="id_raw_lt"]').clear()
 
                    self.driver.find_element_by_xpath('//*[@id="id_raw_lt"]').send_keys(os.path.join(self.Management_Folder,str(element)+'.zip'))
                    
                    
                    self.driver.find_element_by_xpath('/html/body/div/div/div/div[2]/form/input[2]').click()
                    
                    
                    self.driver.get('http://cynora-dd.cynoraad.lan/process/')
                    
                    Select(self.driver.find_element_by_xpath('//*[@id="id_location"]')).select_by_value('%s'%self.Machine)
                    
                    self.driver.find_element_by_xpath('//*[@id="id_OLED_min"]').send_keys('%s' %self.Substrate_I)
                    self.driver.find_element_by_xpath('//*[@id="id_OLED_max"]').send_keys('%s' %self.Substrate_F)
                    
                    self.driver.find_element_by_xpath('//*[@id="id_process_eqe"]').click()
                    
                    self.driver.find_element_by_xpath('//*[@id="id_process_fabSheet"]').click()
                    
                    self.driver.find_element_by_xpath('/html/body/div/div/div/div[2]/form/input[4]').click()
                    
                    time.sleep(2)
                    
                    self.driver.close()
                    
                    
                else:
                    pass
            
            os.chdir(r"C:\Users\Administrator.ELS-109\Desktop\LT_Data_Managemet")  
            try: 
                self.VG_Number=len(self.Email_VG)
            except AttributeError:
                pass
            
            try:
                self.Lesker_Number=len(self.Email_Lesker)
            except AttributeError:
                pass
            
            with open('Life_Time_Email.txt', 'w') as self.Email:
                self.Email.write('Hello, \n')
                self.Email.write('We have the following data uploaded already: \n')
                
                self.Email.write('VG:\n')
                self.Email.write('Batch   Substrate  Pixel \n')
                
                for i in range(0, self.VG_Number):
                    self.Email.write('%s    %s   %s \n' %(self.Email_VG[i][1], self.Email_VG[i][4],self.Email_VG[i][5]))
                    
        
                self.Email.write('Lesker:\n')
                self.Email.write('Batch   Substrate  Pixel \n')
                
                for i in range(0, self.Lesker_Number):
                    self.Email.write('%s    %s   %s \n' %(self.Email_Lesker[i][1], self.Email_Lesker[i][4],self.Email_Lesker[i][5]))
                
                
                self.Email.write('Best Regards, \n')
            
            tkinter.messagebox.showinfo(title = "Info", message = "The process is terminated.")
            
        self.Summon_DIABLO_Button = tkinter.Button(self.parent, text = " Summon! ", width = 15, height = 1, font = ("Arial", 24),command = Summon).place(x = 150, y = 330)
        
        
        
        
        
        
        
        
        
        
def main():
    root=tkinter.Tk()
    DIABLO_AutoUpload(root)
    root.mainloop()

if __name__=="__main__":
    main()