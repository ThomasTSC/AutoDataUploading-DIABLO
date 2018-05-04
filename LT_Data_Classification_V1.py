# -*- coding: utf-8 -*-
"""
Created on Thu Jul 13 11:53:00 2017

@author: chou
"""

import os
import re
import csv
import shutil

def Data_Classification():
    
    os.chdir(r'P:\Forschungs-Projekte\OLED-measurements Bruchsal\LT_FolderInfo')
    
    Lesker_folders = os.listdir("P:\Forschungs-Projekte\OLED-measurements Bruchsal\B 2017 Lesker")
    VG_folders = os.listdir("P:\Forschungs-Projekte\OLED-measurements Bruchsal\B 2017 VG2")
    Lesker_Batch_Number = []
    Lesker_Substrate_I = []
    Lesker_Substrate_F = []
    VG_Batch_Number = []
    VG_Substrate_I = []
    VG_Substrate_F = []


    for L_element in Lesker_folders:
        regex = re.compile(r'\d+')
        regex.findall(L_element)
        Output_filename=[int(x) for x in regex.findall(L_element)]
        if len(Output_filename) > 1:
            Lesker_Batch_Number.append(Output_filename[3])
            Lesker_Substrate_I.append(Output_filename[1])
            Lesker_Substrate_F.append(Output_filename[2])
    
    for V_element in VG_folders:
        regex = re.compile(r'\d+')
        regex.findall(V_element)
        Output_filename=[int(x) for x in regex.findall(V_element)]
        if len(Output_filename) > 1:
            VG_Batch_Number.append(Output_filename[3])
            VG_Substrate_I.append(Output_filename[1])
            VG_Substrate_F.append(Output_filename[2])

    with open("Lesker_Batch_Number.csv", "w",newline='') as Lesker_Batch_Number_file:
        Write_Lesker_Batch_Number = csv.writer(Lesker_Batch_Number_file, quoting=csv.QUOTE_ALL)
        Write_Lesker_Batch_Number.writerow(Lesker_Batch_Number)

    with open("Lesker_Substrate_I.csv", "w",newline='') as Lesker_Substrate_I_file:
        Write_Lesker_Substrate_I = csv.writer(Lesker_Substrate_I_file, quoting=csv.QUOTE_ALL)
        Write_Lesker_Substrate_I.writerow(Lesker_Substrate_I)
    
    with open("Lesker_Substrate_F.csv", "w",newline='') as Lesker_Substrate_F_file:
        Write_Lesker_Substrate_F = csv.writer(Lesker_Substrate_F_file, quoting=csv.QUOTE_ALL)
        Write_Lesker_Substrate_F.writerow(Lesker_Substrate_F)


    with open("VG_Batch_Number.csv", "w",newline='') as VG_Batch_Number_file:
        Write_VG_Batch_Number = csv.writer(VG_Batch_Number_file, quoting=csv.QUOTE_ALL)
        Write_VG_Batch_Number.writerow(VG_Batch_Number)

    with open("VG_Substrate_I.csv", "w",newline='') as VG_Substrate_I_file:
        Write_VG_Substrate_I = csv.writer(VG_Substrate_I_file, quoting=csv.QUOTE_ALL)
        Write_VG_Substrate_I.writerow(VG_Substrate_I)
    
    with open("VG_Substrate_F.csv", "w",newline='') as VG_Substrate_F_file:
        Write_VG_Substrate_F = csv.writer(VG_Substrate_F_file, quoting=csv.QUOTE_ALL)
        Write_VG_Substrate_F.writerow(VG_Substrate_F)
    
    Lesker_Batch_Number_list = []
    Lesker_Substrate_I_list = []
    Lesker_Substrate_F_list = []
    VG_Batch_Number_list = []
    VG_Substrate_I_list = []
    VG_Substrate_F_list = []

    Lesker_Batch_Number = open('Lesker_Batch_Number.csv', 'r')
    Lesker_Index = 0
    for line in csv.reader(Lesker_Batch_Number):
        for Lesker_Batch_element in line:
            Lesker_Batch_Number_list.append(Lesker_Batch_element)
            Lesker_Index = Lesker_Index + 1
    Lesker_Batch_Number.close()
    
    Lesker_Substrate_I = open('Lesker_Substrate_I.csv', 'r')
    for line in csv.reader(Lesker_Substrate_I):
        for Lesker_Substrate_I_element in line:
            Lesker_Substrate_I_list.append(Lesker_Substrate_I_element)
    Lesker_Substrate_I.close()

    Lesker_Substrate_F  = open('Lesker_Substrate_F.csv', 'r')
    for line in csv.reader(Lesker_Substrate_F):
        for Lesker_Substrate_F_element in line:
            Lesker_Substrate_F_list.append(Lesker_Substrate_F_element)
    Lesker_Substrate_F.close()

    VG_Batch_Number = open('VG_Batch_Number.csv', 'r')
    VG_Index = 0
    for line in csv.reader(VG_Batch_Number):
        for VG_Batch_element  in line:
            VG_Batch_Number_list.append(VG_Batch_element)
            VG_Index = VG_Index  + 1
    VG_Batch_Number .close()

    VG_Substrate_I = open('VG_Substrate_I.csv', 'r')
    for line in csv.reader(VG_Substrate_I):
        for VG_Substrate_I_element in line:
            VG_Substrate_I_list.append(VG_Substrate_I_element )
    VG_Substrate_I.close()

    VG_Substrate_F  = open('VG_Substrate_F.csv', 'r')
    for line in csv.reader(VG_Substrate_F):
        for VG_Substrate_F_element in line:
            VG_Substrate_F_list.append(VG_Substrate_F_element)
    VG_Substrate_F .close()    
    
    
    Criteria = 2000  #A criteria for distinguish the substrate from Lesker and VG#
    
    os.chdir(r'P:\Forschungs-Projekte\OLED-measurements Bruchsal\LT_Data')
    LT_Data_Folder = r'P:\Forschungs-Projekte\OLED-measurements Bruchsal\LT_Data' #Set path#
    
    Sorted_folders = sorted(os.listdir(LT_Data_Folder) , key=lambda x: int(x.split('_')[2]))


    
    for element in Sorted_folders:
        print (element)
        regex = re.compile(r'\d+')
        regex.findall(element)
        """Get Info of subtrate"""
        Output_filename=[int(x) for x in regex.findall(element)]
        
        
        for Lesker_element in Lesker_folders:  
            regex = re.compile(r'\d+')
            regex.findall(Lesker_element)
            Lesker_Output_filename=[int(x) for x in regex.findall(Lesker_element)]
            
            if Output_filename[1] < Criteria and Lesker_Output_filename != []:
                if Output_filename[1] >= Lesker_Output_filename[1] and Output_filename[1] <= Lesker_Output_filename[2]:
                            target = r'P:\Forschungs-Projekte\OLED-measurements Bruchsal\B 2017 Lesker\%s' %(str(Lesker_element)+'\LT')
                            shutil.move(os.path.join(LT_Data_Folder, str(element)),target)
        
        for VG_element in VG_folders:   
            regex = re.compile(r'\d+')
            regex.findall(VG_element)
            VG_Output_filename=[int(x) for x in regex.findall(VG_element)]
            if Output_filename[1] > Criteria and VG_Output_filename != [] :
                if Output_filename[1] >= VG_Output_filename[1] and Output_filename[1] <= VG_Output_filename[2] :
                            target = r'P:\Forschungs-Projekte\OLED-measurements Bruchsal\B 2017 VG2\%s' %(str(VG_element) +'\LT')
                            shutil.move(os.path.join(LT_Data_Folder, str(element)), target)
            
            
Data_Classification()            
            
            
            
            
            
            
            
            
            
            
            
            
            
            