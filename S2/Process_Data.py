input_path = "/Users/VAadmin/Desktop/Automated_Histone_Modification/Data/Excel_Files/CSV/Data/"
output_path = "/Users/VAadmin/Desktop/Automated_Histone_Modification/Data/Excel_Files/XLSX/Processed/"


#
#
# Above is for the user to inut their desired input and output folder paths
#
#

#
# Import Statements
#

import pandas as pd
import numpy as np
import xlsxwriter
import openpyxl as pxl
from xlutils.copy import copy as xl_copy
import os
import csv
import matplotlib.pyplot as plt
from scipy.stats import f_oneway
import statsmodels.stats.multicomp as multi
from io import BytesIO
import math
from UliPlot.XLSX import auto_adjust_xlsx_column_width
from operator import itemgetter
import getpass
import platform

#
# Define Constants
#

input_file_prefix = 'Histone_Nucelar_Ratio_'
wt_ciliated_file = input_file_prefix + 'Nonexpressing_Ciliated_Cells_Markers.csv'
transfected_cilia_file = input_file_prefix + 'Expressing_Ciliated_Cells_Markers.csv'
transfected_nonciliated_file = input_file_prefix + 'Expressing_Nonciliated_Cells_Markers.csv'
wt_nonciliated_file = input_file_prefix + 'Nonexpressing_Nonciliated_Cells_Markers.csv'
final_file = "Processed_Histone_Ratio_Data.xlsx"

#
# Declare Arrays for Storing Data for all images and groupings
#

bPSelect = []

wt_ciliated_cell_treatment_group = []
transfected_ciliated_cell_treatment_group = []
wt_nonciliated_cell_treatment_group = []
transfected_nonciliated_cell_treatment_group = []

wt_ciliated_drug_treatment_group = []
transfected_ciliated_drug_treatment_group = []
wt_nonciliated_drug_treatment_group = []
transfected_nonciliated_drug_treatment_group = []

wt_ciliated_starve_treatment_group = []
transfected_ciliated_starve_treatment_group = []
wt_nonciliated_starve_treatment_group = []
transfected_nonciliated_starve_treatment_group = []

wt_ciliated_cell_type_group = []
transfected_ciliated_cell_type_group = []
wt_nonciliated_cell_type_group = []
transfected_nonciliated_cell_type_group = []

wt_ciliated_histone_modification_group = []
transfected_ciliated_histone_modification_group = []
wt_nonciliated_histone_modification_group = []
transfected_nonciliated_histone_modification_group = []

imgNum_wt_ciliated_raw_data = []
imgNum_wt_nonciliated_raw_data = []
imgNum_transfected_ciliated_raw_data = []
imgNum_transfected_nonciliated_raw_data = []

wt_ciliated_cell_image_num = []
transfected_ciliated_cell_image_num = []
wt_nonciliated_cell_image_num = []
transfected_nonciliated_cell_image_num = []

wt_ciliated_cell_object_num = []
transfected_ciliated_cell_object_num = []
wt_nonciliated_cell_object_num = []
transfected_nonciliated_cell_object_num = []

wt_ciliated_cell_DNA_Stain = []
transfected_ciliated_cell_DNA_Stain = []
wt_nonciliated_cell_DNA_Stain = []
transfected_nonciliated_cell_DNA_Stain = []

wt_ciliated_cell_Histone_Stain = []
transfected_ciliated_cell_Histone_Stain = []
wt_nonciliated_cell_Histone_Stain = []
transfected_nonciliated_cell_Histone_Stain = []

wt_ciliated_cell_Ratio = []
transfected_ciliated_cell_Ratio = []
wt_nonciliated_cell_Ratio = []
transfected_nonciliated_cell_Ratio = []

wt_ciliated_cell_images = []
transfected_ciliated_cell_images = []
wt_nonciliated_cell_images = []
transfected_nonciliated_cell_images = []

wt_ciliated_cell_objects = []
transfected_ciliated_cell_objects = []
wt_nonciliated_cell_objects = []
transfected_nonciliated_cell_objects = []

wt_ciliated_cell_DNA_Avg = []
transfected_ciliated_cell_DNA_Avg = []
wt_nonciliated_cell_DNA_Avg = []
transfected_nonciliated_cell_DNA_Avg = []

wt_ciliated_cell_Histone_Avg = []
transfected_ciliated_cell_Histone_Avg = []
wt_nonciliated_cell_Histone_Avg = []
transfected_nonciliated_cell_Histone_Avg = []

wt_ciliated_cell_Labels = []
transfected_ciliated_cell_Labels = []
wt_nonciliated_cell_Labels = []
transfected_nonciliated_cell_Labels = []


#
# Define Workbook Variable for Excel file
#

workbook = xlsxwriter.Workbook(output_path + final_file)
worksheet = workbook.add_worksheet()

#
# Begin Data Processing
#


try:
    # Attempt to navigate to the User Input Path and read the WT Ciliated Cells
    os.chdir(input_path)
    wt_ciliated_raw_data = pd.read_csv(wt_ciliated_file, usecols = ['ImageNumber','ObjectNumber','Intensity_MeanIntensity_Histone','Intensity_MeanIntensity_DNA','PathName_Channel_1'])
except:
    # If above Attempt Fails, end program and display error statement
    print("Error - No File Found at Path: " + wt_ciliated_file)
    quit()
else:
    imgCount1 = wt_ciliated_raw_data['ImageNumber'].max()
    if(math.isnan(imgCount1)):
        imgCount1 = 0;
    # Loops through and inserts raw WT Ciliated data in our data array
    for c in range(round(imgCount1)):
        x = wt_ciliated_raw_data[wt_ciliated_raw_data['ImageNumber'] == c]
        imgNum_wt_ciliated_raw_data.insert(len(imgNum_wt_ciliated_raw_data),x)
    
    # Declare Arrays for WT Ciliated Image #, OBJ #, Total DNA Intensity, Total Histone Stain Intensity, and Treatment Group
    wt_ciliated_cell_image_num = wt_ciliated_raw_data.iloc[0:(len(wt_ciliated_raw_data.iloc[:,0])),0]
    wt_ciliated_cell_object_num = wt_ciliated_raw_data.iloc[0:(len(wt_ciliated_raw_data.iloc[:,1])),1]
    wt_ciliated_cell_DNA_Stain = wt_ciliated_raw_data.iloc[0:(len(wt_ciliated_raw_data.iloc[:,3])),3]
    wt_ciliated_cell_Histone_Stain = wt_ciliated_raw_data.iloc[0:(len(wt_ciliated_raw_data.iloc[:,4])),4]
    wt_ciliated_cell_treatment_group = wt_ciliated_raw_data.iloc[0:(len(wt_ciliated_raw_data.iloc[:,2])),2]
    
    # Loops through and inserts processed WT Ciliated data into arrays to be exported into Excel, including Ratio, Image #, Obj #, DNA Intensity, Histone Stain Intensity, and 4 unique treatment groups
    for n in range(int(len(wt_ciliated_raw_data.iloc[:,0]))):
        wt_ciliated_cell_Ratio.insert(len(wt_ciliated_cell_Ratio),(wt_ciliated_cell_Histone_Stain[n] / wt_ciliated_cell_DNA_Stain[n]))
        wt_ciliated_cell_images.insert(len(wt_ciliated_cell_images),(wt_ciliated_cell_image_num[n]))
        wt_ciliated_cell_objects.insert(len(wt_ciliated_cell_objects),(wt_ciliated_cell_object_num[n]))
        wt_ciliated_cell_DNA_Avg.insert(len(wt_ciliated_cell_DNA_Avg),(wt_ciliated_cell_DNA_Stain[n]))
        wt_ciliated_cell_Histone_Avg.insert(len(wt_ciliated_cell_Histone_Avg),(wt_ciliated_cell_Histone_Stain[n]))
        wt_ciliated_cell_Labels.insert(len(wt_ciliated_cell_Labels),"Nonexpressing Ciliated")
        wt_ciliated_drug_treatment_group.insert(len(wt_ciliated_drug_treatment_group),wt_ciliated_cell_treatment_group[n])
        wt_ciliated_starve_treatment_group.insert(len(wt_ciliated_starve_treatment_group),wt_ciliated_cell_treatment_group[n])
        wt_ciliated_cell_type_group.insert(len(wt_ciliated_cell_type_group),wt_ciliated_cell_treatment_group[n])
        wt_ciliated_histone_modification_group.insert(len(wt_ciliated_histone_modification_group),wt_ciliated_cell_treatment_group[n])
        
    # Loops through and processes the 4 unique treatment groups using the folder tree path from the image
    for n in range(int(len(wt_ciliated_cell_treatment_group))):
        wt_ciliated_drug_treatment_group[n] = str(wt_ciliated_drug_treatment_group[n])
        x = wt_ciliated_drug_treatment_group[n].split("/")
        wt_ciliated_drug_treatment_group[n] = x[len(x)-1]
        wt_ciliated_starve_treatment_group[n] = x[len(x)-2]
        wt_ciliated_cell_type_group[n] = x[len(x)-3]
        wt_ciliated_histone_modification_group[n] = x[len(x)-4]
        
    # Define Exported Data for WT Ciliated Cells
    wt_ciliated_Data = pd.DataFrame([wt_ciliated_histone_modification_group,wt_ciliated_cell_type_group,wt_ciliated_starve_treatment_group,wt_ciliated_drug_treatment_group,wt_ciliated_cell_Labels,wt_ciliated_cell_images,wt_ciliated_cell_objects,wt_ciliated_cell_DNA_Avg,wt_ciliated_cell_Histone_Avg,wt_ciliated_cell_Ratio]).T
    wt_ciliated_Data.columns = ["Histone Modification","Cell Type","Starvation Treatment","Drug Treatment","Phenotype","Image","Object","SPY 505 Avg","Histone Avg","Histone / SPY 505 Ratio Avg"]
    
    try:
        # Attempt to navigate to the User Input Path and read the Transfected Nonciliated Cells
        os.chdir(input_path)
        transfected_nonciliated_raw_data = pd.read_csv(transfected_nonciliated_file, usecols = ['ImageNumber','ObjectNumber','Intensity_MeanIntensity_Histone','Intensity_MeanIntensity_DNA','PathName_Channel_1'])
    except:
        # If above Attempt Fails, end program and display error statement
        print("Error - No File Found at Path: " + transfected_nonciliated_file)
        quit()
    else:
        imgCount2 = transfected_nonciliated_raw_data['ImageNumber'].max()
        if(math.isnan(imgCount2)):
            imgCount2 = 0;
        # Loops through and inserts raw Transfected Nonciliated data in our data array
        for c in range(round(imgCount2)):
            x = transfected_nonciliated_raw_data[transfected_nonciliated_raw_data['ImageNumber'] == c]
            imgNum_transfected_nonciliated_raw_data.insert(len(imgNum_transfected_nonciliated_raw_data),x)
        # Declare Arrays for Transfected Nonciliated Image #, OBJ #, Total DNA Intensity, Total Histone Stain Intensity, and Treatment Group
        transfected_nonciliated_cell_image_num = transfected_nonciliated_raw_data.iloc[0:(len(transfected_nonciliated_raw_data.iloc[:,0])),0]
        transfected_nonciliated_cell_object_num = transfected_nonciliated_raw_data.iloc[0:(len(transfected_nonciliated_raw_data.iloc[:,1])),1]
        transfected_nonciliated_cell_DNA_Stain = transfected_nonciliated_raw_data.iloc[0:(len(transfected_nonciliated_raw_data.iloc[:,3])),3]
        transfected_nonciliated_cell_Histone_Stain = transfected_nonciliated_raw_data.iloc[0:(len(transfected_nonciliated_raw_data.iloc[:,4])),4]
        transfected_nonciliated_cell_treatment_group = transfected_nonciliated_raw_data.iloc[0:(len(transfected_nonciliated_raw_data.iloc[:,2])),2]
        # Loops through and inserts processed Transfected Nonciliated data into arrays to be exported into Excel, including Ratio, Image #, Obj #, DNA Intensity, Histone Stain Intensity, and 4 unique treatment groups
        for n in range(int(len((transfected_nonciliated_raw_data.iloc[:,0])))):
            print("Next")
            transfected_nonciliated_cell_Ratio.insert(len(transfected_nonciliated_cell_Ratio),(transfected_nonciliated_cell_Histone_Stain[n] / transfected_nonciliated_cell_DNA_Stain[n]))
            transfected_nonciliated_cell_images.insert(len(transfected_nonciliated_cell_images),(transfected_nonciliated_cell_image_num[n]))
            transfected_nonciliated_cell_objects.insert(len(transfected_nonciliated_cell_objects),(transfected_nonciliated_cell_object_num[n]))
            transfected_nonciliated_cell_DNA_Avg.insert(len(transfected_nonciliated_cell_DNA_Avg),(transfected_nonciliated_cell_DNA_Stain[n]))
            transfected_nonciliated_cell_Histone_Avg.insert(len(transfected_nonciliated_cell_Histone_Avg),(transfected_nonciliated_cell_Histone_Stain[n]))
            transfected_nonciliated_cell_Labels.insert(len(transfected_nonciliated_cell_Labels),"Expressing Nonciliated")
            transfected_nonciliated_drug_treatment_group.insert(len(transfected_nonciliated_drug_treatment_group),transfected_nonciliated_cell_treatment_group[n])
            transfected_nonciliated_starve_treatment_group.insert(len(transfected_nonciliated_starve_treatment_group),transfected_nonciliated_cell_treatment_group[n])
            transfected_nonciliated_cell_type_group.insert(len(transfected_nonciliated_cell_type_group),transfected_nonciliated_cell_treatment_group[n])
            transfected_nonciliated_histone_modification_group.insert(len(transfected_nonciliated_histone_modification_group),transfected_nonciliated_cell_treatment_group[n])
        # Loops through and processes the 4 unique treatment groups using the folder tree path from the image
        for n in range(int(len(transfected_nonciliated_cell_treatment_group))):
            transfected_nonciliated_drug_treatment_group[n] = str(transfected_nonciliated_drug_treatment_group[n])
            x = transfected_nonciliated_drug_treatment_group[n].split("/")
            transfected_nonciliated_drug_treatment_group[n] = x[len(x)-1]
            transfected_nonciliated_starve_treatment_group[n] = x[len(x)-2]
            transfected_nonciliated_cell_type_group[n] = x[len(x)-3]
            transfected_nonciliated_histone_modification_group[n] = x[len(x)-4]
            
        # Define Exported Data for Transfected Nonciliated Cells
        transfected_nonciliated_Data = pd.DataFrame([transfected_nonciliated_histone_modification_group,transfected_nonciliated_cell_type_group,transfected_nonciliated_starve_treatment_group,transfected_nonciliated_drug_treatment_group,transfected_nonciliated_cell_Labels,transfected_nonciliated_cell_images,transfected_nonciliated_cell_objects,transfected_nonciliated_cell_DNA_Avg,transfected_nonciliated_cell_Histone_Avg,transfected_nonciliated_cell_Ratio]).T
        transfected_nonciliated_Data.columns = ["Histone Modification","Cell Type","Starvation Treatment","Drug Treatment","Phenotype","Image","Object","SPY 505 Avg","Histone Avg","Histone / SPY 505 Ratio Avg"]
    
    try:
        # Attempt to navigate to the User Input Path and read the WT Nonciliated Cells
        os.chdir(input_path)
        wt_nonciliated_raw_data = pd.read_csv(wt_nonciliated_file, usecols = ['ImageNumber','ObjectNumber','Intensity_MeanIntensity_Histone','Intensity_MeanIntensity_DNA','PathName_Channel_1'])
    except:
        # If above Attempt Fails, end program and display error statement
        print("Error - No File Found at Path: " + wt_nonciliated_file)
        quit()
    else:
        imgCount3 = wt_nonciliated_raw_data['ImageNumber'].max()
        if(math.isnan(imgCount3)):
            imgCount3 = 0;
            
        # Loops through and inserts raw WT Nonciliated data in our data array
        for c in range(round(imgCount3)):
            x = wt_nonciliated_raw_data[wt_nonciliated_raw_data['ImageNumber'] == c]
            imgNum_wt_nonciliated_raw_data.insert(len(imgNum_wt_nonciliated_raw_data),x)
            
        # Declare Arrays for WT Nonciliated Image #, OBJ #, Total DNA Intensity, Total Histone Stain Intensity, and Treatment Group
        wt_nonciliated_cell_image_num = wt_nonciliated_raw_data.iloc[0:(len(wt_nonciliated_raw_data.iloc[:,0])),0]
        wt_nonciliated_cell_object_num = wt_nonciliated_raw_data.iloc[0:(len(wt_nonciliated_raw_data.iloc[:,1])),1]
        wt_nonciliated_cell_DNA_Stain = wt_nonciliated_raw_data.iloc[0:(len(wt_nonciliated_raw_data.iloc[:,3])),3]
        wt_nonciliated_cell_Histone_Stain = wt_nonciliated_raw_data.iloc[0:(len(wt_nonciliated_raw_data.iloc[:,4])),4]
        wt_nonciliated_cell_treatment_group = wt_nonciliated_raw_data.iloc[0:(len(wt_nonciliated_raw_data.iloc[:,2])),2]
        
        # Loops through and inserts processed WT Nonciliated data into arrays to be exported into Excel, including Ratio, Image #, Obj #, DNA Intensity, Histone Stain Intensity, and 4 unique treatment groups
        for n in range(int(len(wt_nonciliated_raw_data.iloc[:,0]))):
            wt_nonciliated_cell_Ratio.insert(len(wt_nonciliated_cell_Ratio),(wt_nonciliated_cell_Histone_Stain[n] / wt_nonciliated_cell_DNA_Stain[n]))
            wt_nonciliated_cell_images.insert(len(wt_nonciliated_cell_images),(wt_nonciliated_cell_image_num[n]))
            wt_nonciliated_cell_objects.insert(len(wt_nonciliated_cell_objects),(wt_nonciliated_cell_object_num[n]))
            wt_nonciliated_cell_DNA_Avg.insert(len(wt_nonciliated_cell_DNA_Avg),(wt_nonciliated_cell_DNA_Stain[n]))
            wt_nonciliated_cell_Histone_Avg.insert(len(wt_nonciliated_cell_Histone_Avg),(wt_nonciliated_cell_Histone_Stain[n]))
            wt_nonciliated_cell_Labels.insert(len(wt_nonciliated_cell_Labels),"Nonexpressing Nonciliated")
            wt_nonciliated_drug_treatment_group.insert(len(wt_nonciliated_drug_treatment_group),wt_nonciliated_cell_treatment_group[n])
            wt_nonciliated_starve_treatment_group.insert(len(wt_nonciliated_starve_treatment_group),wt_nonciliated_cell_treatment_group[n])
            wt_nonciliated_cell_type_group.insert(len(wt_nonciliated_cell_type_group),wt_nonciliated_cell_treatment_group[n])
            wt_nonciliated_histone_modification_group.insert(len(wt_nonciliated_histone_modification_group),wt_nonciliated_cell_treatment_group[n])
            
        # Loops through and processes the 4 unique treatment groups using the folder tree path from the image
        for n in range(int(len(wt_nonciliated_drug_treatment_group))):
            wt_nonciliated_drug_treatment_group[n] = str(wt_nonciliated_drug_treatment_group[n])
            x = wt_nonciliated_drug_treatment_group[n].split("/")
            wt_nonciliated_drug_treatment_group[n] = x[len(x)-1]
            wt_nonciliated_starve_treatment_group[n] = x[len(x)-2]
            wt_nonciliated_cell_type_group[n] = x[len(x)-3]
            wt_nonciliated_histone_modification_group[n] = x[len(x)-4]
            
        # Define Exported Data for WT Nonciliated Cells
        wt_nonciliated_Data = pd.DataFrame([wt_nonciliated_histone_modification_group,wt_nonciliated_cell_type_group,wt_nonciliated_starve_treatment_group,wt_nonciliated_drug_treatment_group,wt_nonciliated_cell_Labels,wt_nonciliated_cell_images,wt_nonciliated_cell_objects,wt_nonciliated_cell_DNA_Avg,wt_nonciliated_cell_Histone_Avg,wt_nonciliated_cell_Ratio]).T
        wt_nonciliated_Data.columns = ["Histone Modification","Cell Type","Starvation Treatment","Drug Treatment","Phenotype","Image","Object","SPY 505 Avg","Histone Avg","Histone / SPY 505 Ratio Avg"]
        
        try:
            # Attempt to navigate to the User Input Path and read the Transfected Ciliated Cells
            os.chdir(input_path)
            transfected_ciliated_raw_data = pd.read_csv(transfected_cilia_file, usecols = ['ImageNumber','ObjectNumber','Intensity_MeanIntensity_Histone','Intensity_MeanIntensity_DNA','PathName_Channel_1'])
        except:
            # If above Attempt Fails, end program and display error statement
            print("Error - No File Found at Path: " + transfected_cilia_file)
            quit()
        else:
            imgCount4 = transfected_ciliated_raw_data['ImageNumber'].max()
            if(math.isnan(imgCount4)):
                imgCount4 = 0;
                
            # Loops through and inserts raw WT Nonciliated data in our data array
            for c in range(round(imgCount4)):
                x = transfected_ciliated_raw_data[transfected_ciliated_raw_data['ImageNumber'] == c]
                imgNum_transfected_ciliated_raw_data.insert(len(imgNum_transfected_ciliated_raw_data),x)
                
            # Declare Arrays for Transfected Ciliated Image #, OBJ #, Total DNA Intensity, Total Histone Stain Intensity, and Treatment Group
            transfected_ciliated_cell_image_num = transfected_ciliated_raw_data.iloc[0:(len(transfected_ciliated_raw_data.iloc[:,0])),0]
            transfected_ciliated_cell_object_num = transfected_ciliated_raw_data.iloc[0:(len(transfected_ciliated_raw_data.iloc[:,1])),1]
            transfected_ciliated_cell_DNA_Stain = transfected_ciliated_raw_data.iloc[0:(len(transfected_ciliated_raw_data.iloc[:,3])),3]
            transfected_ciliated_cell_Histone_Stain = transfected_ciliated_raw_data.iloc[0:(len(transfected_ciliated_raw_data.iloc[:,4])),4]
            transfected_ciliated_cell_treatment_group = transfected_ciliated_raw_data.iloc[0:(len(transfected_ciliated_raw_data.iloc[:,2])),2]
            
            # Loops through and  inserts processed Transfected Ciliated data into arrays to be exported into Excel, including Ratio, Image #, Obj #, DNA Intensity, Histone Stain Intensity, and 4 unique treatment groups
            for n in range(int(len(transfected_ciliated_raw_data.iloc[:,0]))):
                transfected_ciliated_cell_Ratio.insert(len(transfected_ciliated_cell_Ratio),(transfected_ciliated_cell_Histone_Stain[n] / transfected_ciliated_cell_DNA_Stain[n]))
                transfected_ciliated_cell_images.insert(len(transfected_ciliated_cell_images),(transfected_ciliated_cell_image_num[n]))
                transfected_ciliated_cell_objects.insert(len(transfected_ciliated_cell_objects),(transfected_ciliated_cell_object_num[n]))
                transfected_ciliated_cell_DNA_Avg.insert(len(transfected_ciliated_cell_DNA_Avg),(transfected_ciliated_cell_DNA_Stain[n]))
                transfected_ciliated_cell_Histone_Avg.insert(len(transfected_ciliated_cell_Histone_Avg),(transfected_ciliated_cell_Histone_Stain[n]))
                transfected_ciliated_cell_Labels.insert(len(transfected_ciliated_cell_Labels),"Expressing Ciliated")
                transfected_ciliated_drug_treatment_group.insert(len(transfected_ciliated_drug_treatment_group),transfected_ciliated_cell_treatment_group[n])
                transfected_ciliated_starve_treatment_group.insert(len(transfected_ciliated_starve_treatment_group),transfected_ciliated_cell_treatment_group[n])
                transfected_ciliated_cell_type_group.insert(len(transfected_ciliated_cell_type_group),transfected_ciliated_cell_treatment_group[n])
                transfected_ciliated_histone_modification_group.insert(len(transfected_ciliated_histone_modification_group),transfected_ciliated_cell_treatment_group[n])
                
            # Loops through and processes the 4 unique treatment groups using the folder tree path from the image
            for n in range(int(len(transfected_ciliated_drug_treatment_group))):
                transfected_ciliated_drug_treatment_group[n] = str(transfected_ciliated_drug_treatment_group[n])
                x = transfected_ciliated_drug_treatment_group[n].split("/")
                transfected_ciliated_drug_treatment_group[n] = x[len(x)-1]
                transfected_ciliated_starve_treatment_group[n] = x[len(x)-2]
                transfected_ciliated_cell_type_group[n] = x[len(x)-3]
                transfected_ciliated_histone_modification_group[n] = x[len(x)-4]
                
            # Define Exported Data for Transfected Ciliated Cells
            transfected_ciliated_Data = pd.DataFrame([transfected_ciliated_histone_modification_group,transfected_ciliated_cell_type_group,transfected_ciliated_starve_treatment_group,transfected_ciliated_drug_treatment_group,transfected_ciliated_cell_Labels,transfected_ciliated_cell_images,transfected_ciliated_cell_objects,transfected_ciliated_cell_DNA_Avg,transfected_ciliated_cell_Histone_Avg,transfected_ciliated_cell_Ratio]).T
            transfected_ciliated_Data.columns = ["Histone Modification","Cell Type","Starvation Treatment","Drug Treatment","Phenotype","Image","Object","SPY 505 Avg","Histone Avg","Histone / SPY 505 Ratio Avg"]
                
            try:
                #
                # Try to consolidate all groups into one exportable Excel File (.xlsx)
                #
                
                # Prepare to write data
                os.chdir(output_path)
                bPAll_1 = pd.DataFrame((np.array(wt_ciliated_cell_Ratio)).transpose(),columns = ["Nonexpressing Ciliated"])
                bPAll_2 = pd.DataFrame((np.array(wt_nonciliated_cell_Ratio)).transpose(),columns = ["Nonexpressing Nonciliated"])
                bPAll_3 = pd.DataFrame((np.array(transfected_ciliated_cell_Ratio)).transpose(),columns = ["Expressing Ciliated"])
                bPAll_4 = pd.DataFrame((np.array(transfected_nonciliated_cell_Ratio)).transpose(),columns = ["Expressing Nonciliated"])
                combineAll = [bPAll_1,bPAll_2,bPAll_3,bPAll_4]
                bPAll = pd.concat(combineAll)
                
                # Write Data into Excel File
                with pd.ExcelWriter(final_file, engine='xlsxwriter') as writer:
                    # Write data into excel sheet 'Nonexpressing Ciliated'
                    wt_ciliated_Data.to_excel(writer, sheet_name = "Nonexpressing Ciliated", index = False)
                    sheet1 = writer.sheets["Nonexpressing Ciliated"]
                    sheet1.set_column(0,0,20)
                    sheet1.set_column(1,2,20)
                    sheet1.set_column(1,2,20)
                    sheet1.set_column(3,3,20)
                    sheet1.set_column(4,4,50)
                    sheet1.set_column(5,5,5)
                    sheet1.set_column(6,6,5)
                    sheet1.set_column(7,7,20)
                    sheet1.set_column(8,8,20)
                    sheet1.set_column(9,9,20)
                    
                    # Write data into excel sheet 'Nonexpressing Nonciliated'
                    wt_nonciliated_Data.to_excel(writer, sheet_name = "Nonexpressing Nonciliated", index = False)
                    sheet2 = writer.sheets["Nonexpressing Nonciliated"]
                    sheet2.set_column(0,0,20)
                    sheet2.set_column(1,2,20)
                    sheet2.set_column(1,2,20)
                    sheet2.set_column(3,3,20)
                    sheet2.set_column(4,4,50)
                    sheet2.set_column(5,5,5)
                    sheet2.set_column(6,6,5)
                    sheet2.set_column(7,7,20)
                    sheet2.set_column(8,8,20)
                    sheet2.set_column(9,9,20)
                    
                    # Write data into excel sheet 'Expressing Ciliated'
                    transfected_ciliated_Data.to_excel(writer, sheet_name = "Expressing Ciliated", index = False)
                    sheet3 = writer.sheets["Expressing Ciliated"]
                    sheet3.set_column(0,0,20)
                    sheet3.set_column(1,2,20)
                    sheet3.set_column(1,2,20)
                    sheet3.set_column(3,3,20)
                    sheet3.set_column(4,4,50)
                    sheet3.set_column(5,5,5)
                    sheet3.set_column(6,6,5)
                    sheet3.set_column(7,7,20)
                    sheet3.set_column(8,8,20)
                    sheet3.set_column(9,9,20)
                    
                    # Write data into excel sheet 'Expressing Nonciliated'
                    transfected_nonciliated_Data.to_excel(writer, sheet_name = "Expressing Nonciliated", index = False)
                    sheet4 = writer.sheets["Expressing Nonciliated"]
                    sheet4.set_column(0,0,20)
                    sheet4.set_column(1,2,20)
                    sheet4.set_column(1,2,20)
                    sheet4.set_column(3,3,20)
                    sheet4.set_column(4,4,50)
                    sheet4.set_column(5,5,5)
                    sheet4.set_column(6,6,5)
                    sheet4.set_column(7,7,20)
                    sheet4.set_column(8,8,20)
                    sheet4.set_column(9,9,20)

            except:
                # If above Attempt Fails, end program and display error statement
                print("Error - Compilation of Data Failed")
                quit()
    print("DATA PROCESSING DONE")
