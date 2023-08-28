# created 8/28/2023 from "Testing XML Parser for cobas 8800" section for getting  just the results info from the XML file
# running this script will get the Results data a from a cobas 8800 run .XML file
# output as an excel file

import bs4 as bs
import pandas as pd
from functools import reduce
import numpy as np
from openpyxl import load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows
import os
from datetime import datetime
import time


def rename_keys(my_dict,suffix):
    """ rename all keys in a dictionary with an added _suffix"""
    # create new key names from current by adding _suffix to it
    #modified from : https://www.alixaprodev.com/2022/07/rename-dictionary-key-in-python.html#:~:text=Rename%20a%20dictionary%20key%20in%20Python%201%20Get,value%20of%20the%20older%20key%20to%20knew%20key
    new_keys=[]
    suffix='_'+str(suffix)
    for x in my_dict:
        new=x+suffix
        new_keys.append(new)
    
    #create and return new dictionary
    d1 = dict( zip( list(my_dict.keys()), new_keys) )
    return {d1[oldK]: value for oldK, value in my_dict.items()}



def infoFromTestOrder(one):
    """ get Specimen and TestResult data from a TestOrder tag"""
    row={}
    for x in one.children:
        # get sample info like barcode
        if x.name =="Specimen":
            for y in x.children:
                row.update(y.attrs)

        # get test results info
        elif x.name =="TestResults":
            for num,y in enumerate(x.children):
                if num==0:
                    row.update(y.attrs)
                else:
                    new_dict= rename_keys(y.attrs,num)
                    row.update(new_dict)
                    multipleTargets=False
               
    return row

### declare the function for parsing specific query to return a list of dicts
def parse_xml(query):
    search = soup.find_all(query)
    # parse content into dicts
    dicts=[]
    for data in search:
        if data.Carrier:
            data.attrs.update(data.Carrier.attrs)
        dicts.append(data.attrs)
    return dicts


if __name__=="__main__":
    print("Select an exported XML file for results data extraction")
    ### Select file name 
    # copied from https://stackoverflow.com/questions/20790926/ipython-notebook-open-select-file-with-gui-qt-dialog 
    try:
        from tkinter import Tk
        from tkFileDialog import askopenfilenames
    except:
        from tkinter import Tk
        from tkinter import filedialog
    Tk().withdraw()
    filenames = filedialog.askopenfilenames() 
    print (filenames[0])
    print("Extracting...")
    
    name=filenames[0]
    # reading content
    file = open(name, "r")
    contents = file.read()
    # parsing
    soup = bs.BeautifulSoup(contents, 'xml')


    #### display Sample info from parseing XML file:
    #query xml for sample info
    query='Sample'
    dicts=parse_xml(query)
    #assemble into dataframe
    dfs=pd.DataFrame.from_dict(dicts)
    try:
        #dfs=dfs.drop_duplicates()
        dfs=dfs.sort_values("CreationDateTime")
        dfs=dfs.reset_index()
        dfs=dfs.reset_index()
        dfs['temp']=dfs['level_0']
        dfs#.to_excel("temp_dfs.xlsx")
    except:
        print("No samples found")


    search = soup.find_all('TestOrder')
    one=search[0]


    ### create dfResults
    dfResults="" # added so that it doesnt get longer infinitely with re-runs

    #get testOrder info (each row of df)
    for one in search:
        row = infoFromTestOrder(one)

        #convert row to df
        df=pd.DataFrame([row])

        #outline general header terms desired
        GeneralHeaders=['CreationDateTime','Name','Barcode','FinalInterpretation','CT','Position','Info','SpecimenClass','Target']

        #get all fields that contain a general header
        SpecificHeaders=[]
        for x in df:
            for y in GeneralHeaders:
                if y in x:
                    SpecificHeaders.append(x)

        #df w/ desired fields
        df=df[SpecificHeaders]

        #append to main df
        try:
            dfResults=pd.concat([dfResults,df],axis=0)
        except:
            dfResults=df
    dfResults

    # # Name and Create the output file
    ### Get the test name
    testName=""
    for x in dfResults:
        if "Name" in x:
            testName=dfResults[x].unique()[0] 
    # add barcode values for controls
    TestName=""
    HxV=['hiv','hbv','hcv']
    if testName.lower() in HxV:
        if 'hiv' in testName:
            TestName="HIV"
        elif 'hbv' in testName:
            TestName="HBV"
        elif 'hcv' in testName:
            TestName="HCV"        
    elif "HPV" in testName:
        TestName="HPV"
    elif "CT" in testName or "NG" in testName:
        TestName="CTNG"
    elif "tgt" in testName.lower():
        TestName="SARS"
    else:
        print("test name not found for labelling control samples")
    print("test is: ",TestName)

    ### get batch # 
    query='OrderGroup'  # batch num 
    dicts=parse_xml(query)
    dfb=pd.DataFrame.from_dict(dicts)
    BatchNum=max([int(x) for x in dfb['OrderId'].unique()])
    BatchNum="b"+str(BatchNum)
    print("Batch Number is: ",BatchNum)

    ### create excel file
    OutputFileName_prefix="C:/Users/PCA0551/Desktop/"
    OutputFileName_base="CobasResults"
    OutputFileName_suffix=".xlsx"
    OutputFileName_date=datetime.today().strftime('%Y_%m_%d')

    OutputFileName=OutputFileName_prefix+OutputFileName_base+"_"+TestName+"_"+BatchNum+"_"+OutputFileName_date+""+OutputFileName_suffix
    print(OutputFileName)
    dfResults.to_excel(OutputFileName,index=False)

    print("\n \n File created, it can be found on the Desktop. You may now exit or Window will close shortly")        
    time.sleep(30)
