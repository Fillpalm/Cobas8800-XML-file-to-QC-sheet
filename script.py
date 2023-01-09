############## This script splits up the QC sheet info needed into two categories: Reagent info and Results info

import bs4 as bs
import pandas as pd
from functools import reduce
import numpy as np
from openpyxl import load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows
import os
import time

### declare the function for parsing specific query to return a list of dicts
def parse_xml(query,soup):
    search = soup.find_all(query)
    # parse content into dicts
    dicts=[]
    for data in search:
        if data.Carrier:
            data.attrs.update(data.Carrier.attrs)
        dicts.append(data.attrs)   
    return dicts
    
def getReagents(soup):
    #query xml for reagents
    dicts=parse_xml('ReagentContainer',soup)
    #assemble into dataframe
    dfRs=pd.DataFrame.from_dict(dicts)
    dfRs=dfRs[['ReagentName','SerialNumber','LotNumber','Expiration','OnboardTime']] # for reagents
    dfRs=dfRs.drop_duplicates()
    dfRs=dfRs.replace('(-) C', '(-) Ctrl')
    return dfRs

def getReagentKit(soup,test):
    KitMaterialNumbers={"HBV":"9040820190_A",
             "HCV":"9040765190_A",
             "HIV":"9040803190_A", 
             "CTNG":"7460066190_A",
             "SARS": "9343733190_A",
             "HPV":"7460155190_A",}
    #query xml for kit
    dicts=parse_xml('InventoryItemTracking',soup)

    #assemble into dataframe
    dfK=pd.DataFrame.from_dict(dicts)
    dfK=dfK[['SerialNumber','LotNumber','Expiration','OnboardTime','MaterialNumber']] # for reagents
    dfK=dfK.drop_duplicates()
    dfK['ReagentName']="Reagent Kit"
    dfK=dfK.loc[dfK['MaterialNumber']==KitMaterialNumbers[test]]
    return dfK

def rename_keys(my_dict,suffix):
    """ rename all keys in a dictionary with an added _suffix"""
    # create new key names from current by adding _suffix to it
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

def getResults(soup):
    search = soup.find_all('TestOrder')
    rows=[]
    for one in search:
        rows.append(infoFromTestOrder(one))
    dfResults=pd.DataFrame(rows)
    return dfResults

def getTestName(dfResults):
    testName=""
    for x in dfResults:
        if "name" in x.lower():
            testName=dfResults[x][0]
    return testName

def addControlLabels(dfResults,testName):
    ### add barcode values for control     
    # add barcode values for controls
    PosControlName=""
    HxV=['hiv','hbv','hcv']
    if testName.lower() in HxV:
        print("control is HxV!!!")
        dfResults.loc[len(dfResults)-3,'Barcode']="HxV H (+) C"
        PosControlName="HxV L (+) C"
    elif "HPV" in testName:
        print("control is HPV!!!")
        PosControlName="HPV (+) C"
    elif "CT" in testName or "NG" in testName:
        print("control is CT/NG!!!")
        PosControlName="CT/NG (+) C"
    elif "tgt" in testName.lower():
        print("control is SARS!!!")
        PosControlName="SARS-CoV-2 (+) C"
    else:
        print("test name not found for labelling control samples")
    dfResults.loc[len(dfResults)-2,'Barcode']=PosControlName
    dfResults.loc[len(dfResults)-1,'Barcode']= "(-) Ctrl"
    return dfResults,PosControlName

def assignReagentVariables(dfRs,PosControlName):
    ### reagent info
    #kit lot
    ReagentKitLot=str(dfRs.loc[dfRs['ReagentName']=='Reagent Kit']['LotNumber'].values[0])
    #PC lot #
    PCLotNo=str(dfRs.loc[dfRs['ReagentName']==PosControlName]['LotNumber'].values[0])
    #NC lot#
    NCLotNo=str(dfRs.loc[dfRs['ReagentName']=='(-) Ctrl']['LotNumber'].values[0])
    return(ReagentKitLot,PCLotNo,NCLotNo)

def assignResultsVariablesHIV(dfResults):
    ### HIV
    #HPC CT
    HPC_CT=float(dfResults.loc[dfResults['Barcode']=='HxV H (+) C']['CT'].values[0])
    #HPC IU
    HPC_IU=float(dfResults.loc[dfResults['Barcode']=='HxV H (+) C']['Value'].values[0])
    #LPC CT
    LPC_CT=float(dfResults.loc[dfResults['Barcode']=='HxV L (+) C']['CT'].values[0])
    #LPC IU
    LPC_IU=float(dfResults.loc[dfResults['Barcode']=='HxV L (+) C']['Value'].values[0])
    #NC Result
    NC_CT=float(dfResults.loc[dfResults['Barcode']=='(-) Ctrl']['CT'].values[0])
    if NC_CT >0:
        pass
    else:
        NC_CT="None"
    return(HPC_CT,HPC_IU,LPC_CT,LPC_IU,NC_CT)
    
def assignResultsVariablesSARS(dfResults):
    tgt1_positives=0
    tgt2_positives=0
    invalidSamples=0
    for x in dfResults:
        if "FinalInterpretation" in x:
            for y in dfResults[x].unique():
                if "Positive"in y:
                    tgt1_positives=len(dfResults[x].loc[dfResults[x]==y])
        if "FinalInterpretation_1" in x:
            for y in dfResults[x].unique():
                if "Positive"in y:
                    tgt2_positives=len(dfResults[x].loc[dfResults[x]==y])
                if "invalid"in y.lower():
                    invalidSamples=len(dfResults[x].loc[dfResults[x]==y])


    return(tgt1_positives,tgt2_positives,invalidSamples)  
    
    
def assignResultsVariablesHPV(dfResults):
    ### HPV
    HPV16Positives=0
    HPV18Positives=0
    HPVotherPositives=0
    invalidSamples=0
    for x in dfResults:
        if "FinalInterpretation" in x:
            for y in dfResults[x].unique():
                if "Positive"in y:
                    #print(y,"positives: ",len(dfResults[x].loc[dfResults[x]==y]), '\n')
                    if "16" in y:
                        HPV16Positives=len(dfResults[x].loc[dfResults[x]==y]) # no minus 1 for controls, they are valid not positive
                    elif "18" in y:
                        HPV18Positives=len(dfResults[x].loc[dfResults[x]==y])
                    elif "other" in y.lower():
                        HPVotherPositives=len(dfResults[x].loc[dfResults[x]==y])
                if "invalid"in y.lower():
                    invalidSamples=len(dfResults[x].loc[dfResults[x]==y])
    return(HPV16Positives,HPV18Positives,HPVotherPositives,invalidSamples)

def assignResultVariablesCTNG(dfResults):
    ### CTNG    
    swabSampleCount=0
    urineSampleCount=0
    thinprepSampleCount=0
    CTpositives=0
    NGpostives=0
    invalidSamples=0
    for x in dfResults:
        if "FinalInterpretation" in x:
            for y in dfResults[x].unique():
                if "Positive"in y:
                    #print(y,"positives: ",len(dfResults[x].loc[dfResults[x]==y]), '\n')
                    if "CT" in y:
                        CTpositives=len(dfResults[x].loc[dfResults[x]==y]) #no minus 1 for controls, they are valid not positive
                    elif "NG" in y:
                        NGpostives=len(dfResults[x].loc[dfResults[x]==y])
                if "invalid"in y.lower():
                    invalidSamples=len(dfResults[x].loc[dfResults[x]==y])
        if "Info" in x:
            for y in dfResults[x].unique():
                ymod=str(y).lower()
                if "swab"in ymod:
                    swabSampleCount=len(dfResults[x].loc[dfResults[x]==y])
                elif "urine"in ymod:
                    urineSampleCount=len(dfResults[x].loc[dfResults[x]==y])
                elif "preserv"in ymod:
                    thinprepSampleCount=len(dfResults[x].loc[dfResults[x]==y])
    #print(" Swabs ",swabSampleCount,"\n","Urine: ",urineSampleCount,"\n","thinpreps: : ",thinprepSampleCount)
    #print(" CT positives: ",CTpositives,"\n","NG positives: ",NGpostives,"\n","invalids: : ",invalidSamples)
    return(swabSampleCount,urineSampleCount,thinprepSampleCount,CTpositives,NGpostives,invalidSamples)

def prepQCdata(test,soup):
    # path as copied from file explorer window
    #added extra "\" in front of "\n"'s as required
    path=str('M:\MP Molecular Pathology\\NJ_Mol_Virology\\NJ Routine\QC Sheets\COBAS 8800\\test auto.xlsx')
    #convert to linux
    cpath=path.replace('\\','/')
    ### get batch # for joining 
    dicts=parse_xml('OrderGroup',soup)
    dfb=pd.DataFrame.from_dict(dicts)
    BatchNum=max([int(x) for x in dfb['OrderId'].unique()])

    #get and read correct sheet
    sheets={"HCV":"HCVrunQC",
           "HIV":"HIVrunQC",
           "HBV":"HBVrunQC",
           "HPV":"HPVrunQC",
           "SARS": "SARSrunQC",
           "CTNG":"CTNGrunQC"}
    sheet=sheets[test.lower().upper()]
    dftest=pd.read_excel(cpath, sheet_name=sheet,header=2)
    dftest=dftest.loc[dftest['CONTROL BATCH #']==BatchNum]
    if len(dftest)>0:
        print("QC file loaded!")
    else:
        print("Control batch ID not found")
    return dftest,sheet,BatchNum
        
def writeToQCSheets(dftest,sheet,BatchNum):
    # path as copied from file explorer window
    #added extra "\" in front of "\n"'s as required
    path=str('M:\MP Molecular Pathology\\NJ_Mol_Virology\\NJ Routine\QC Sheets\COBAS 8800\\test auto.xlsx')
    ### use pyxl to write to spreadsheet in order to handle dateTime object
    #get lastrow number to write to
    wb = load_workbook(path)
    ws = wb[sheet]
    try:
        dftestValues=[x for x in dataframe_to_rows(dftest, index=False, header=False)][0]
    except:
        print("No batch number found in sheet: ",sheet,"for control batch ID: ",BatchNum)
    #Find Row for desired batch number
    for row in ws.iter_cols(min_col=3,max_col=3):
        for RowNum,cell in enumerate(row):
            if (cell.value == BatchNum):
                targetRow=RowNum+1
    #Write values to excel sheet            
    for num,x in enumerate(ws[targetRow]):
        x.value=dftestValues[num]
    wb.save(path)
    wb.close()
    
    
def Main(name):
    # reading file content
    file = open(name, "r")
    contents = file.read()
    # parsing
    soup = bs.BeautifulSoup(contents, 'xml')
    file.close()
    
    ### get test name from file ##################################
    # requires test name in file
    # USE DefinitionId ID OR MaterialNumber NOT SERIAL NUMBER    
    tests=['HBV','HCV','HIV','HPV','CTNG','SARS']
    test=""
    try:
        for x in tests:
            if x.lower() in name.lower():
                print("test is: ",x)
                test=x
    except:
        print("test not found")
    

    ### Reagents info extraction################################################################
    #get regeagents
    dfRs=getReagents(soup)
    #get reagentKit
    dfKit=getReagentKit(soup,test)
    #join reagents w kits info
    dfRs=pd.concat([dfRs,dfKit])
    
    
    ### Results info extraction ################################################################
    dfResults=getResults(soup) 
        
    #get testName
    testName=getTestName(dfResults)
    #add control labels for results
    dfResults,PosControlName=addControlLabels(dfResults,testName)
    sampleNum=len(dfResults)
    
    #assign reagent variables that will be written to QC sheet
    ReagentKitLot,PCLotNo,NCLotNo=assignReagentVariables(dfRs,PosControlName) 
    
    #assign result variables that will be written to QC sheet
    testName=testName.strip().lower()
    HxV=['hiv','hbv','hcv']
    if testName in HxV:
        HPC_CT,HPC_IU,LPC_CT,LPC_IU,NC_CT=assignResultsVariablesHIV(dfResults)
        dftest,sheet,BatchNum=prepQCdata(test,soup)
    elif "hpv" in testName:
        HPV16Positives,HPV18Positives,HPVotherPositives,invalidSamples=assignResultsVariablesHPV(dfResults)
        dftest,sheet,BatchNum=prepQCdata(test,soup)
    elif "ct" in testName or "ng" in testName:
        swabSampleCount,urineSampleCount,thinprepSampleCount,CTpositives,NGpostives,invalidSamples=assignResultVariablesCTNG(dfResults)
        dftest,sheet,BatchNum=prepQCdata(test,soup)
    elif "tgt" in testName.lower(): # SARS
        tgt1_positives,tgt2_positives,invalidSamples=assignResultsVariablesSARS(dfResults) 
        dftest,sheet,BatchNum=prepQCdata(test,soup)
    else:
        print("testName not found: ", testName)
        
    ### assign values
    #assign reagent values
    dftest['auto REAGENT KIT LOT']=ReagentKitLot
    dftest['auto POSITIVE CTRL KIT LOT#']=PCLotNo
    dftest['auto NEGATIVE CTRL LOT#']=NCLotNo
    #assign result values
    dftest['Samples + controls']=sampleNum
    
    if testName in HxV:
        dftest['CT VALUE OF HIGH POS']=HPC_CT
        dftest['HIGH POS CTRL RESULT (IU/mL)']=HPC_IU
        dftest['HIGH POS CTRL Result (Log IU/mL)']=np.log10(HPC_IU)
        dftest['CT VALUE OF LOW POS']=LPC_CT
        dftest['LOW POS CONTROL RESULT (IU/mL)']=LPC_IU
        dftest['LOW POS CONTROL (Log IU/mL)']=np.log10(LPC_IU)
        dftest['NEGATIVE CTRL RESULT']=NC_CT
    elif "hpv" in testName:
        dftest['OTHER HR HPV POSITIVE']=HPVotherPositives
        dftest['HPV 16 POSITIVE']=HPV16Positives
        dftest['HPV 18 POSITIVE']=HPV18Positives
        dftest['INVALID']=invalidSamples
    elif "ct" in testName or "ng" in testName:
        dftest['CT POSITIVE']=CTpositives
        dftest['NG POSITIVE']=NGpostives
        dftest['INVALID']=invalidSamples
        dftest['SWABS']=swabSampleCount
        dftest['URINES']=urineSampleCount
        dftest['THINPREPS']=thinprepSampleCount
    elif "tgt" in testName.lower(): # SARS
        dftest['Target1 positive']=tgt1_positives
        dftest['Target2 positive']=tgt2_positives
        dftest['INVALID']=invalidSamples
    else:
        print("test name not found")

    writeToQCSheets(dftest,sheet,BatchNum)
    
if __name__=="__main__":
    #list files
    newFileFolder="new XML files/"
    for x in os.listdir(newFileFolder):
        if x[0]=='b':
            print("Processing file: ",x)
            path=newFileFolder+x 
            try:
                Main(path)
                newPath="old XML files/"+x  
                os.rename(path,newPath)
                print(" complete!",x)
            except:
                print("\n \n error found, see above text \n \n")
    print("\n \n Program complete. Window will close shortly")        
    time.sleep(10)
