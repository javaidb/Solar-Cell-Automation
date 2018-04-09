import os
import glob
#Find relevant CSVs in folder
path = os.getcwd()
extension = 'txt'
os.chdir(path)
csvresult = [i for i in glob.glob('*.{}'.format(extension))]
print("Tabulating Solar Cell Data from CSVs...")
print(csvresult)

import numpy as np
import matplotlib.pyplot as plt
import pandas as pd
import statistics
import re

k=0
#iterate through TXTs
while k < len(csvresult):
    allsample=[]
    path = os.getcwd()
    path = path + "/" + csvresult[k]
    solar=pd.read_csv(path, delimiter=",")
    m=0
    #iterate through each line
    while m < len(solar):
        #Remove tabs (\t) from array
        solar.values[m][0]=re.split(r'\t+',solar.values[m][0].rstrip('\t'))
        n=0
        solartry=[]
        #iterate through each element
        while n < len(solar.values[m][0]):
            #Place every two values in an array
            solarvolt=solar.values[m][0][0+n]
            solarcurr=solar.values[m][0][1+n]
            solartry.append([solarvolt,solarcurr])
            n=n+2
            continue
        solar.values[m][0]=solartry
        m=m+1
        continue
    #Loop through samples (grouped from previous lines of code)
    samp=0
    print('Iterating through samples:')
    sizecount=0
    while samp < n/2:
        eachsample=[]
        #Loop through each line
        lin=0
        solar_volt=[]
        solar_curr=[]
        while lin < len(solar):
            solar_volt.append(solar.values[lin][0][samp][0])
            solar_curr.append(solar.values[lin][0][samp][1])
            lin=lin+1
            continue
        if solar_curr[2] == '0.0572':
            sizecount = sizecount + 1
        #Extract values
        voltvals=solar_volt[11:]
        voltvals=[float(x) for x in voltvals]
        currvals=solar_curr[11:]
        currvals=[float(x) for x in currvals]
        currvals=[x*(-1) for x in currvals]
        r=0
        voltcount=0
        currcount=0
        while r < len(voltvals):
            if voltvals[r] < 0:
                voltcount=voltcount+1
            if currvals[r] < 0:
                currcount=currcount+1
            r=r+1
            continue
        i=0
        solar_pow=[]
        #Make power column
        while i < len(voltvals):
            solar_pow.append(voltvals[i]*currvals[i])
            i=i+1
            continue
        eachsample.append(solar_curr[0])
        #Do calculations
        jsc=currvals[0]
        if currcount == len(voltvals):
            voc=voltvals[currcount-1]
        else:
            voc=(voltvals[currcount-1]+voltvals[currcount])/2
        pmax=min(solar_pow)
        jmp=currvals[solar_pow.index(pmax)]
        vmp=voltvals[solar_pow.index(pmax)]
        ff=abs(pmax/(voc*jsc))
        pce=abs((pmax*1000)/float(solar_curr[2]))
        eachsample.extend((jsc,voc,pmax,jmp,vmp,ff,pce))
        allsample.append(eachsample) 
        samp=samp+1
        print(samp)
        continue
    k=k+1
    #----------------------------------------------- EXCEL EXPORT-----------------------------------------------------

    # Create a Pandas dataframe title for data.
    df0 = pd.DataFrame({'Data': ['Sample','Jsc (A)','Voc (V)', 'Pmax (W)','Jmp (A)','Vmp (V)','FF','PCE (%)']}).T

    writer = pd.ExcelWriter('Solar Cell Data Summary.xlsx', engine='xlsxwriter')

    #Loop through organic cell data
    j=1
    index=1
    jsctot=[]
    voctot=[]
    pmaxtot=[]
    jmptot=[]
    vmptot=[]
    fftot=[]
    pcetot=[]
    while j < sizecount + 1:
        samplename=allsample[j][0]
        jsc=allsample[j][1]
        jsctot.append(jsc)
        voc=allsample[j][2]
        voctot.append(voc)
        pmax=allsample[j][3]
        pmaxtot.append(pmax)
        jmp=allsample[j][4]
        jmptot.append(jmp)
        vmp=allsample[j][5]
        vmptot.append(vmp)
        ff=allsample[j][6]
        fftot.append(ff)
        pce=allsample[j][7]
        pcetot.append(pce)
        individualdata=(samplename,jsc,voc,pmax,jmp,vmp,ff,pce)
        df = pd.DataFrame({'invisible header': individualdata}).T
        df.to_excel(writer, sheet_name='Sheet1', startrow=index, index=False,header =False)
        j=j+1
        index=index+1
        continue

    jscavg=sum(jsctot)/len(jsctot)
    jscdev=statistics.stdev(jsctot)
    vocavg=sum(voctot)/len(voctot)
    vocdev=statistics.stdev(voctot)
    pmaxavg=sum(pmaxtot)/len(pmaxtot)
    pmaxdev=statistics.stdev(pmaxtot)
    jmpavg=sum(jmptot)/len(jmptot)
    jmpdev=statistics.stdev(jmptot)
    vmpavg=sum(vmptot)/len(vmptot)
    vmpdev=statistics.stdev(vmptot)
    ffavg=sum(fftot)/len(fftot)
    ffdev=statistics.stdev(fftot)
    pceavg=sum(pcetot)/len(pcetot)
    pcedev=statistics.stdev(pcetot)
    averagedata=(jscavg,vocavg,pmaxavg,jmpavg,vmpavg,ffavg,pceavg)
    stdevdata=(jscdev,vocdev,pmaxdev,jmpdev,vmpdev,ffdev,pcedev)
    df1 = pd.DataFrame({'invisible header': averagedata}).T
    df2 = pd.DataFrame({'more invisible header': stdevdata}).T
    df3 = pd.DataFrame({'Data': ['AVERAGE','STD DEV']})

    # Convert the dataframe to an XlsxWriter Excel object.
    df0.to_excel(writer, sheet_name='Sheet1', header = False, index=False)
    df1.to_excel(writer, sheet_name='Sheet1', startrow=index, startcol=1, header = False, index=False)
    df2.to_excel(writer, sheet_name='Sheet1', startrow=(index+1), startcol=1, header = False, index=False)
    df3.to_excel(writer, sheet_name='Sheet1', startrow=index, header = False, index=False)

    #Loop through silicon cell data
    index=index+2
    jsctot=[]
    voctot=[]
    pmaxtot=[]
    jmptot=[]
    vmptot=[]
    fftot=[]
    pcetot=[]
    while j < len(allsample):
        samplename=allsample[j][0]
        jsc=allsample[j][1]
        jsctot.append(jsc)
        voc=allsample[j][2]
        voctot.append(voc)
        pmax=allsample[j][3]
        pmaxtot.append(pmax)
        jmp=allsample[j][4]
        jmptot.append(jmp)
        vmp=allsample[j][5]
        vmptot.append(vmp)
        ff=allsample[j][6]
        fftot.append(ff)
        pce=allsample[j][7]
        pcetot.append(pce)
        individualdata=(samplename,jsc,voc,pmax,jmp,vmp,ff,pce)
        df = pd.DataFrame({'invisible header': individualdata}).T
        df.to_excel(writer, sheet_name='Sheet1', startrow=index, index=False,header =False)
        j=j+1
        index=index+1
        continue
    
    # Close the Pandas Excel writer and output the Excel file.
    writer.save()
    print('Solar Cell Data Summary exported ---->')
    continue
