import os
import glob
#Find relevant CSVs in folder
path = os.getcwd()
extension = 'csv'
os.chdir(path)
csvresult = [i for i in glob.glob('*.{}'.format(extension))]
print("Tabulating Solar Cell Data from CSVs...")

import numpy as np
import matplotlib.pyplot as plt
import pandas as pd
import statistics

#iterate through CSVs
k=0
allsample=[]
while k < len(csvresult):
    eachsample=[]
    solar_volt=[]
    solar_curr=[]
    solar_pow=[]
    path = os.getcwd()
    path = path + "/" + csvresult[k]
    samplename=csvresult[k][:-4]
    eachsample.append(samplename)
    solar=pd.read_csv(path, delimiter=",",usecols=["Solar Cell Voltage","Solar Cell Current","Solar Cell Power"])
    i=0
    voltcount=0
    currcount=0
    while i < len(solar):
        solar_volt.append(solar.values[i][0])
        if solar.values[i][0] < 0:
            voltcount=voltcount+1
        solar_curr.append(solar.values[i][1])
        if solar.values[i][1] < 0:
            currcount=currcount+1
        solar_pow.append(solar.values[i][2])
        i=i+1
        continue
    jsc=(solar_curr[voltcount-1]+solar_curr[voltcount])/2
    voc=(solar_volt[currcount-1]+solar_volt[currcount])/2
    pmax=min(solar_pow)
    jmp=solar_curr[solar_pow.index(pmax)]
    vmp=solar_volt[solar_pow.index(pmax)]
    ff=abs(pmax/(voc*jsc))
    pce=abs((pmax*1000)/0.0572)
    eachsample.extend((jsc,voc,pmax,jmp,vmp,ff,pce))
    allsample.append(eachsample)
    k=k+1
    continue
   
#----------------------------------------------- EXCEL EXPORT-----------------------------------------------------

# Create a Pandas dataframe title for data.
df0 = pd.DataFrame({'Data': ['Sample','Jsc (A)','Voc (V)', 'Pmax (W)','Jmp (A)','Vmp (V)','FF','PCE (%)']}).T

writer = pd.ExcelWriter('Solar Cell Data Summary.xlsx', engine='xlsxwriter')

#Loop through data from total export list
j=0
index=1
#Loop through data
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

jscavg=sum(jsctot)/len(allsample)
jscdev=statistics.stdev(jsctot)
vocavg=sum(voctot)/len(allsample)
vocdev=statistics.stdev(voctot)
pmaxavg=sum(pmaxtot)/len(allsample)
pmaxdev=statistics.stdev(pmaxtot)
jmpavg=sum(jmptot)/len(allsample)
jmpdev=statistics.stdev(jmptot)
vmpavg=sum(vmptot)/len(allsample)
vmpdev=statistics.stdev(vmptot)
ffavg=sum(fftot)/len(allsample)
ffdev=statistics.stdev(fftot)
pceavg=sum(pcetot)/len(allsample)
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

# Close the Pandas Excel writer and output the Excel file.
writer.save()
print('Solar Cell Data Summary exported ---->')
