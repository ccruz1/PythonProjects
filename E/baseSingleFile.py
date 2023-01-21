
# importing pandas as pd
from datetime import datetime 
import pandas as pd
import openpyxl 
from openpyxl import load_workbook
import os
  
start_time = datetime.now() 

##List files per year
##Change the year as needed
path_Default = "hubcopy/"
list_Files = os.listdir(path_Default)

##List here the data frames
list_Of_DFrames = list()

##Create the sheets in excel file
for files in list_Files:
    print("Escribiendo...")
    df = pd.DataFrame(pd.read_excel(path_Default+files))
    df = df[(df['sector_economico_4'].astype('str').str.contains('8701')) & ((df['cve_municipio']=='A01') | (df['cve_municipio']=='A02'))]
    list_Of_DFrames.append(df)

##Convert to data frames to excel
with pd.ExcelWriter('testing1.xlsx') as writer:
    ##DECLARATION OF THE WRITER
    list_Of_DFrames[0].to_excel(writer, engine="xlsxwriter", startrow=0)
    ##VAR WITH THE CURRENT WORKING ROW
    startrow = list_Of_DFrames[0].shape[0] + 1
    ##CYCLE THE DF's AND ADDING TO SPACES TO THE CURRENT ROW 
    x = 0
    for dframe in list_Of_DFrames[1:]:
        dframe.to_excel(writer, engine="xlsxwriter", startrow=startrow, header=False)
        startrow += (dframe.shape[0] + 2)
        x += 1
        print("Archivos recorridos: " + str(x))

print(len(list_Of_DFrames))

##Gets End Script Time
end_time = datetime.now()
print('Duration: {}'.format(end_time - start_time))





