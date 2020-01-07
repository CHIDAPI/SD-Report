import pandas as pd
from pandas import DataFrame
from pyxlsb import open_workbook as open_xlsb
import codecs

pd.set_option('display.max_columns',None)
pd.set_option('display.expand_frame_repr', False)
pd.set_option('max_colwidth', -1)

df_cols = ['Skills','Interval Start Time','Interval End Time','CSQ Name','Total','Avg Queue Time','Max Queue Time','Total','Avg Handle Time','Max Handle Time','Total','Avg Queue Time','Max QueueTime','Service Level']
#df_cols = ['Interval End Time','CSQ Name','Calls Presented','Calls Handled','Calls Abandoned','Service Level']
df = pd.read_excel('C:/temp/JJ/Daily CSQ Performance Summary.xlsx',names=df_cols,header=1)


for (index, row) in df.iterrows():
    if(pd.isnull(row[5]) and pd.isnull(row[6]) and pd.isnull(row[11]) and pd.isnull(row[12])):
        pass
    else:
        Date = row[1].strftime('%Y-%m-%d')
        CP_AQ = row[5].strftime('%H:%M:%S')
        CP_MQ = row[6].strftime('%H:%M:%S')
        CH_AH = row[8].strftime('%H:%M:%S')
        CH_MH = row[9].strftime('%H:%M:%S')
        CA_AQ = row[11].strftime('%H:%M:%S')
        CA_MQ = row[12].strftime('%H:%M:%S')
        with open("C:/Users/l_li/Desktop/PowerShell/SD Cisco Report _Jamie/Daily CSQ Performance Summary.txt", "a") as file:
            file.write("Date  " + Date + '\n')
            file.write("Avg Queue Time - Calls Presented  " + CP_AQ + '\n')
            file.write("Max Queue Time - Calls Presented  " + CP_MQ + '\n')
            file.write("Avg Handle Time  " + CH_AH + '\n')
            file.write("Max Handle Time  " + CH_MH + '\n')
            file.write("Avg Queue Time - Calls Abandoned  " + CA_AQ + '\n')
            file.write("Avg Queue Time - Calls Abandoned  " + CA_MQ + '\n')
        Out_Data = {'Date':[Date],'CP Avg Queue Time':[CP_AQ],'CP Max Queue Time':[CP_MQ],'Avg Handle Time':[CH_AH],'Max Handle Time':[CH_MH],'CA Avg Queue Time':[CA_AQ],'CA Max Queue Time':[CA_MQ]}
        df_new = DataFrame(Out_Data, columns = ['Date','CP Avg Queue Time','CP Max Queue Time','Avg Handle Time','Max Handle Time','CA Avg Queue Time','CA Max Queue Time'])
        df_new.to_csv('G:/systems/End User Support/Schedule/Sched_Tracker/Daily CSQ Performance Summary.csv', header=None,mode='a', index=False)
        #with codecs.open('C:/Users/l_li/Desktop/PowerShell/SD Cisco Report _Jamie/Daily CSQ Performance Summary.csv','a') as file2:
            #csv_out_file = csv.DicWrittter(file2)
         #   export_csv = df_new.to_csv(file2,index=False,header=False)