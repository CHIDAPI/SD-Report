import pandas as pd
from pandas import DataFrame
from pyxlsb import open_workbook as open_xlsb
import codecs

pd.set_option('display.max_columns',None)
pd.set_option('display.expand_frame_repr', False)
pd.set_option('max_colwidth', -1)

df_cols2 = ['CSQ Name','CSQ ID','Call Skills','Interval Start Time','Interval End Time','Service Level (sec)','Calls Handled < Service Level','Calls Abandoned < Service Level','Only Handled','With No Abandoned Calls','With Abandoned Calls Counted Positively','With Abandoned Calls Counted Negatively','Calls Presented','Handled','%','Abandoned','%','Dequeued','%']
#df_cols = ['Interval End Time','CSQ Name','Calls Presented','Calls Handled','Calls Abandoned','Service Level']
df2 = pd.read_excel('C:/temp/JJ/Oxford SD - Daily CSQ Report.xlsx',names=df_cols2,header=1)

for (index,row) in df2.iterrows():
    if (pd.notnull(row[4])):
        global Date2
        Date2 = row[4].strftime('%Y-%m-%d')
        print(Date2)

    if(pd.isnull(row[5]) and pd.notnull(row[6])):
        CH=str(row[6])
        CA=str(row[7])
        CP=str(row[12])
        H=str(row[13])
        A=str(row[15])



        Out_Data2 = {'Date': [Date2], 'Calls Handled < Service Level': [CH], 'Calls Abandoned < Service Level': [CA],
                    'Calls Presented': [CP], 'Calls Handled': [H], 'Calls Abandoned': [A]
                    }
        df_new2 = DataFrame(Out_Data2, columns=['Date', 'Calls Handled < Service Level', 'Calls Abandoned < Service Level', 'Calls Presented',
                                              'Calls Handled', 'Calls Abandoned' ])
        df_new2.to_csv('G:/systems/End User Support/Schedule/Sched_Tracker/Daily CSQ Report.csv',header=None,mode='a',index=False)
        #with codecs.open('C:/Users/l_li/Desktop/PowerShell/SD Cisco Report _Jamie/Daily CSQ Report.csv','a') as file3:
         #   export_csv = df_new2.to_csv(file3,index=False,header=False)

        with open("C:/Users/l_li/Desktop/PowerShell/SD Cisco Report _Jamie/Daily CSQ Report.txt", "a") as file:
            file.write("Calls Handled  " + CH + '\n')
            file.write("Calls Abandoned  " + CA + '\n')
            file.write("Calls Presented  " + CP + '\n')
            file.write("Handled  " + H + '\n')
            file.write("Abandoned  " + A + '\n')