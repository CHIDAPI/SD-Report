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


df_cols2 = ['CSQ Name','CSQ ID','Call Skills','Interval Start Time','Interval End Time','Service Level (sec)','Calls Handled < Service Level','Calls Abandoned < Service Level','Only Handled','With No Abandoned Calls','With Abandoned Calls Counted Positively','With Abandoned Calls Counted Negatively','Calls Presented','Handled','%','Abandoned','%','Dequeued','%']
#df_cols = ['Interval End Time','CSQ Name','Calls Presented','Calls Handled','Calls Abandoned','Service Level']
df2 = pd.read_excel('C:/temp/JJ/Oxford SD - Daily CSQ Report.xlsx',names=df_cols2,header=1)
#print(df.dtypes)
#for i in range(0, len(df)):
#    print (df.iloc[i,5])
#print(df2)
print(df2.dtypes)

print ("========================================")
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
