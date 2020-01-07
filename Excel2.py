import pandas as pd
import datetime
import numpy as np
import matplotlib.pyplot as plt

pd.set_option('display.max_columns',None)
pd.set_option('display.expand_frame_repr', False)
pd.set_option('max_colwidth', -1)

df_cols = ['Agent_Name','Agent_ID','Agent_Extension','Interval_Start_Time','Interval_End_Time','Total_Logged_In','Reason_Code','Duration']
df = pd.read_excel('C:/temp/JJ/Oxford_SD_Agent_State_Summary_LFW.xlsx',names=df_cols,header=1)
df = df.drop(columns=['Agent_Name','Agent_Extension','Total_Logged_In']).dropna(subset=['Reason_Code']).fillna(method='ffill')

df=df[df['Reason_Code'].isin([16,15])]

df['Duration'] = df['Duration'].astype(str)
df['Duration'] = df['Duration'].apply(pd.Timedelta)



print(df.dtypes)

Agent = df['Agent_ID'].drop_duplicates()
#print(Agent)


def Calculation(WhoIsThis):

    Who = ((df.loc[df.Agent_ID == WhoIsThis])[['Interval_Start_Time','Interval_End_Time','Reason_Code','Duration']])
    #print(Who)


    for i in range(1,len(Who)):
        if len(Who) > 1:
            Who.iloc[i,3] = Who.iloc[i,3] + Who.iloc[i-1,3]
        #print(Who.iloc[i, 3])
            Who=Who.drop_duplicates(subset='Interval_Start_Time',keep='last')

        #print(WhoIsThis)
        #print(Who)

    for (index,row) in Who.iterrows():
        Start=(row[0]).strftime('%Y-%m-%d')
        End=(row[1]).strftime('%Y-%m-%d')
        Duration=(str(row[3])).split(' ')[2]
        #Duration=pd.to_datetime(row[3]).strftime('%H:%M:%S')
    with open("G:/systems/End User Support/Schedule/Sched_Tracker/BreakTime.txt", "a") as file:
        file.write(WhoIsThis + '\n')
        file.write(Start + '\n')
        file.write(End + '\n')
        file.write(Duration + '\n')


    #print(Who)
#print(df)
#print(Agent)


X=Agent.values
print(X)

for x in X:
    Calculation(x)