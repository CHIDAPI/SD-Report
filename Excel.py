import pandas as pd
import numpy as np
#import matplotlib.pyplot as plt

pd.set_option('display.max_columns',None)
pd.set_option('display.expand_frame_repr', False)
pd.set_option('max_colwidth', -1)

df_cols = ['Agent Name','Agent_ID','Extension','LBLT','Login_Time','LOALT','LogOut_Time','Logout Reason','Logged_In_Duration']
df = pd.read_excel('C:/temp/JJ/Oxford_SD_Login_logout_LFW.xlsx',names=df_cols,header=1)

df['Login_Time'] = pd.to_datetime(df.Login_Time)
df['LogOut_Time'] = pd.to_datetime(df.LogOut_Time)
df['Duration'] = (df.LogOut_Time - df.Login_Time)
#print(df.dtypes)

Agent=df['Agent_ID'].drop_duplicates()
Agent=Agent.dropna()
#print(Agent)
X=Agent.values


def TimeCheck( WhoIsThis ):
    Who = ((df.loc[df.Agent_ID == WhoIsThis])[['Login_Time','LogOut_Time','Duration','Agent_ID']])
    for (index, row) in Who.iterrows():
        #In = (row[0].date()).strftime('%Y-%m-%d')
        #Out = (row[1].date()).strftime('%Y-%m-%d')    ########
        if (pd.isnull(row[1])):
            Who.loc[index, 'LogOut_Time'] = pd.to_datetime(((row[0].date()).strftime('%Y-%m-%d')) + " " + "23:59:59")

        if (pd.isnull(row[0])):     ################
            Who.loc[index,'Login_Time'] = pd.to_datetime(((row[1].date()).strftime('%Y-%m-%d')) + " " + "00:00:01") ###############


    ####### Check for duplicated Agent ID #########
    if len(Who) > 1:
        for k in range(1,len(Who)):
            Who.iloc[k,0] = Who.iloc[k-1,0]
            Who.iloc[k-1,1] = Who.iloc[k,1]

    Who = Who.drop_duplicates(subset='Login_Time', keep='last')

    for (index, row) in Who.iterrows():
        In = (row[0].date()).strftime('%Y-%m-%d')
        Out = (row[1].date()).strftime('%Y-%m-%d')
        Out1 = (row[1].date())
        Out2 = (row[1])
        if In != Out:
            Who.loc[index, 'R_LogOut_Time'] = (In + " " + "23:59:59")
            Who['R_LogOut_Time'] = pd.to_datetime(Who.R_LogOut_Time)
            insert = pd.Series({'Login_Time': Out, 'LogOut_Time': Out2, 'R_LogOut_Time': Out2}, name=index * 2)
            Who = Who.append(insert)
        else:
            Who.loc[index, 'R_LogOut_Time'] = (row[1])
    Who = Who.drop(columns='LogOut_Time')
    Who['Duration'] = (Who.R_LogOut_Time - Who.Login_Time)

    for (index, row) in Who.iterrows():
        In = (row[0].date())

        In_Time = (row[0].time())
        Who.loc[index, 'Login_Date'] = (In)
    Who = Who[['Duration', 'Login_Date', 'Login_Time', 'R_LogOut_Time','Agent_ID']]
    Who = Who.sort_values(by='Login_Time')


    for i in range(1, len(Who)):

        if (Who.iloc[i, 1] == (Who.iloc[i - 1, 1])):
            Who.iloc[i, 0] += Who.iloc[i - 1, 0]

            while (Who.iloc[i, 2] != (Who.iloc[i - 1, 2]) or Who.iloc[i - 1, 3] != Who.iloc[i, 3] or Who.iloc[
                i - 1, 0] != Who.iloc[i, 0]):
                if (Who.iloc[i, 2] > (Who.iloc[i - 1, 2])):
                    Who.iloc[i, 2] = Who.iloc[i - 1, 2]
                if (Who.iloc[i, 3] > (Who.iloc[i - 1, 3])):
                    Who.iloc[i - 1, 3] = Who.iloc[i, 3]
                if (Who.iloc[i, 0]) > (Who.iloc[i - 1, 0]):
                    Who.iloc[i - 1, 0] = Who.iloc[i, 0]
    Who = Who.drop_duplicates(subset='Login_Date', keep='last')
    #with open("C:/temp/JJ/In_Out.txt", "a") as file:
        #file.write("==================================" + "\n")
        #file.write(WhoIsThis + "\n")



    for (index,row) in Who.iterrows():

        Date=(row[1]).strftime('%Y-%m-%d')
        In=(row[2].time()).strftime('%H:%M')
        Out=(row[3].time()).strftime('%H:%M')

        with open("G:/systems/End User Support/Schedule/Sched_Tracker/In_Out.txt", "a") as file:
        #with open("C:/temp/JJ/In_Out.txt", "a") as file:
            file.write(WhoIsThis + '\n')
            file.write(Date + '\n')
            file.write(In + '\n')
            file.write(Out + '\n')
    print(Who)
    #print(Who.dtypes())

for x in X:

    TimeCheck(x)