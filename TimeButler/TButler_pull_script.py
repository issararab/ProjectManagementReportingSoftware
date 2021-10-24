#-- ========================================================================
#-- Author      : <Issar Arab>
#-- Created     : <date: 29.01.2021>
#-- Description : <In Germany, companies make use of TimeButler to keep  
#--                track of their employees vacation. This script pulls 
#--                customized reports for all vacation/absences reported
#--                by employees in a given company that uses TimeButler.
#--                The administrator needs to activate API connections in 
#--                his account to generate the connection Token and start
#--                pulling data.>
#-- ========================================================================

import requests
from requests.auth import AuthBase
import pandas as pd
import os
import json
import threading
import itertools
import sys
import time

def animate():
    while True:
        for c in itertools.cycle(['|', '/', '-', '\\']):
            sys.stdout.write('\rPulling: ' + c)
            sys.stdout.flush()
            time.sleep(0.1)
            
def toDataFrame(body):
    body = str(body).split('\n')[:-1]#Remove the last empty token
    #exit(0)
    #body[0] = body[0][2:]
    body = [entry.replace(',','=*=') for entry in body]
    with open('temp_list.csv', 'w') as f:
        for line in [entry.replace(';',',') for entry in body]:
            f.write(line + "\n")
    result = pd.read_csv('temp_list.csv',encoding = "ISO-8859-1")
    result = result.replace(to_replace=r'=*=', value=',', regex=True)
    os.remove("temp_list.csv")
    return result
    
with open('TButler_Config.json') as config_file:
    config = json.load(config_file)
 
if __name__ == "__main__":
    URL = config['url']
    PARAMS = {'auth':config['token'], 'year':config['year']} 
    
    # Animation
    t = threading.Thread(target=animate)
    t.daemon = True
    t.start()
    
    # Users 
    users = requests.post(url = URL+'users', params = PARAMS).content
    users = toDataFrame(users.decode("utf8"))
    #print(users.columns)
    # holidayentitlement
    holidayentitlement = requests.post(url = URL+'holidayentitlement', params = PARAMS).content
    holidayentitlement = toDataFrame(holidayentitlement.decode("utf8"))
    #print(holidayentitlement)
    # Absences 
    absences = requests.post(url = URL+'absences', params = PARAMS).content
    absences = toDataFrame(absences.decode("utf8"))
    absences.to_excel("absences.xlsx", index= False)

    #workdays
    workdays = requests.post(url = URL+'workdays', params = PARAMS).content
    workdays = toDataFrame(workdays.decode("utf8"))
    workdays.to_excel("workdays.xlsx", index= False)
    
    
    #Process Abwesenheiten
    absences = users.join(absences.set_index('User ID'), on='User ID', how='inner',lsuffix='u')
    absences = absences[['First name','Last name', 'From','To','Half a day','Morning','Type','Extra vacation day','State','Workdays','Medical certificate (sick leave only)']]
    #print(absences)

    #Compute approved vacation days
    approved_absenses = absences[absences['State'].isin(['Done','Approved'])]
    approved_absenses_grouped = approved_absenses.groupby(['First name','Last name'])['Workdays'].sum().reset_index()
    #print(approved_absenses_grouped)

    #Generate urlaubkonto
    konto = users.join(holidayentitlement.set_index('User ID'), on='User ID', how='left')
    #print(konto.columns)
    konto = konto[['First name','Last name','Date of entry (dd/mm/yyyy)', 'User account locked', 'Vacation contingent','Remaining vacation','Extra vacation days','Additional vacation for severely challenged persons','Expired Vacation','Paid out vacation']]

    ##Write tables
    absences.to_excel(config['abwesenheiten_output_file'], index= False)
    konto.to_excel(config['konto_output_file'], index= False)
    print("\rDone      !")
    
