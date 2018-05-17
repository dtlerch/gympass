# Combined program to account for hiring, firing, and employee transfers between companies
# Please contact Dylan Lerch dylan.lerch@gympass.com with questions, comments, suggestions, and especially with test cases
# that do not work. Thanks!

# PROCESS GUIDE:
# 1. Save the previous month's bases from the CRM as .xlsx. Be sure to change any numbers-as-text into numeric values instead.
# 2. Save the current month's upload from the H.R. Portal as .xlsx and check column headers. Required* case-sensitive headers:
#
#    EmpID
#    Name
#    Home Phone
#    Email Address
#
#    (EmpID refers to our Token. Name is formatted Last, First but this can be adjusted in the code.)
#    *Any other columns can be left or removed, and any columns left in can be present in any order.
#
# 3. Note the company names and company IDs for input values.
# 4. Run the Python script and follow the dialog prompts precisely. Be sure not to mix up the current/previous bases or IDs.
# 5. Review the QC() function results for inconsistencies
# 6. Review the operation_batch.xlsx file for errors
# 7. Upload the operation_batch.xlsx file, review the results, and submit.

import pandas as pd
import string
import re
from tabulate import tabulate
import datetime
import time
import numpy as np
from tkinter import filedialog
from tkinter import *
from sys import exit

month = datetime.date.today().strftime("%B")
name_list = []
id_list = []
dfold = pd.DataFrame()
dfnew = pd.DataFrame()

# Quality Check function to check the results at the end of the process

def QC():
    uniqueness = len(id_list)-len(set(id_list))
    if len(id_list) > len(set(id_list)):
        print("There are at least %d duplicated Company ID(s).\n" % (uniqueness))
    
    x_old = dfold.shape
    x_new = dfnew.shape
    num_hires = dfmaster[dfmaster["operation_type"]==1].shape
    num_fires = dfmaster[dfmaster["operation_type"]==200].shape
    num_trans = dfmaster[dfmaster["operation_type"]==130].shape
    
    if (x_old[0] + num_hires[0] - num_fires[0] - x_new[0]) == 0:
        print("The net movement numbers for %s in the %s family of companies check out:" % (month, name_list[0]))
        print("%d old employees, %d net hires, %d net fires, and %d current employees." % (x_old[0], num_hires[0], num_fires[0], x_new[0]))
        print("There were also %d net transfers." % (num_trans[0]))
    else:
        print("WARNING: The net movement numbers DO NOT check out.")
        print("Please double-check the base files and try again.")
    
    # Cute code to add a checkmark or an X to the Values Report below
    
    oldnu = x_old[0] - dfold["Full name"].nunique()
    newnu = x_new[0] - dfnew["Name"].nunique()
    oldtu = x_old[0] - dfold["Token"].nunique()
    newtu = x_new[0] - dfnew["EmpID"].nunique()
    
    checklist = [oldnu, newnu, oldtu, newtu]
        
    for i, s in enumerate(checklist):
        if s == 0:
            checklist[i] = u'\u2713'
        else:
            checklist[i] = "X"
    
    print("\nValues report:\n%s %d old names with %d unique. %s" % (checklist[0], x_old[0], dfold["Full name"].nunique(), checklist[0]))
    print("%s %d new names with %d unique. %s" % (checklist[1], x_new[0], dfnew["Name"].nunique(), checklist[1]))
    print("%s %d old tokens with %d unique. %s" % (checklist[2], x_old[0], dfold["Token"].nunique(), checklist[2]))
    print("%s %d new tokens with %d unique. %s" % (checklist[3], x_new[0], dfnew["EmpID"].nunique(), checklist[3]))
    
    if (x_old[0] != dfold["Token"].nunique()) | (x_new[0] != dfnew["EmpID"].nunique()):
        print("\nWARNING: There is a problem with Token uniqueness. Please double-check the base files and correct the problem.")
    
        if x_old[0] != dfold["Token"].nunique():
            print("The problem exists in the old base.")
            dupes1 = dfold[dfold["Token"].isin(dfold["Token"].value_counts()[dfold["Token"].value_counts()>1].index)].sort_values(by="Token")
            
            co_list = []
            for id in dupes1["Company"]:
                co_list.append(name_list[id])
            dupes1["Company"] = co_list           
            
            print(tabulate(dupes1[["Full name", "Token", "Company"]], headers='keys'))
            
        if x_new[0] != dfnew["EmpID"].nunique():
            print("The problem exists in the new base.")
            dupes2 = dfnew[dfnew["EmpID"].isin(dfnew["EmpID"].value_counts()[dfnew["EmpID"].value_counts()>1].index)].sort_values(by="EmpID")
            
            co_list = []
            for id in dupes2["Company"]:
                co_list.append(name_list[id])
            dupes2["Company"] = co_list
            
            print(tabulate(dupes2[["Full name", "Token", "Company"]], headers='keys'))
    
    if (x_old[0] != dfold["Full name"].nunique()) | (x_new[0] != dfnew["Name"].nunique()):
        print("\nThere are duplicated names in the base. Please ensure that they are all unique employees.")
        
        if x_old[0] != dfold["Full name"].nunique():
            print("\nPlease review the following repeated OLD names for accuracy:\n")
            dubs1 = dfold[dfold["Full name"].isin(dfold["Full name"].value_counts()[dfold["Full name"].value_counts()>1].index)].sort_values(by="Full name")
            
            co_list = []
            for id in dubs1["Company"]:
                co_list.append(name_list[id])
            dubs1["Company"] = co_list
            
            print(tabulate(dubs1[["Full name", "Token", "Company"]], headers='keys'))
        
        if x_new[0] != dfnew["Name"].nunique():
            print("\nPlease review the following repeated NEW names for accuracy:\n")
            dubs2 = dfnew[dfnew["Name"].isin(dfnew["Name"].value_counts()[dfnew["Name"].value_counts()>1].index)].sort_values(by="Name")
            
            co_list = []
            for id in dubs2["Company"]:
                co_list.append(name_list[id])
            dubs2["Company"] = co_list
 
    
            print(tabulate(dubs2[["Name", "EmpID", "Company"]], headers='keys'))

# User input for the number of companies in the current batch.
# Note that they must all be Parent-Dependent with each other and there is no feedback check to verify.
    
try:
    num = int(input("How many total IDs are associated with this company? Parent + daughters: "))
except ValueError:
    num = 0
    print("\nInvalid response. Please run the script again but submit a positive integer value instead.")

if num > 0:
    for x in range(1, num + 1):
        name = input("What is the name of company %d?: " % (x))
        name_list.append(name)
        try:
            id = int(input("What is the ID of %s?: " % (name)))
            id_list.append(id)
        except ValueError:
            print("\nInvalid response. Company IDs are numeric only. Please re-run and try again")
            break

    # User file selection dialog to load the files, two (one old, one new) for each named company
    # In theory, this could be replaced by just one upload each for old and new, if combined properly in advance
    # This also adds a column to the DataFrame to with an incremental value to denote each separate company
    # Will be changed at some point to write in the Company ID rather than a dummy variables
    
    root = Tk()

    i,k = 0,0

    for company in name_list:
        root.old =  filedialog.askopenfilename(initialdir = "/",title = "PREVIOUS month " + str(company) + " base",filetypes = (("spreadsheets","*.xlsx"),("all files","*.*")))
        old = pd.read_excel(root.old)
        old["Company"] = i
        dfold = dfold.append(old, ignore_index=True)
        i += 1

    for company in name_list:
        root.new =  filedialog.askopenfilename(initialdir = "/",title = "CURRENT month " + str(company) + " base",filetypes = (("spreadsheets","*.xlsx"),("all files","*.*")))
        new = pd.read_excel(root.new)
        new["Company"] = k
        dfnew = dfnew.append(new, ignore_index=True)
        k += 1

    root.withdraw()

    # Remove CRM-designated previously disabled users from our DataFrame to avoid incorrect duplicates

    dfold = dfold[dfold["Disabled at"].isnull()]

    consistent = pd.DataFrame()
    master_term = pd.DataFrame()
    master_hire = pd.DataFrame()
    master_tran = pd.DataFrame()

    # Loop through the companies and build a master DataFrame to analyze

    for x in range(num):
        terminated = pd.DataFrame(columns = ["company_id", "Token", "Full name"])
        hired = pd.DataFrame(columns = ["company_id", "EmpID", "Name", "Email Address", "Home Phone"])
        conlist = []
        conlist2 = []
        termlist = []
        hirelist = []
        namelist = []
        phonelist = []
        temp_old = dfold[dfold["Company"]==x].copy()
        temp_new = dfnew[dfnew["Company"]==x].copy()
        temp_old.loc[:,"Token"] = temp_old["Token"].astype(int)
        temp_new.loc[:,"EmpID"] = temp_new["EmpID"].astype(int)

        for tk in temp_old['Token']:
            if any(temp_new.EmpID == int(tk)):
                conlist.append(tk)
            else:
                termlist.append(tk)

        for id in temp_new['EmpID']:
            if id in conlist:
                conlist2.append(id)
            else:
                hirelist.append(id)

        for id in termlist:
            df = dfold[dfold["Token"]==id]
            terminated = terminated.append(df, ignore_index=True)

        terminated["company_id"] = id_list[x]

        for id in hirelist:
            df = dfnew[dfnew["EmpID"]==id]
            hired = hired.append(df,ignore_index=True)

        hired["company_id"] = id_list[x]
        terminated = terminated[["company_id", "Token", "Full name"]]
        terminated["email"] = np.nan
        terminated["phone"] = np.nan
        terminated["phone_country_code"] = np.nan
        terminated["send_email"] = 0
        terminated["operation_type"] = 200
        terminated.rename(columns={"Token": "token", "Full name": "name"}, inplace=True)
        master_term = master_term.append(terminated, ignore_index=True)

        for person in hired["Name"]:
            person = str(person)
            person = person.split(',')
            if len(person) > 1:
                person = person[1].strip() + " " + person[0].strip()
                namelist.append(person)
            else:
                person = person[0].strip()
                namelist.append(person)

        for phone in hired["Home Phone"]:
            phone = str(phone)
            phone = re.sub("[^0-9]", "", phone)
            phonelist.append(phone) 

        hired["name"] = namelist
        hired["phone"] = phonelist
        hired = hired[["company_id", "EmpID", "name", "Email Address", "phone"]]
        hired["phone_country_code"] = 1
        hired["send_email"] = 2
        hired["operation_type"] = 1
        hired.rename(columns={"EmpID": "token", "Email Address" : "email"}, inplace=True)
        master_hire = master_hire.append(hired, ignore_index=True)

    # Build the master DataFrame to export

    dfmaster = pd.DataFrame()
    dfmaster = dfmaster.append([master_hire, master_term], ignore_index=True)
    transfers = dfmaster[dfmaster["token"].isin(dfmaster["token"].value_counts()[dfmaster["token"].value_counts()>1].index)]
    dfmaster = dfmaster[dfmaster["token"].isin(dfmaster["token"].value_counts()[dfmaster["token"].value_counts()<=1].index)]

    # Account for transferred employees

    transfers["operation_type"].replace([200,1],[np.nan, 130], inplace=True)
    transfers = transfers[np.logical_not(transfers["operation_type"].isnull())]
    transfers["send_email"].replace([2],[0], inplace=True)

    dfmaster = dfmaster.append(transfers, ignore_index=True)

    dfmaster.operation_type = dfmaster.operation_type.astype(int)
    dfmaster.company_id = dfmaster.company_id.astype(int)
    dfmaster.send_email = dfmaster.send_email.astype(int)

    # Export to .csv in the same folder that the script is run from
    # Should be ready to upload directly to https://www.gympass.com/operation_batches/new

    dfmaster.to_csv('__MAINTENANCE__ ' + month + " " + name_list[0] + ' Operation_Batch File.csv', index=False)

    print("\nMovement for " + month + " is completed. Please check the results for quality and accuracy.\n")

    QC() # Quality Check function defined at the start
elif num == 0:
    print("Well then I guess your work here is done already, isn't it?")