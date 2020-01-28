#!/usr/bin/env python
# coding: utf-8

# In[1]:


import pandas as pd
import numpy as np
import datetime
from datetime import date
from datetime import datetime
import re
import os
from dateutil.relativedelta import relativedelta


# In[2]:


file_path = "YOUR PATH"
file_name = "sender_details.xlsx"
output_file_name = "read_email_dets"
file_dtls = file_path+file_name


# In[3]:


if os.path.exists(file_dtls):
    print("file exists")
    pdf_file_dtls = pd.read_excel(file_dtls)
else:
    print("file not there...please provide file")
    raise SystemExit("Exiting as program can not move forward without file")
print("location of file is {}".format(file_dtls))


# In[4]:


Retailer = pdf_file_dtls["Retailer"].iloc[0]
Email =  pdf_file_dtls["Email"].iloc[0]
StartDate = pdf_file_dtls["StartDate"].iloc[0].date()


# In[5]:


v2_StartDate = StartDate
columns = ['Retailer','Email','StartDate','EndDate','Status','NoofEmails','JobStart','JobEnd']
df_retailer_run_dts = pd.DataFrame()
v3_todayDate = datetime.today().date()


# In[6]:


while v2_StartDate < v3_todayDate:
    v3_StartDate = v2_StartDate
    v2_StartDate = v2_StartDate + relativedelta(years=1)
    print("current v2_StartDate is {}  and v3_startDate is {}".format(v2_StartDate,v3_StartDate))
    if v2_StartDate > v3_todayDate:
        v2_StartDate = v3_todayDate
    df = pd.DataFrame([[Retailer,Email,v3_StartDate,v2_StartDate,'pending',0,'','']])
    df.columns = ['Retailer','Email','StartDate','EndDate','Status','NoofEmails','JobStart','JobEnd']
    df_retailer_run_dts = df_retailer_run_dts.append(df,ignore_index=True)

print("final v2_startdate is {}".format(v3_todayDate))


# In[7]:


df_retailer_run_dts.to_excel(file_path + output_file_name + "_" + Retailer + ".xlsx",index=False)


# In[ ]:




