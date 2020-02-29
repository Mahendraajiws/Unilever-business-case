#!/usr/bin/env python
# coding: utf-8

# In[1]:


#Import modules
import pandas as pd
from datetime import datetime

#read data excel by sheet
Shipmentreport1=pd.read_excel(r'E:\Mahendra\INTERNSHIP\Logistic Data Analytics.xlsx', sheet_name=0)
Shipmentreport2=pd.read_excel(r'E:\Mahendra\INTERNSHIP\Logistic Data Analytics.xlsx', sheet_name=1)
MaterialMD=pd.read_excel(r'E:\Mahendra\INTERNSHIP\Logistic Data Analytics.xlsx', sheet_name=2)
PlantMD=pd.read_excel(r'E:\Mahendra\INTERNSHIP\Logistic Data Analytics.xlsx', sheet_name=3)


# In[2]:


#Pre processing
#drop unused row in shipment report 2
x=636468
y=1048555
Shipmentreport2.drop(Shipmentreport2.loc[x:y].index, inplace=True)
Shipmentreport2


# In[3]:


#Join report shipment 1 & 2
Shipmentreport=pd.concat([Shipmentreport1,Shipmentreport2])
Shipmentreport


# In[4]:


#Drop missing values in shipment report
Shipmentreport=Shipmentreport.dropna(how='any', axis=0)
Shipmentreport


# In[5]:


#rename columns in shipment report
Shipmentreport4=Shipmentreport.rename(columns={Shipmentreport.columns[1]: "Qty",Shipmentreport.columns[6]: "PLANT CODE" })
Shipmentreport4


# In[43]:


#Add new columns for define month
Shipmentreport4['Month']=Shipmentreport4.Date.dt.month
Shipmentreport4


# In[13]:


#Aggregate shipment report by referencing material and plant unique id
Referencing1=Shipmentreport4.merge(MaterialMD, on='Material', how='left')
Aggregate=Referencing1.merge(PlantMD, on='PLANT CODE')
Aggregate


# In[44]:


#Rename month
Aggregate['Month']= Aggregate['Month'].replace({1: 'January',2: 'February',3: 'March',4: 'April',5: 'May',6: 'June',7: 'July',8: 'August',9: 'September',10: 'October',11: 'November',12: 'December'})
Aggregate


# In[56]:


#import modul
import numpy as py

#Make pivot table to aggregate data
Pivot=pd.pivot_table(Aggregate,values='Qty',index=['PLANT CODE','Category'],columns=['Month'],aggfunc=py.sum,margins=True,margins_name='Totals')
Pivot


# In[58]:


#Write pivot table to excel
Pivot.to_excel("E:\Mahendra\INTERNSHIP\Aggregate-Category Plant.xlsx") 


# In[ ]:




