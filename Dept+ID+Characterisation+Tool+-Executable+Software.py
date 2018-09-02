
# coding: utf-8

# # Filter Member Tool

# In[29]:


# Our usual suspects for data manipulation
import pandas as pd
import numpy as np

# Import libraries for data visualisation
import matplotlib.pyplot as plt

# Import the Iris dataset from SciKit Learn
from sklearn import datasets


# In[47]:


import sys
x = input("Enter Product Code: ")
print('Your product code is' + x )


# # Import Sales Data

# In[30]:


#identifies dataset location
dataset_location = 'data.xlsx'
#imports dataset into Azure Notebook
dataFrame = pd.read_excel('data.xlsx', encoding='ISO-8859-1')

#Return data type
#print("Dataset class type: ")
#print(type(dataFrame))

#Return size of dataset
#print("Dataset size: ")
#print(dataFrame.shape)

#Print dataFrame
#dataFrame.head()
print ('Sales data uploaded')


# # Import Product Master File

# In[31]:


#identifies dataset location
dataset_location = 'CKC Prod Master (Full).xlsx'
#imports dataset into Azure Notebook
ProdMaster = pd.read_excel('CKC Prod Master (Full).xlsx', encoding='ISO-8859-1')

#Return data type
#print("Dataset class type: ")
#print(type(ProdMaster))

#Return size of dataset
#print("Dataset size: ")
#print(ProdMaster.shape)

#Print dataFrame
#ProdMaster.head()
print ('Product Master data uploaded')


# # Merge Data Sets

# In[32]:


dataFrame['Prod Code'] = dataFrame['Prod Code'].astype(str)
ProdMaster['Prod Code'] = ProdMaster['Prod Code'].apply(str)

mergeresult = pd.merge(dataFrame,
                 ProdMaster[['Prod Code', 'Dept ID', 'Cat ID', 'Prod Name (Chi)', 'Prod Name (Eng)', 'Subcat ID']],
                 on='Prod Code')

BusinessDate = mergeresult.pop('Business Date')
MonthID = mergeresult.pop('Month ID')
HourID = mergeresult.pop('Hour ID')
DayID = mergeresult.pop('Day ID')
MOPID = mergeresult.pop('MOP ID')
MemberGrade = mergeresult.pop('Member Grade')
MinID = mergeresult.pop('Min ID')
NetSales = mergeresult.pop('Net Sales')
ProdCode = mergeresult.pop('Prod Code')
SoldQty = mergeresult.pop('Sold Qty')
TranKey = mergeresult.pop('Tran Key')
SiteID = mergeresult.pop('Site ID')
mergeresult= mergeresult.sort_values(['Member ID'])
#mergeresult.head()


# # Visit Schedule

# In[33]:


dataFrame2 = dataFrame.copy()
#print('Number of Transactions: ')
#print(dataFrame2.shape)
#Remove duplicates (Same customer, on the same day, at the same hour)
dataFrame2.drop_duplicates(subset = ['Business Date','Hour ID', 'Member ID'], keep = 'first', inplace = True)
#print('Number of Unique Visits: ')
#print(dataFrame2.shape)


# In[34]:


BusinessDate = dataFrame2.pop('Business Date')
MonthID = dataFrame2.pop('Month ID')
DayID = dataFrame2.pop('Day ID')
MOPID = dataFrame2.pop('MOP ID')
MemberGrade = dataFrame2.pop('Member Grade')
MinID = dataFrame2.pop('Min ID')
NetSales = dataFrame2.pop('Net Sales')
ProdCode = dataFrame2.pop('Prod Code')
SoldQty = dataFrame2.pop('Sold Qty')
TranKey = dataFrame2.pop('Tran Key')

dataFrame3 = dataFrame2.copy()
SiteID = dataFrame2.pop('Site ID')
#dataFrame2.head()


# In[35]:


table= pd.pivot_table(dataFrame2, index = ['Member ID'], columns = ['Hour ID'], aggfunc=np.sum)
table.fillna(0, inplace = True)
table = pd.DataFrame(table.to_records())
table.columns = ['Member ID', '0', '1', '2', '3', '4', '5', '6', '7', '8', '9', '10', '11', '12', '13', '14', '15', '16', '17', '18', '19', '20', '21', '22', '23']
#table.head()


# In[36]:


table2 = table.copy()
table2['0000-0700'] = table2['1'] + table2['2'] + table2['3'] + table2['4'] + table2['5'] + table2['6'] + table2['7']
table2['0700-0900'] = table2['8'] + table2['9']
table2['0900-1200'] = table2['10'] + table2['11'] + table2['12']
table2['1200-1400'] = table2['13'] + table2['14']
table2['1400-1600'] = table2['15'] + table2['16']
table2['1600-1800'] = table2['17'] + table2['18']
table2['1800-2100'] = table2['19'] + table2['20'] + table2['21']
table2['2100-0000'] = table2['22'] + table2['23'] + table2['0']
#table2.head()


# In[37]:


table3 = table2.copy()
table3.drop(table3.iloc[:, 1:25], inplace=True, axis=1)
MemberID = table3.pop('Member ID')
table3['Unique Visits'] = table3.sum(axis=1)
table3['Member ID'] = MemberID
#table3.head()


# In[38]:


table4 = table3.copy()
MemberID = table4.pop('Member ID')
Sum = table4.pop('Unique Visits')

#Calculates top 2 time periods
arank = table4.apply(np.argsort, axis=1)
ranked_cols = table4.columns.to_series()[arank.values[:,::-1][:,:2]]
table5 = pd.DataFrame(ranked_cols, index=table4.index)
table5.columns = ['Top Visit Period', 'Second Top Visit Period']
table5['Second Top Visit Count'] = table4.apply(lambda row: row.nlargest(2).values[-1],axis=1)
table5['Top Visit Count'] = table4.max(axis=1)
table5['Member ID'] = MemberID
table5['Unique Visits'] = Sum
table5 = table5 [['Member ID','Top Visit Period','Top Visit Count','Second Top Visit Period','Second Top Visit Count', 'Unique Visits']]
#table5.head()


# In[39]:


table5.to_excel('Top2VisitPeriod.xlsx')
print('Top2VisitPeriod Export Complete')


# # Store Location Characterisation

# In[40]:


HourID = dataFrame3.pop('Hour ID')

locationtable= pd.pivot_table(dataFrame3, values = 'Number of Records', index = ['Member ID'], columns = ['Site ID'], aggfunc=np.sum)
locationtable.fillna(0, inplace = True)
locationtable = pd.DataFrame(locationtable.to_records())

#Calculates top 2 store locations
MemberID = locationtable.pop('Member ID')
arank = locationtable.apply(np.argsort, axis=1)
ranked_cols = locationtable.columns.to_series()[arank.values[:,::-1][:,:2]]
locationtable1 = pd.DataFrame(ranked_cols, index=locationtable.index)
locationtable1.columns = ['Top Visit Store', 'Second Top Visit Store']
locationtable1['Second Top Visit Store Count'] = locationtable.apply(lambda row: row.nlargest(2).values[-1],axis=1)
locationtable1['Top Visit Store Count'] = locationtable.max(axis=1)
locationtable1['Member ID'] = MemberID
locationtable1 = locationtable1 [['Member ID','Top Visit Store','Top Visit Store Count','Second Top Visit Store','Second Top Visit Store Count']]
#locationtable1.head()


# In[41]:


locationtable1.to_excel('Top2VisitLocation.xlsx')
print('Top2VisitLocation Export Complete')


# # Find Top Five Highest Dept ID

# In[42]:


ItemCheck1 = mergeresult.copy()

ItemCheck1= pd.pivot_table(ItemCheck1, values = 'Number of Records', index = ['Member ID'], columns = ['Dept ID'], aggfunc=np.sum)
ItemCheck1.fillna(0, inplace = True)
ItemCheck1 = pd.DataFrame(ItemCheck1.to_records())
#ItemCheck1.head()


# In[43]:


#Calculates top 2 Dept ID
MemberID = ItemCheck1.pop('Member ID')
arank = ItemCheck1.apply(np.argsort, axis=1)
ranked_cols = ItemCheck1.columns.to_series()[arank.values[:,::-1][:,:5]]
ItemCheck2 = pd.DataFrame(ranked_cols, index=ItemCheck1.index)
ItemCheck2.columns = ['Top Dept ID', 'Second Top Dept ID', 'Third Top Dept ID', 'Fourth Top Dept ID', 'Fifth Top Dept ID']
ItemCheck2['Top Dept ID Count'] = ItemCheck1.max(axis=1)
ItemCheck2['Second Top Dept ID Count'] = ItemCheck1.apply(lambda row: row.nlargest(2).values[-1],axis=1)
ItemCheck2['Third Top Dept ID Count'] = ItemCheck1.apply(lambda row: row.nlargest(3).values[-1],axis=1)
ItemCheck2['Fourth Top Dept ID Count'] = ItemCheck1.apply(lambda row: row.nlargest(4).values[-1],axis=1)
ItemCheck2['Fifth Top Dept ID Count'] = ItemCheck1.apply(lambda row: row.nlargest(5).values[-1],axis=1)
ItemCheck2['Member ID'] = MemberID
ItemCheck2 = ItemCheck2 [['Member ID','Top Dept ID','Top Dept ID Count','Second Top Dept ID','Second Top Dept ID Count','Third Top Dept ID','Third Top Dept ID Count', 'Fourth Top Dept ID','Fourth Top Dept ID Count', 'Fifth Top Dept ID','Fifth Top Dept ID Count']]
#MemberID = ItemCheck2.pop('Member ID')
#ItemCheck2.head()


# In[44]:


ItemCheck2.to_excel('Top5DeptID.xlsx')
print('Top5DeptID Export Complete')


# # Find Top Five Highest Cat ID

# In[45]:


ItemCheck3 = mergeresult.copy()

ItemCheck3= pd.pivot_table(ItemCheck3, values = 'Number of Records', index = ['Member ID'], columns = ['Cat ID'], aggfunc=np.sum)
ItemCheck3.fillna(0, inplace = True)
ItemCheck3 = pd.DataFrame(ItemCheck3.to_records())

#Calculates top 2 Cat ID
MemberID = ItemCheck3.pop('Member ID')
ItemCheck3['TotTrans'] =  ItemCheck3.sum(axis=1)
TotTrans = ItemCheck3.pop('TotTrans')
arank = ItemCheck3.apply(np.argsort, axis=1)
ranked_cols = ItemCheck3.columns.to_series()[arank.values[:,::-1][:,:5]]
ItemCheck4 = pd.DataFrame(ranked_cols, index=ItemCheck3.index)
ItemCheck4.columns = ['Top Cat ID', 'Second Top Cat ID', 'Third Top Cat ID', 'Fourth Top Cat ID', 'Fifth Top Cat ID']
ItemCheck4['Top Cat ID Count'] = ItemCheck3.max(axis=1)
ItemCheck4['Second Top Cat ID Count'] = ItemCheck3.apply(lambda row: row.nlargest(2).values[-1],axis=1)
ItemCheck4['Third Top Cat ID Count'] = ItemCheck3.apply(lambda row: row.nlargest(3).values[-1],axis=1)
ItemCheck4['Fourth Top Cat ID Count'] = ItemCheck3.apply(lambda row: row.nlargest(4).values[-1],axis=1)
ItemCheck4['Fifth Top Cat ID Count'] = ItemCheck3.apply(lambda row: row.nlargest(5).values[-1],axis=1)
ItemCheck4['Member ID'] = MemberID
ItemCheck4 = ItemCheck4 [['Member ID','Top Cat ID','Top Cat ID Count','Second Top Cat ID','Second Top Cat ID Count','Third Top Cat ID','Third Top Cat ID Count','Fourth Top Cat ID','Fourth Top Cat ID Count','Fifth Top Cat ID','Fifth Top Cat ID Count']]

ItemCheck4['TotTrans'] = TotTrans
#ItemCheck4['Unique Visits'] = Sum
#MemberID = ItemCheck4.pop('Member ID')
#ItemCheck4.head()


# In[46]:


ItemCheck4.to_excel('Top5CatID.xlsx')
print('Top5CatID Export Complete')


# # Filter Members who bought certain product

# In[48]:


#Input Product Code and apply Filter
filtertable = dataFrame[dataFrame['Prod Code'] == x]
#filtertable.head()


# In[49]:


frame = filtertable['Member ID'].value_counts().to_frame('Bought Count')
frame['Member ID'] = frame.index
frame.reset_index(inplace= True, drop = True)
frame['Bought Product'] = 1
frame = frame[['Member ID', 'Bought Product', 'Bought Count']]
frame = frame.sort_values(['Member ID'])
#frame.head()


# In[50]:


# Filtered List- Top Dept ID
filteredlist1 = pd.merge(ItemCheck2,
                 frame[['Member ID', 'Bought Product', 'Bought Count']],
                 on='Member ID')
#filteredlist1.head()


# In[51]:


filteredlist1.to_excel('FilteredTop5DeptID.xlsx')
print('FilteredTop5DeptID Export Complete')


# In[52]:


# Filtered List- Top Cat ID
filteredlist2 = pd.merge(ItemCheck4,
                 frame[['Member ID', 'Bought Product', 'Bought Count']],
                 on='Member ID')
#filteredlist2.head()


# In[53]:


filteredlist2.to_excel('FilteredTop5CatID.xlsx')
print('FilteredTop5CatID Export Complete')


# In[54]:


#print(filteredlist2.shape)

