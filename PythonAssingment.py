#!/usr/bin/env python
# coding: utf-8

# In[1]:


import pandas as pd

# Read the input sheets into pandas dataframes
xls = pd.ExcelFile(r"C:\Users\deepi\Downloads\dataset2.xlsx")
df1 = pd.read_excel(xls,'Sheet2')
df2 = pd.read_excel(xls,'Sheet3')

# merging both datasets
df3 =  pd.merge(df1,df2,left_on='User ID' , right_on = 'uid',how='left').drop('uid',axis=1)

df1 = df3.groupby('Team Name').agg({'total_statements': 'mean', 'total_reasons': 'mean'})
df1 = df1.rename(columns={'total_statements': 'Average Statements', 'total_reasons': 'Average Reasons'})
df1 = df1.sort_values('Average Statements', ascending=False).reset_index()
df1['Team Rank'] = df1.index + 1
df1 = df1.rename(columns={'Team Name': 'Thinking Teams Leaderboard'})
df1 = df1[['Team Rank' , 'Thinking Teams Leaderboard' , 'Average Statements' , 'Average Reasons']]
df1.to_excel('output1.xlsx', index=False)

# Sort the users by their number of statements and reasons
df2 = df2.sort_values(['total_statements', 'total_reasons'], ascending=False)
# Add a rank column based on the sorted order of the users
df2['Rank'] = range(1, len(df2) + 1)
# Reorder the columns to match the output sheet
df2 = df2[['Rank', 'name', 'uid', 'total_statements', 'total_reasons']]

# Rename the columns to match the output sheet
df2 = df2.rename(columns={'name': 'Name', 'uid': 'UID', 'total_statements': 'No. of Statements', 'total_reasons': 'No. of Reasons'})
df2.to_excel('output2.xlsx', index=False)


# In[2]:


df1


# In[3]:


df2


# In[6]:


df1.to_clipboard()


# In[7]:


df2.to_clipboard()


# In[ ]:




