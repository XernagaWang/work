#!/usr/bin/env python
# coding: utf-8

# In[86]:


import pandas as pd
import numpy as np


# In[87]:


vk_spec = pd.read_excel("auto_spec_db_trial_volkswagen (2).xlsx", sheet_name="V2")


# In[88]:


vk_spec1 = vk_spec[['OEM', 'Brand', 'Vehicle classification', 'Model family']]


# In[89]:


vk_spec1 = vk_spec1.drop_duplicates()


# In[90]:


vk_spec1 = vk_spec1.reset_index(drop=True)


# In[91]:


vk_spec1 = vk_spec1.sort_values(by=['Vehicle classification', 'Model family'])


# In[92]:


vk_spec1


# In[93]:


gp = list(vk_spec1.groupby(['Vehicle classification']))
row_list = []

for ind, row in gp:
    n = row.shape[1]
    v_line = pd.DataFrame([[np.NAN]*n], columns=row.columns)
    row_nan = pd.concat([row, v_line], axis=0)
    row_list.append(row_nan)
step2_vk_temp = pd.concat(row_list) 


# In[94]:


step2_vk_temp = step2_vk_temp.reset_index(drop=True)


# In[95]:


step2_vk_temp


# In[96]:


cols =["Australia", "India", "South East Asia", "Japan", "Rest of Asia Pacific", "South Korea", "People's Republic of China (mainland)", "Rest of Greater China", "Rest of Central & Eastern Europe",
                "Rest of Middle East", "France", "Germany", "UK", "Rest of Latin America", "Canada", "United States", "space1", "space2", "Australia-r", "India-r", "South East Asia-r", "Japan-r", "Rest of Asia Pacific-r", "South Korea-r", "People's Republic of China (mainland)-r", "Rest of Greater China-r", "Rest of Central & Eastern Europe-r",
                "Rest of Middle East-r", "France-r", "Germany-r", "UK-r", "Rest of Latin America-r", "Canada-r", "United States-r"]


# In[97]:


df2 = pd.DataFrame(columns=cols)


# In[98]:


df2


# In[99]:


step2_vk_temp = pd.concat([step2_vk_temp, df2], axis=1)


# In[100]:


step2_vk_temp


# In[101]:


# colA = ["OEM", "Brand", "Vehicle classification", "Model family", "Australia", "India", "South East Asia", "Japan", "Rest of Asia Pacific", "South Korea", "People's Republic of China (mainland)", "Rest of Greater China", "Rest of Central & Eastern Europe",
#                 "Rest of Middle East", "France", "Germany", "UK", "Rest of Latin America", "Canada", "United States", " ", " ", "Australia-r", "India-r", "South East Asia-r", "Japan-r", "Rest of Asia Pacific-r", "South Korea-r", "People's Republic of China (mainland)-r", "Rest of Greater China-r", "Rest of Central & Eastern Europe-r",
#                 "Rest of Middle East-r", "France-r", "Germany-r", "UK-r", "Rest of Latin America-r", "Canada-r", "United States-r"]


# In[102]:


all_row = ['ALL', 'ALL', 'ALL', 'ALL']
ALL_df = pd.DataFrame(data = [all_row], columns= ["OEM", "Brand", "Vehicle classification", "Model family"])

for c in cols:
    ALL_df[c] = 0

ALL_df


# In[103]:


nan_row = pd.DataFrame([[np.NAN]*38], columns = ALL_df.columns)


# In[104]:


gp = list(step2_vk_temp.groupby(['Vehicle classification']))
row_list = []

for ind, row in gp:
    for c in cols:
        ALL_df[c] = 0
        res = pd.concat([ALL_df, row, nan_row], axis=0)
    row_list.append(res)


# In[105]:


fin = pd.concat(row_list).reset_index(drop=True)


# In[106]:


fin


# In[108]:


get_ipython().system('pip install xlsxwriter')


# In[117]:


from pandas import ExcelWriter, read_excel
 
with pd.ExcelWriter('1234.xlsx', engine='xlsxwriter') as writer:
    fin.to_excel(writer, sheet_name='Sheet1')
    formatObj = writer.book.add_format({'num_format': '0.00%'})
    writer.book.sheetnames['Sheet1'].set_column('X:AN', cell_format=formatObj)
    workbook = writer.book
    worksheet = writer.sheets['Sheet1']
    
    header_no_color = workbook.add_format(
    {
        "bold": True,
        "text_wrap": True,
        "valign": "top",
        "fg_color": "#FFFFFF",
        "border": 1,
    }
)
    colors = ['#CD5C5C', '#F08080', '#FA8072', '#E9967A', 
          '#FFA07A', '#DC143C', '#FF0000', '#B22222', 
          '#8B0000', '#3CB371', '#2E8B57', '#228B22', 
          '#008000', '#006400', '#9ACD32', '#6B8E23',
          '#808000', '#556B2F', '#66CDAA', '#8FBC8B',
          '#20B2AA', '#008B8B', '#008080', '#ADD8E6', '#87CEEB',
          '#87CEFA', '#00BFFF', '#1E90FF', '#6495ED', '#7B68EE',
          '#4169E1', '#0000FF', '#E6E6FA', '#D8BFD8']
    for col_num, value in enumerate(fin.columns.values):
    
        if value in fin.columns.values[4:]:

            # print(col_num)    
            header_format = workbook.add_format(
                {
                 "bold": True,
                 "text_wrap": True,
                 "valign": "top",
                 "fg_color": colors[col_num - 10],
                 "border": 1,
                 }
                )

            worksheet.write(0, col_num + 1, value, header_format)

        else:
            worksheet.write(0, col_num + 1, value, header_no_color)

    
writer.close()


# In[ ]:





# In[121]:


countries =["Australia", "India", "South East Asia", "Japan", "Rest of Asia Pacific", "South Korea", "People's Republic of China (mainland)", "Rest of Greater China", "Rest of Central & Eastern Europe",
                "Rest of Middle East", "France", "Germany", "UK", "Rest of Latin America", "Canada", "United States"]


# In[122]:


type = ['ICE', 'BEV', 'PHEV', 'Others']


# In[118]:


# 创建MultiIndex数据帧
cols1 = [['组1', '组1', '组2', '组2'], ['A', 'B', 'A', 'B']]
data = [[1, 2, 3, 4], [5, 6, 7, 8]]
index = ['第1行', '第2行']
df = pd.DataFrame(data, index=index, columns=cols1)

# 导出到Excel
df.to_excel('multi_header.xlsx', engine='openpyxl')


# In[161]:


a = [(con, con, con, con) for con in countries]


# In[162]:


a


# In[164]:


b = [j for i in a for j in i]


# In[169]:


c = [n for m in range(len(countries)) for n in type]


# In[165]:


b


# In[170]:


c


# In[171]:


df1 = pd.DataFrame(columns=[b,c])


# In[172]:


df1


# In[187]:


df1.to_excel("df1_test2.xlsx", engine='openpyxl')


# In[182]:


to_see


# In[184]:


too_see = pd.concat([vk_spec1, df1])


# In[185]:


to_see.columns


# In[186]:


vk_spec1.to_excel("vk_spec1.xlsx", index=False)


# In[267]:


df1 = pd.read_excel("df1_test1.xlsx", header=[0,1])
df2 = pd.read_excel("vk_spec1.xlsx", header=[0,1])


# In[268]:


df2.head()


# In[269]:


fin = pd.concat([df2, df1.drop([('Unnamed: 0_level_0', 'Unnamed: 0_level_1')], axis = 1)], axis = 1)
fin


# In[270]:


# fin.loc[0] = {
#   fin.columns[0] : fin.columns[0][1],
#   fin.columns[1] : fin.columns[1][1],
#   fin.columns[2] : fin.columns[2][1],
#   fin.columns[3] : fin.columns[3][1],
# }


# In[224]:


# fin


# In[218]:


# fin.columns = pd.MultiIndex.from_tuples(fin.set_axis(fin.columns.values, axis=1).rename(columns={
#    (fin.columns[0][0], fin.columns[0][1]): (fin.columns[0][0], ''),
#    (fin.columns[1][0], fin.columns[1][1]): (fin.columns[1][0], ''),
#    (fin.columns[2][0], fin.columns[2][1]): (fin.columns[2][0], ''),
#    (fin.columns[3][0], fin.columns[3][1]): (fin.columns[3][0], ''),

#   }))


# In[219]:


# fin.columns[1][1]


# In[271]:


l1 = [fin.columns[0][1], fin.columns[1][1], fin.columns[2][1], fin.columns[3][1]]

for _ in range(64):
  l1.append(np.NAN)
  
df = pd.DataFrame(np.insert(fin.values, 0, values= l1, axis=0))
df.head()


# In[272]:


fin.columns = pd.MultiIndex.from_tuples(fin.set_axis(fin.columns.values, axis=1).rename(columns={
   (fin.columns[0][0], fin.columns[0][1]): (fin.columns[0][0], ''),
   (fin.columns[1][0], fin.columns[1][1]): (fin.columns[1][0], ''),
   (fin.columns[2][0], fin.columns[2][1]): (fin.columns[2][0], ''),
   (fin.columns[3][0], fin.columns[3][1]): (fin.columns[3][0], ''),

  }))


# In[273]:


df.columns = fin.columns
df.head()


# In[231]:


df.to_excel("df_sample.xlsx")


# In[238]:


df.columns.levels[0]


# In[274]:


all_row2 = ['ALL', 'ALL', 'ALL', 'ALL', 'ICE', 'BEV', 'PHEV', 'Others', 'ICE', 'BEV', 'PHEV', 'Others', 'ICE', 'BEV', 'PHEV', 'Others', 'ICE', 'BEV', 'PHEV', 'Others', 'ICE', 'BEV', 'PHEV', 'Others', 'ICE', 'BEV', 'PHEV', 'Others', 'ICE', 'BEV', 'PHEV', 'Others', 'ICE', 'BEV', 'PHEV', 'Others', 'ICE', 'BEV', 'PHEV', 'Others', 'ICE', 'BEV', 'PHEV', 'Others', 'ICE', 'BEV', 'PHEV', 'Others', 'ICE', 'BEV', 'PHEV', 'Others', 'ICE', 'BEV', 'PHEV', 'Others', 'ICE', 'BEV', 'PHEV', 'Others', 'ICE', 'BEV', 'PHEV', 'Others', 'ICE', 'BEV', 'PHEV', 'Others']
ALL_df2 = pd.DataFrame(data = [all_row2], columns= df.columns)

# for c in countries:
#     for t in ['ICE', 'BEV', 'PHEV', 'Others']:
#         ALL_df2[c] = c
#         ALL_df[t] = t

ALL_df2


# In[275]:


nan_row2 = pd.DataFrame([[np.NAN]*68], columns = ALL_df2.columns)


# In[276]:


df.columns[2]


# In[279]:


gp4 = list(df.groupby(df.columns[2]))
row_list4 = []

for ind, row in gp4:
    print(row)
    res4 = pd.concat([ALL_df2, row, nan_row2], axis=0)
    row_list4.append(res4)


# In[280]:


fin4 = pd.concat(row_list4).reset_index(drop=True)


# In[281]:


fin4


# In[282]:


# fin4.to_excel("fin4_to_see.xlsx")


# In[287]:


fin4 = fin4.rename(columns={fin4.columns[0][1]:'ALL'})


# In[288]:


fin4


# In[290]:


fin4.shape


# In[295]:


with pd.ExcelWriter('12345.xlsx', engine='xlsxwriter') as writer:
    fin4.to_excel(writer, sheet_name='Sheet1')
    formatObje = writer.book.add_format({'num_format': '0.00%'})
    writer.book.sheetnames['Sheet1'].set_column('F:BQ', cell_format=formatObje)
    workbook = writer.book
    worksheet = writer.sheets['Sheet1']
    
    worksheet.write_array_formula(4,5,len(fin4)+1, 5, '{=1-SUM(G:I)}')
  
    
    
    
    
    
    
    
    
writer.close()
    
#     header_no_color = workbook.add_format(
#     {
#         "bold": True,
#         "text_wrap": True,
#         "valign": "top",
#         "fg_color": "#FFFFFF",
#         "border": 1,
#     }
# )
#     colors = ['#CD5C5C', '#F08080', '#FA8072', '#E9967A', 
#           '#FFA07A', '#DC143C', '#FF0000', '#B22222', 
#           '#8B0000', '#3CB371', '#2E8B57', '#228B22', 
#           '#008000', '#006400', '#9ACD32', '#6B8E23',
#           '#808000', '#556B2F', '#66CDAA', '#8FBC8B',
#           '#20B2AA', '#008B8B', '#008080', '#ADD8E6', '#87CEEB',
#           '#87CEFA', '#00BFFF', '#1E90FF', '#6495ED', '#7B68EE',
#           '#4169E1', '#0000FF', '#E6E6FA', '#D8BFD8']
#     for col_num, value in enumerate(fin.columns.values):
    
#         if value in fin.columns.values[4:]:

#             # print(col_num)    
#             header_format = workbook.add_format(
#                 {
#                  "bold": True,
#                  "text_wrap": True,
#                  "valign": "top",
#                  "fg_color": colors[col_num - 10],
#                  "border": 1,
#                  }
#                 )

#             worksheet.write(0, col_num + 1, value, header_format)

#         else:
#             worksheet.write(0, col_num + 1, value, header_no_color)

    
# writer.close()


# In[ ]:




