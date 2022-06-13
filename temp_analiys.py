#!/usr/bin/env python
# coding: utf-8

# https://comrade-xl.ru/2021/08/05/py-combine-xl-sheets/

# In[1]:


#!pip install xlrd

#!pip install --py luxwidget
#!pip enable --py luxwidget


# In[2]:


import pandas as pd
import xlwings as xw
import xlrd
import seaborn as sns
import matplotlib.pyplot as plt


# Створюємо об'єкт ExcelFile.

# In[3]:


xlsx=pd.ExcelFile('temp_datas.xlsx')


# Отримуємо список книг

# In[4]:


worksheets = xlsx.sheet_names


# Створюємо список, в якому кожний елемент - це датафрейм із однієї книги.

# In[5]:


worksheets_dfs = []

for worksheet in worksheets:
    data = pd.read_excel(xlsx,
                         sheet_name=worksheet)
    data['worksheet'] = worksheet
    worksheets_dfs.append(data)


# Обєднати всі датафрейми із списку.

# In[6]:


meteo_data = pd.concat(worksheets_dfs)


# In[7]:


meteo_data["date"]=pd.to_datetime(meteo_data["Timestamp for sample frequency every 1 min"]).dt.date
meteo_data["months"]=pd.to_datetime(meteo_data["Timestamp for sample frequency every 1 min"]).dt.month_name()


# In[8]:


#meteo_data.head()
#meteo_data.sample(5)


# In[9]:


get_ipython().run_line_magic('matplotlib', 'inline')

sns.boxplot(x="months", y="Temperature_Celsius", data=meteo_data)


# In[10]:


meteo_data.describe()


# # Середні значення показників по місяцях

# In[11]:


meteo_data_grouped_mean=meteo_data.groupby('months').mean()

#Групування зведеної таблиці по місяцях в послідовному порядку (відповідно до послідовності в списку 'cats'),
#а не за алфавітом
#https://stackoverflow.com/questions/40816144/pandas-series-sort-by-month-index
cats = ['January', 'February', 'March', 'April','May','June', 'July', 'August','September', 'October', 'November', 'December']
meteo_data_grouped_mean.index=pd.CategoricalIndex(meteo_data_grouped_mean.index, categories=cats, ordered=True)
meteo_data_grouped_mean=meteo_data_grouped_mean.sort_index()

meteo_data_grouped_mean


# In[12]:


grouped_mean=meteo_data.groupby('date').mean().rename(columns={'Temperature_Celsius':  'mean_temperature_celsius',
                                                               'Relative_Humidity': 'mean_relative_humidity',
                                                               'VPD': 'mean_vpd',
                                                               'DEWPOINT_Celsius': 'mean_dewpoint_celsius'})


# In[13]:


#Розрахунок середніх денних показників та виокремлення окремої колонки активних середньодобових температур,
#де середньодобова температура більша 10 град. цельсія.
grouped_mean['act_mean_temperature_celsius']=grouped_mean['mean_temperature_celsius'].apply(lambda x: x if x >10 else 0)
grouped_mean["months"]=pd.to_datetime(grouped_mean.index).month_name()

#grouped_mean.sample(5)


# # Сума активних температур

# In[14]:


grouped_sum=grouped_mean.groupby('months').sum()

#Групування зведеної таблиці по місяцях в послідовному порядку (відповідно до послідовності в списку 'cats'),
#а не за алфавітом
#https://stackoverflow.com/questions/40816144/pandas-series-sort-by-month-index
cats = ['January', 'February', 'March', 'April','May','June', 'July', 'August','September', 'October', 'November', 'December']
grouped_sum.index=pd.CategoricalIndex(grouped_sum.index, categories=cats, ordered=True)
grouped_sum=grouped_sum.sort_index()

#Налаштування відображення гістограми
#https://matplotlib.org/3.5.0/plot_types/basic/bar.html#sphx-glr-plot-types-basic-bar-py

labels=grouped_sum.index
sum_act_temp=grouped_sum['act_mean_temperature_celsius']
width = 0.90

fig, ax = plt.subplots()

bar1=ax.bar(labels, sum_act_temp, width, label='Сума активних температур', color='orange')

ax.set_ylabel('Активні температури, гр. цельсія')
ax.set_title('Сума активних температур по місяцях')
ax.legend()

ax.bar_label(bar1, padding=1)

plt.show()


# ## Значення суми активних температур 2022 р. (в р-ні с. Новичка, Калуського району)

# In[15]:


print('Сума активних температур (>10 гр. цельсія) становить', grouped_mean['act_mean_temperature_celsius'].sum().round(decimals=2),' гр. цельсія')


# In[ ]:





# In[ ]:




