#tools I will need to import, visualize, combine, and export data
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import openpyxl as pxl
import xlsxwriter as xlw
import random
import xlrd

#files from where data will be pulled and from where data will be manipulated
excel_file_1 = 'buttons.xlsx'
excel_file_2 = 'fabrics.xlsx'
excel_file_3 = 'dyes.xlsx'
excel_file_4 = 'clients.xlsx'
excel_file_5 = 'dye_fabric_costs.xlsx'
excel_file_6 = 'order_form.xlsx'

# data frames of all the excel spreedsheets that will be used in this final project
df_button = pd.read_excel(excel_file_1, sheet_name='buttons')
df_fabric_1 = pd.read_excel(excel_file_2, sheet_name='silk')
df_fabric_2 = pd.read_excel(excel_file_2, sheet_name='chiffon')
df_fabric_3 = pd.read_excel(excel_file_2, sheet_name='cotton')
df_fabric_4 = pd.read_excel(excel_file_2, sheet_name='linen')
df_fabric_5 = pd.read_excel(excel_file_2, sheet_name='velvet')
df_fabric_6= pd.read_excel(excel_file_2, sheet_name='crepe')
df_fabric_7= pd.read_excel(excel_file_2, sheet_name='wool')
df_fabric_8 = pd.read_excel(excel_file_2, sheet_name='denim')
df_fabric_9 = pd.read_excel(excel_file_2, sheet_name='polyester')
df_fabric_10 = pd.read_excel(excel_file_2, sheet_name='flannel')
df_fabric_11 = pd.read_excel(excel_file_2, sheet_name='rayon')
df_fabric_12 = pd.read_excel(excel_file_2, sheet_name='corduroy')
df_fabric_13 = pd.read_excel(excel_file_2, sheet_name='organza')
df_fabric_14 = pd.read_excel(excel_file_2, sheet_name='fleece')
df_fabric_15 = pd.read_excel(excel_file_2, sheet_name='satin')
df_fabric_16 = pd.read_excel(excel_file_2, sheet_name='chenille')
df_fabric_17 = pd.read_excel(excel_file_2, sheet_name='muslin')
df_fabric_18 = pd.read_excel(excel_file_2, sheet_name='georgette')
df_fabric_19 = pd.read_excel(excel_file_2, sheet_name='lace')
df_fabric_20 = pd.read_excel(excel_file_2, sheet_name='poplin')
df_fabric_21 = pd.read_excel(excel_file_2, sheet_name='batiste')
df_fabric_22 = pd.read_excel(excel_file_2, sheet_name='lawn')
df_fabric_23 = pd.read_excel(excel_file_2, sheet_name='nylon')
df_fabric_24 = pd.read_excel(excel_file_2, sheet_name='gabardine')
df_dye = pd.read_excel(excel_file_3, sheet_name='dye')
df_client = pd.read_excel(excel_file_4, sheet_name='clients')

# Combining all types of fabrics to visualize them in one single data frame
df_all = pd.concat([df_fabric_1,df_fabric_2,df_fabric_3,df_fabric_4,
df_fabric_5,df_fabric_6,df_fabric_7,df_fabric_8,df_fabric_9,df_fabric_10,df_fabric_11,
df_fabric_12,df_fabric_13,df_fabric_14,df_fabric_15,df_fabric_16,df_fabric_17,df_fabric_18,
df_fabric_19, df_fabric_20, df_fabric_21, df_fabric_22, df_fabric_23, df_fabric_24])

print(df_all)
df_all.to_excel('output_all_fabrics.xlsx')

# Filtering 1 yard of each fabric into one single dataframe
df_fabric_all = pd.concat([df_fabric_1.iloc[:1],df_fabric_2.iloc[:1],df_fabric_3.iloc[:1],df_fabric_4.iloc[:1],
df_fabric_5.iloc[:1],df_fabric_6.iloc[:1],df_fabric_7.iloc[:1],df_fabric_8.iloc[:1],df_fabric_9.iloc[:1],df_fabric_10.iloc[:1],df_fabric_11.iloc[:1],
df_fabric_12.iloc[:1],df_fabric_13.iloc[:1],df_fabric_14.iloc[:1],df_fabric_15.iloc[:1],df_fabric_16.iloc[:1],df_fabric_17.iloc[:1],df_fabric_18.iloc[:1],
df_fabric_19.iloc[:1], df_fabric_20.iloc[:1], df_fabric_21.iloc[:1], df_fabric_22.iloc[:1], df_fabric_23.iloc[:1], df_fabric_24.iloc[:1]])

print(df_fabric_all)
df_fabric_all.to_excel('fabric_cost_per_yard.xlsx')

# Transposing the Dye dataframe
df_dye_t = df_dye.T
print(df_dye_t[1:3])
df_dye_t.to_excel('transposed_colors.xlsx')

# Combined Fabric and Dye Cost Dataframe
df_dye_fabric_costs = pd.read_excel(excel_file_5, sheet_name='Sheet1')  
print (df_dye_fabric_costs)

# Transposing the Client Roster
client_t = df_client.T
print(client_t)
client_t.to_excel('transposed_clients.xlsx')

import os
import random
from numpy import number
import pandas as pd

# CONSTANTS
projects_dir = '/Users/Sabrina/projects'
excel_client_file = os.path.join(projects_dir, 'clients.xlsx')
excel_order_file = os.path.join(projects_dir, 'order_form.xlsx')

RATE_COLUMN_NAME = 'Raw Price'
TOTAL_COLUMN_NAME = 'Raw Total'
MARKUP = 'Markup'
PROFIT = 'Profit'

df = pd.read_excel(excel_order_file, sheet_name='buttons')
df_client = pd.read_excel(excel_client_file, sheet_name='clients')
clients = list(df_client['Clients'])

number_rows = len(df)

# for each client add a column using client name as column name
for client in clients:
    rows = []
    for row in range(0, number_rows):
        value = random.randrange(0, 50)
        rows.append(value)

    # use client name for column name
    column_name = client

    # add rows to column
    df[column_name] = rows

# to store the total values for every row
totals = []

for row in range(0, number_rows):
    # the rate for this row
    rate = df[RATE_COLUMN_NAME][row]

    # sum the values for all the client columns for this row
    client_sum = 0
    for client in clients:
        client_sum += df[client][row]

    # multiply the rate by the client sum
    row_total = rate * client_sum

    # add the total for this row to the list
    totals.append(row_total)

# add a column called 'Total'
df[TOTAL_COLUMN_NAME] = totals

# Assigning a markup value for raw totals
df[MARKUP] = df[TOTAL_COLUMN_NAME] * 4

# Calculating Sales Profits
profits = []
for row in range(0,len(df)):
    profit = df[MARKUP][row] - totals[row]

    profits.append(profit)

df[PROFIT] = profits

print(df)
df.to_excel('order_form_sample random.xlsx')