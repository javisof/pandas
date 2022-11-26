# -------------------------------------------------------------------------------------------------
# NOTE: Rename Excel files to our file names
# -------------------------------------------------------------------------------------------------

# -------------------------------------------------------------------------------------------------
# Eliminate warnings when opening Excel. It appears with some files downloaded from Screaming Frog..
import warnings
warnings.simplefilter("ignore")
# -------------------------------------------------------------------------------------------------

import pandas as pd

# ---------------------------------------------------
# Path Excel files
# ---------------------------------------------------
path_screaming = 'file_screaming_frog.xlsx'
path_semrush = 'file_positions_semrush.xlsx'
# ---------------------------------------------------

# ---------------------------------------------------
# Get data from 1º Excel - file .xlsx
# ---------------------------------------------------
data_screaming = pd.read_excel(path_screaming)
print('\n-------------------------------------------------------------------------')
print('Screaming Frog Data (URLS and Word Count) - 10 sample rows:')
print('-------------------------------------------------------------------------')
data_screaming = data_screaming.sort_values(by=['Dirección'])[['Dirección', 'Recuento de palabras']]
print(data_screaming.head(10))

# ---------------------------------------------------
# Get data from 2º Excel - file .xlsx
# ---------------------------------------------------
data_semrush = pd.read_excel(path_semrush)
data_semrush = data_semrush.sort_values(by=['URL'])[['URL', 'Position']]
print('\n--------------------------------------------------------')
print('Semrush data (URLS and Position) - 10 sample rows:')
print('--------------------------------------------------------')
print(data_semrush.head(10))
# ---------------------------------------------------


# ----------------------------------------------------
# JOIN THE 2 DATAFRAMES (Screaming Frog + Semrush):
# * Only the necessary columns:
# ** Screaming Frog: Address, Word Count
# ** Semrush: URL, Position
# ----------------------------------------------------
# ***** BEFORE JOINING THEM, YOU HAVE TO COMPARE THE URLS TO ASSIGN THE VALUE THAT CORRESPONDES TO THEM.
# ***** Joining with merge removes rows that do not match.
print('\n---------------------------------')
print('Get size of the dataframes (data obtained from Excel):')
print('---------------------------------')
print('data_screaming size:', data_screaming.shape)
print('data_semrush size:', data_semrush.shape)
print('---------------------------------')
print('Unique urls in data_screaming:', len(data_screaming['Dirección'].unique()))
print('Unique urls in data_screaming data_semrush:', len(data_semrush['URL'].unique()))


# ----------------------------------------------------------------------------------
# Join the 2 dataframes with merge function pandas.
# ----------------------------------------------------------------------------------
inner_merge = data_semrush.merge(data_screaming, left_on='URL', right_on='Dirección')
print(inner_merge.head())

# Sort by URL and Position
inner_merge = inner_merge.sort_values(by=['URL', 'Dirección'])

# Remove Duplicate Columns - Duplicate URL and Position
inner_merge = inner_merge.drop('Dirección', axis=1)

# Change dataframe column name
inner_merge = inner_merge.rename(columns={'Recuento de palabras': 'Text length'})

# Export the union to Excel.
path = 'data_all.xlsx'
inner_merge.to_excel(path, index=False)

# **** In the python script folder an .xlsx file with name data_all.xlsx will be created.
