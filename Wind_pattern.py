import pandas as pd
import numpy as np

df = pd.read_excel(r'C:\Users\jones\PycharmProjects\Energy_data\generationData.xlsx', skiprows=1,
                   sheet_name='Sheet1', header=[0, 1, 2])
df = df.iloc[:, 1:]
pd.set_option('display.max_column', None, 'display.max_rows', None)

# Drop NAN values in df
df.dropna(axis=1, inplace=True)

# Create new column 'Hours' == df[:,0] delete original df[:,0] first column of df
df['Hours'] = df.iloc[:, 0]

# Delete original df[:,0] first column of df
df.drop(df.columns[0], inplace=True, axis=1)

# Set index
df.set_index('Hours', inplace=True)

# Drop tail(1) which is the already calculated average
df.drop(df.tail(1).index, inplace=True)

# Renaming columns to be more specific
df.rename({'Wind': 'Onshore Wind', 'Li-Ion': 'Li-Ion Battery'}, axis=1, level=1, inplace=True)

# Creating a dataframe of the columns.
columns_df = pd.DataFrame(df.columns)
#print(columns_df)
#print('......................')
#print(np.where(df.columns.get_level_values(0) == 'Bus1'))
#print(np.where(df.columns.get_level_values(1) == 'Coal'))

# Conditions
condition1 = df.columns.get_level_values(1) == 'Offshore wind'
condition2 = df.columns.get_level_values(2) == 'Generation2025(MWh)'

indices = np.where((condition1 & condition2))
#print(type(indices[0]))
# Convert array to a list
indices = indices[0].tolist()
#print(indices)

print('___________')
#print(len(indices))

for i in range(0, len(indices)):
    indices[i] = indices[i] - 1
    #print(indices[i])

array1 = np.arange(1, len(indices) + 1)
list1 = array1.tolist()
#print(list1)

#print(list(zip(list1, indices)))

#
insertion_locations = [a+b for (a, b) in zip(list1, indices)]
#print(insertion_locations)
#print(len(insertion_locations))

# Insert an empty column after offshore  wind, Generation 2025(Mwh) column for a given bus
for i in range(0, len(insertion_locations)):
    df.insert(insertion_locations[i], (df.columns[insertion_locations[i]-1][0], 'Offshore wind',
                                       'Normalized values'),'')
    #print(pd.DataFrame(df.columns))

# Create new row
df.loc['Maximum', :] = df.max()
#print(df)

wind_factor_locations = np.where(df.columns.get_level_values(2) == 'Normalized values') #returns a tuple with two values (an array, type)

# Convert from array to list
wind_factor_locations = wind_factor_locations[0]
wind_factor_locations = wind_factor_locations.tolist()
print(wind_factor_locations)

print(df.iloc[-1, 8])

# For loop
for k in wind_factor_locations:
    df.iloc[:, k] = df.iloc[:, k+1] / df.iloc[-1, k+1]
#print(df)

# Dropping the last column
df.drop(df.tail(1).index, inplace=True)
#print(df)

relevant_cols = df.columns[(df.columns.get_level_values(1) != 'Offshore wind') &
                              (df.columns.get_level_values(2) != 'Generation2025(MWh)')
                              | (df.columns.get_level_values(1) != 'Offshore wind') &
                              (df.columns.get_level_values(2) != 'Generation2020(MWh)')]
df_no_offsh_wind = df[relevant_cols]
#print(df_no_offsh_wind)


relevant_columns = df.columns[(df.columns.get_level_values(1) == 'Offshore wind') &
                              (df.columns.get_level_values(2) == 'Generation2025(MWh)')]
df_offshore_wind = df[relevant_columns]
#print(df_offshore_wind)

# Capturing the wind profile only
df_relevant_cols = df.columns[(df.columns.get_level_values(2) == 'Normalized values')]
df_wind_profile = df[df_relevant_cols]
print(df_wind_profile)

# Copy to excel
with pd.ExcelWriter('output generation.xlsx') as writer:
    df.to_excel(writer, sheet_name='df')
    df_offshore_wind.to_excel(writer, sheet_name='Offshore_wind')
    df_wind_profile.to_excel(writer, sheet_name='Wind profile ')


