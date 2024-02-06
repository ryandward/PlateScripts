import pandas as pd
import numpy as np
import sys

def process_excel(file_path):
    
    pd.set_option('display.max_columns', None)
    pd.set_option('display.max_rows', None)
    
    # Read the Excel file into a DataFrame
    df = pd.read_excel(file_path, header=None)

    # Find the row index where 'Cycle Nr.' is located
    header_row = df[df[0] == 'Cycle Nr.'].index[0]

    # Find the row index where 'H12' is located
    footer_row = df[df[0] == 'H12'].index[0]

    # Slice the DataFrame to include only the rows from 'Cycle Nr.' to 'H12'
    df = df.loc[header_row:footer_row]

    # Transpose the DataFrame
    df = df.T

    # Set the column names to the first row
    df.columns = df.iloc[0]

    # Drop the first row as it's now the column names
    df = df.iloc[1:]

    # Reset the index
    df.reset_index(drop=True, inplace=True)

    # Melt the DataFrame
    df_melted = df.melt(id_vars=['Cycle Nr.', 'Time [s]', 'Temp. [°C]'], var_name='Well', value_name='OD600')

    # rename Cycle Nr. to Cycle, Time [s] to Time, Temp. [°C] to Temp
    df_melted.rename(columns={'Cycle Nr.': 'Cycle', 'Time [s]': 'Time', 'Temp. [°C]': 'Temp'}, inplace=True)

    # rename rows to Cycle, Time, Temp, Well, OD600, Column(*1-12 from Well), Row (A-H* from Well)
    df_melted['Row'] = df_melted['Well'].str.extract('([A-H])')
    df_melted['Column'] = df_melted['Well'].str.extract('(\d+)')
    
    # Reorder the columns
    df_melted = df_melted[['Well', 'Row', 'Column', 'OD600', 'Time', 'Temp', 'Cycle']]

    
    # filter where OD600 is NaN
    df_melted = df_melted[df_melted['OD600'].notna()]


    print(df_melted.to_csv(sep='\t', index=False))

if __name__ == "__main__":
    process_excel(sys.argv[1])