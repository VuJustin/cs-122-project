import pandas as pd
import os
localPath = os.path.abspath(os.getcwd()) + '\\'


"""
Function read_file()
reads Excel datafile (sheet(s)) and returns dataframe(s)
Inputs: (1) path name of Excel data file
        (2) Excel sheet(s) name(s) where data is located
Output: Dataframe if one sheet is input
        Or, dictionary of dataframes if multiple sheets
        are input, keys are sheet names.
"""
def read_file(dataFile, *sheetName):
    try:
        # if only one sheet name is input,
        # return corresponding dataframe
        if (len(sheetName) == 1):
            df = pd.read_excel(dataFile, sheetName[0])
        else:
            # if multiple sheet names are input,
            # return a dictionary of dataframes
            n = len(sheetName)
            df = {}
            for i in sheetName:
                df[i] = pd.read_excel(dataFile, i)

    except FileNotFoundError:
        print('Not a valid filename. Try again.')
    except Exception:
        print('Uncaught error')
        print(dataFile, sheetName)
        exit(0)

    return df

# main
# path = Path()
# dataPath = "path/../Data/"
# print("Pwd = ", path)
# print("local working dir = ", localPath)
dataPath = localPath + 'Data\\'
# GHG data file
datafile = dataPath + "EDGARv8.0_FT2022_GHG_booklet_2023.xlsx"
print(datafile)

# sheets in the data
GHG_totals = "GHG_totals_by_country"
GHG_by_sector = "GHG_by_sector_and_country"
GHG_per_GDP = "GHG_per_GDP_by_country"
GHG_per_capita = "GHG_per_capita_by_country"
lulucf = "LULUCF_macroregions"

# create dataframes
# GHG totals
df_GHG_totals = read_file(datafile, GHG_totals)
df_GHG_totals = df_GHG_totals.set_index('Country')
# GHG by sector
df_GHG_sector = read_file(datafile, GHG_by_sector)
# GHG per GDP
df_GHG_per_GDP = read_file(datafile, GHG_per_GDP)
df_GHG_per_GDP = df_GHG_per_GDP.set_index('Country')
# GHG per capita
df_GHG_per_capita = read_file(datafile, GHG_per_capita)
df_GHG_per_capita = df_GHG_per_capita.set_index('Country')
# lulucf
df_lulucf = read_file(datafile, lulucf)

df = read_file(datafile, [GHG_totals, GHG_by_sector, GHG_per_GDP, GHG_per_capita, lulucf])
# print info
print(df_GHG_totals.info())
print("1st three numeric columns")
print(df_GHG_totals.iloc[0:4, 2:5])
print("GHG_per_GDP")
print(df_GHG_per_GDP.describe())
print(df_GHG_per_GDP.iloc[:, 30:])
print("GHG_per_capita")
print(df_GHG_per_capita.describe())
print(df_GHG_per_capita.iloc[:, 50:])

print("1st 3 numeric columns from df dict of GHG totals")
df[GHG_totals] = df[GHG_totals].set_index("Country")
print(df[GHG_totals].iloc[0:4, 2:5])

# GHG by sector
dummy = 1