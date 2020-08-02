import pandas as pd
import numpy as np
import copy


# Each tab of the original file stores data for one year,
# and each tab is divided into several data blocks, 
# and each data block stores a miles range of data
# This function is used to clear each tab data into a unique dataframe dictionary
def cleardf(df):
    newdfs = {}
    # the columns in each tab stored key information, retrieve it
    year, milesRange = df.columns[0].split()[0], df.columns[0].split()[-1]
    prebottom = -2                                  # used to store the bottom of previous data block
    for bottom in df[df.iloc[:,0].isnull()].index:  # cutting the original df to multiple dfs by empty row
        newdf = copy.deepcopy(df.iloc[prebottom+3:bottom, :16])
        newdf.columns = list(df.iloc[prebottom+2, :16])
        newdfs[year + ' ' + milesRange] = newdf     # each data block is transfer to a df which is stored in newdfs dictionary
                                                    # the key of newdfs keep key information is originally stored on the head of each data block
        prebottom = bottom
        year, milesRange = df.iloc[bottom+1, 0].split()[0], df.iloc[bottom+1, 0].split()[-1]  # renew the key infomation
    # append the last df
    newdf = copy.deepcopy(df.iloc[prebottom+3:len(df), :16])
    newdf.columns = list(df.iloc[prebottom+2, :16])
    newdfs[year + ' ' + milesRange] = newdf
    return newdfs
        
# This function is used to reorganize data structure
# The key information on the beginning of each block, include year, min_miles, max_mile are converted to columns
# Term/Miles data need to be splited and stored in multiple columns
# Multiple class data in one row need to be converted to multiple rows with one class column
def createNewDf(df, title):    # the title is derived from the key of dictionary which return by cleardf()
    newdf = pd.DataFrame(columns=['Policy', 'Months', 'Miles', 'nPolicy', 'Year', 'Min_Miles', 'Max_Miles', 'Class', 'Supplier_cost'])
    year, min_miles, max_miles = title.split(' ')[0], *title.split(' ')[1].split('-')
    print(min_miles,max_miles,sep='-')
    term = df['Term/Miles']
    # reorganize the data structure
    for clmnName in list(df.columns)[1:]:  # loop along with columns
        tmpdf = pd.DataFrame(columns=['Policy', 'Months', 'Miles', 'nPolicy', 'Year', 'Min_Miles', 'Max_Miles', 'Class', 'Supplier_cost'])
        tmpdf['Policy'] = term
        tmpdf[['Months','Miles']] = df['Term/Miles'].str.split('/', expand=True)
        tmpdf['Miles'] = tmpdf['Miles'].str.replace(r'k$', '')    # remove the 'k' char
        tmpdf['nPolicy'] = term.str.replace(r'k$', '')
        tmpdf.Year, tmpdf.Min_Miles, tmpdf.Max_Miles, tmpdf.Class = year, min_miles, max_miles, clmnName.split()[1]
        tmpdf.Supplier_cost = df[clmnName]
        newdf = newdf.append(tmpdf)
    return newdf

# Convert data type and display format, save file
def saveDst(df, dstfile, path):
    df[['Months', 'Miles', 'Year', 'Min_Miles', 'Max_Miles', 'Class']] = \
        df[['Months', 'Miles', 'Year', 'Min_Miles', 'Max_Miles', 'Class']].astype(int)
    df.columns = ['Policy', 'Months', 'Miles', 'Policy', 'Year', 'Min_Miles', 'Max_Miles', 'Class', 'Supplier_cost']
    writer = pd.ExcelWriter(path+dstfile, engine='xlsxwriter')
    df.to_excel(writer, sheet_name='Converted')
    wb = writer.book
    ws = writer.sheets['Converted']
    num_fmt = wb.add_format({'num_format': '#####0'})
    ws.set_column('B:C', None, num_fmt)
    ws.set_column('F:I', None, num_fmt)
    money_fmt = wb.add_format({'num_format': '$###,###'})
    ws.set_column('J:J', None, money_fmt)
    writer.save()
    
    


if __name__ == '__main__':
    path = 'G:/Projects/Upwork/Snoopdrive/'
    scrfile = 'RoadVantage.xlsx'
    dstfile = 'ConvertedRoadVantage.xlsx'
    dfs = pd.read_excel(path+scrfile, sheet_name=None)
    convertedDf = pd.DataFrame(columns=['Policy', 'Months', 'Miles', 'nPolicy', 'Year', 'Min_Miles', 'Max_Miles', 'Class', 'Supplier_cost'])
    for sht in list(dfs.keys())[1:]:    # iterate tabs in dictionary of original file
        print(sht)
        newdfs = cleardf(dfs[sht])      # each tab data is convert to unique dictionary, 'year min_miles-max_miles': df
        for newsht in list(newdfs.keys()):    # iterate converted dictionary
            newdf = createNewDf(newdfs[newsht], newsht)
            convertedDf = convertedDf.append(newdf)
    saveDst(convertedDf, dstfile, path)

    
        
        
    

