import pandas as pd
import numpy as np
import copy


    
def cleardf(df):
    newdfs = {}
    year, milesRange = df.columns[0].split()[0], df.columns[0].split()[-1]
    # year, min_miles, max_miles = df.columns[0][:4], *df.columns[0][-7:].split('-')
    # min_miles, max_miles = int(min_miles), int(max_miles)
    prebottom = -2
    for bottom in df[df.iloc[:,0].isnull()].index:  # cutting the original df to multiple dfs by empty row
        print(bottom)
        newdf = copy.deepcopy(df.iloc[prebottom+3:bottom, :16])
        newdf.columns = list(df.iloc[prebottom+2, :16])
        newdfs[year + ' ' + milesRange] = newdf
        prebottom = bottom
        year, milesRange = df.iloc[bottom+1, 0].split()[0], df.iloc[bottom+1, 0].split()[-1]
    # append the last df
    newdf = copy.deepcopy(df.iloc[prebottom+3:len(df), :16])
    newdf.columns = list(df.iloc[prebottom+2, :16])
    newdfs[year + ' ' + milesRange] = newdf
    return newdfs
        
        
def createNewDf(df, title):
    newdf = pd.DataFrame(columns=['Policy', 'Months', 'Miles', 'nPolicy', 'Year', 'Min_Miles', 'Max_Miles', 'Class', 'Supplier_cost'])
    year, min_miles, max_miles = title.split(' ')[0], *title.split(' ')[1].split('-')
    term = df['Term/Miles']
    for clmnName in list(df.columns)[1:]:  # loop along with columns
        # print(clmnName)
        tmpdf = pd.DataFrame(columns=['Policy', 'Months', 'Miles', 'nPolicy', 'Year', 'Min_Miles', 'Max_Miles', 'Class', 'Supplier_cost'])
        tmpdf['Policy'] = term
        tmpdf[['Months','Miles']] = df['Term/Miles'].str.split('/', expand=True)
        tmpdf['Miles'] = tmpdf['Miles'].str.replace(r'k$', '')    # remove the 'k' char
        tmpdf['nPolicy'] = term.str.replace(r'k$', '')
        tmpdf.Year, tmpdf.Min_Miles, tmpdf.Max_Miles, tmpdf.Class = year, min_miles, max_miles, clmnName.split()[1]
        tmpdf.Supplier_cost = df[clmnName]
        # print('tmpdf:')
        # print(tmpdf.head())
        newdf = newdf.append(tmpdf)
        # print('newdf: ')
        # print(newdf.head())
    return newdf


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
    # newdfs = {}
    convertedDf = pd.DataFrame(columns=['Policy', 'Months', 'Miles', 'nPolicy', 'Year', 'Min_Miles', 'Max_Miles', 'Class', 'Supplier_cost'])
    for sht in list(dfs.keys())[1:2]:
        # newdfs.update(cleardf(dfs[sht]))
        newdfs = cleardf(dfs[sht])
        for newsht in list(newdfs.keys()):
            newdf = createNewDf(newdfs[newsht], newsht)
            convertedDf = convertedDf.append(newdf)
    saveDst(convertedDf, dstfile, path)
    # convertedDf.to_excel(path+dstfile, index=False)
    
        
        
    

