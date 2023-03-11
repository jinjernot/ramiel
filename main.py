import pandas as pd
import glob
folder_path = "./xlsx/"
xlsx_files = glob.glob(folder_path + "*.xlsx")


for xlsx_file in xlsx_files: #loop through all the files (WIP)

    def main():
        df = pd.read_excel("./xlsx/SummaryReport.xlsx", sheet_name='Directly pulled from prism') #load the file
        df.drop(index=range(5), inplace=True) #remove the first 5 rows
        df = df.rename(columns=df.iloc[0]).drop(df.index[0])


        split_string = lambda x: '/'.join(x.split('/')[2:]) if x and isinstance(x, str) and len(x.split('/')) >= 2 else x #clean up the string
        df['ContainerName'] = df['ContainerName'].apply(split_string)#split the string

        df['Tag'] = df['ContainerName'].str.extract('\[(.*?)\]', expand=False) #create new column with the tag
        df['ContainerName'] = df['ContainerName'].str.replace('\[.*?\]', '', regex=True) #clean the tag

        idx_chunk = (df.iloc[0] == 'ChunkValue').values #search for the columns chunk and M to swap
        idx_m = (df.iloc[0] == 'M').values
        df.loc[:, idx_chunk], df.loc[:, idx_m] = df.loc[:, idx_m].values, df.loc[:, idx_chunk].values

        nan_columns = df.columns[pd.isna(df.columns)].tolist() #look for NaN headers
        df = df.drop(columns=nan_columns)#remove columns with NaN header

        first_col = df.iloc[:, 0] #move last column to the second position
        last_col = df.iloc[:, -1]
        middle_cols = df.iloc[:, 1:-1]
        new_df = pd.concat([first_col, last_col, middle_cols], axis=1)
        print(new_df)


        writer = pd.ExcelWriter('./xlsx/updated.xlsx', engine='xlsxwriter') #create a writer object
        new_df.to_excel(writer, sheet_name="oli", index=False) #create the excel
                    
        workbook = writer.book
        worksheet = writer.sheets['oli']

        for i, column in enumerate(df.columns): #loop through each column and set the width to 25
            column_width = 25
            worksheet.set_column(i, i, column_width)
        writer.save()

if __name__ == "__main__":
    main()