import pandas as pd

def main():
    df = pd.read_excel("./xlsx/SummaryReport.xlsx", sheet_name='Directly pulled from prism')
    df.drop(index=range(5), inplace=True)
    df = df.rename(columns=df.iloc[0]).drop(df.index[0])
    #df.insert(1,"Tag",0)
    split_string = lambda x: '/'.join(x.split('/')[2:]) if x and isinstance(x, str) and len(x.split('/')) >= 2 else x
    df['ContainerName'] = df['ContainerName'].apply(split_string)

    df['Tag'] = df['ContainerName'].str.extract('\[(.*?)\]', expand=False)
    df['ContainerName'] = df['ContainerName'].str.replace('\[.*?\]', '', regex=True)
    cols_to_select = df.iloc[0].isin(['ChunkValue'])
    selected_cols = df.loc[:, cols_to_select]
    print(selected_cols)
    selected_cols = selected_cols.shift(axis=1, periods=-3)
    #nan_columns = df.columns[pd.isna(df.columns)].tolist()
    #df = df.drop(columns=nan_columns)
    
    df.to_excel('./xlsx/updated.xlsx', sheet_name='test', index=False)
    

if __name__ == "__main__":
    main()