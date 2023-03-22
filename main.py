import pandas as pd
import glob
import tkinter as tk

def summaryReport(): 
    """load the prism report from the folder specified, must be a xlsx file"""
    folder_path = "./Summary/"
    xlsx_files = glob.glob(folder_path + "*.xlsx")

    for xlsx_file in xlsx_files: #loop through all the files
        cleanSummary(xlsx_file)

def cleanSummary(xlsx_file):
    """clean the file"""
    df = pd.read_excel(xlsx_file) #load the file
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
    writer = pd.ExcelWriter(xlsx_file, engine='xlsxwriter') #create a writer object
    new_df.to_excel(writer, sheet_name="oli", index=False) #create the excel
                    
    workbook = writer.book
    worksheet = writer.sheets['oli']

    for i, column in enumerate(df.columns): #loop through each column and set the width to 25
        column_width = 25
        worksheet.set_column(i, i, column_width)
    writer.save()

def exportReport(): 
    """load the prism report from the folder specified, must be a xlsx file"""
    folder_path = "./Export/"
    xlsx_files = glob.glob(folder_path + "*.xlsx")

    for xlsx_file in xlsx_files: #loop through all the files
        cleanExport(xlsx_file)

def cleanExport(xlsx_file):
    df = pd.read_excel(xlsx_file) #load the file
    cols_to_drop = ['Length', 'Definition', 'Example', 'Format', 'Business Rule']
    cols_to_drop.extend([col for col in df.columns if col.startswith('[Model')])
    df = df.drop(cols_to_drop, axis=1)
    df = df.drop([0, 1, 2])

    new_column2 = df['ContainerName'].str.split('/', n=1, expand=True)[1].str.split('/', n=1, expand=True)[0]
    df.insert(loc=0, column='root2', value=new_column2)

    new_column = df['ContainerName'].str.split('/', n=1, expand=True)[0]
    df.insert(loc=0, column='root1', value=new_column)

    df['ContainerName'] = df['ContainerName'].str.split('/', n=2, expand=True)[2]

    container_type = df.pop('ContainerType')
    df.insert(loc=0, column='ContainerType', value=container_type)

    df.to_excel(xlsx_file, index=False)


def main():
    """create the GUI"""
    window = tk.Tk()
    window.title("Ramiel v1.0")
    window_width = 420
    window_height = 420

    canvas = tk.Canvas(window, width=window_width, height=window_height)
    canvas.pack()

    text = "( ͡° ͜ʖ ͡°)"
    text_color = "black"
    text_size = 20

    canvas.create_text(window_width/2, window_height/4, text=text, fill=text_color, font=("Arial", text_size), anchor=tk.CENTER)
    buttonSummary = tk.Button(window, text="SummaryReport",font=("Arial", text_size), command=summaryReport)
    buttonSummary.place(relx=0.35, rely=0.5, anchor=tk.CENTER)

    buttonExport = tk.Button(window, text="Export",font=("Arial", text_size), command=exportReport)
    buttonExport.place(relx=0.8, rely=0.5, anchor=tk.CENTER)


    window.mainloop()

if __name__ == "__main__":
    main()
    