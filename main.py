import pandas as pd 

class excel_file():
    '''
    Class to build object containing excel file data into pandas dataframe
    '''
    def __init__(self, file_name, sheet_name, column):
        self.file_name = file_name
        self.sheet_name = sheet_name
        self.column = column
        self.df = pd.read_excel(file_name,sheet_name= sheet_name).astype(str)

def find_in_column(string, column, trunc_len = 3):
        '''
        Find index of a (truncated) string in a given column
        '''
        if isinstance(string, str) is False:
            string = str(string)
        if trunc_len != -1:
            string = string[:trunc_len]
        is_in_column = column.str.contains(string, case=False)
        index = is_in_column[is_in_column[:] == True].index
        return index

def find_in_column_recurssive(string, column):
    '''
    Use of the function find_in_column recurssively until only one index remains
    '''
    index = find_in_column(string, column)
    term = 0
    while len(index) > 1:
         index = find_in_column(string.split()[term], column, -1)
         term +=1 
    return index 

def column_parser(row, target, ref):
    '''
    Parse column of ref dataframe and execute the function find_in_column_recurssive until it find a matching index or parsed all data
    '''
    index = []
    for key in ref.column.keys():
        if len(index) == 0:
            string = target.df.loc[row, target.column[key]]
            if string != 'nan':
                index = find_in_column_recurssive(string, ref.df[ref.column[key]])
        else:
            return index.values[0]
    try: 
        return index.values[0]
    except:
        return -1 

def failed_list_to_txt(failed_list):
    '''
    Write list of failed line into failed_list.txt file
    '''
    with open("failed_list.txt", 'w') as file:
        for item in failed_list:
            file.write(str(item) + '\n\n')
        file.close()

if __name__ == "__main__":
    '''
    Main code, will be executed if the code is run as a main and not as a library
    '''

    ###### This is the info you need to modify for your own usage #####
    ###### Syntax:
     #------=-----------('your file name.xlsx', 'the sheet name in excel',{'name' : 'column name for Name in excel', 'address' : 'column name for Address in excel', 'zipcode' : 'column name for Zipcode in excel'})
    # target is the file you want to normalize 
    target = excel_file('target.xlsx', 'Sheet1', {'name' : 'Name:', 'address' : 'Address:', 'zipcode' : 'Zipcode:'})
    # ref_file is your reference file 
    ref_file = excel_file('ref_file.xlsx', 'Sheet1', {'name' : 'Name:', 'address' : 'Address:', 'zipcode' : 'Zipcode:'})
    
    failed_list=[]

    for row in range(len(target.df)):       
        index = column_parser(row, target, ref_file)
        if index != -1:   
            target.df.loc[row, target.column['name']] = ref_file.df.loc[index, ref_file.column['name']] 
            target.df.loc[row, target.column['address']] = ref_file.df.loc[index, ref_file.column['address']] 
            target.df.loc[row, target.column['zipcode']] = ref_file.df.loc[index, ref_file.column['zipcode']] 
            print(target.df.loc[row, target.column['name']], '- Index:', index)
        else:
            print('FAILED', 'Name:', target.df.loc[row, target.column['name']])
            failed_list.append(target.df.loc[row])
        
    failed_list_to_txt(failed_list)
    target.df.to_excel("output.xlsx", index=False)
    