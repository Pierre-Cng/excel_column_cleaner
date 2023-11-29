import pandas as pd 

class excel_file():
    def __init__(self, file_name, sheet_name, column):
        self.file_name = file_name
        self.sheet_name = sheet_name
        self.column = column
        self.df = pd.read_excel(file_name,sheet_name= sheet_name).astype(str)

def find_in_column(string, column, trunc_len = 3):
        '''
        Return the index position of a string in a dataframe column.
        The string can be truncated to do a pattern research and increase the 
        chance to find the term.
        '''
        #print(type(string))
        if isinstance(string, str) is False:
            string = str(string)
        if trunc_len != -1:
            string = string[:trunc_len]
        #print(column)
        is_in_column = column.str.contains(string, case=False)
        index = is_in_column[is_in_column[:] == True].index
        return index

def find_in_column_recurssive(string, column):
    index = find_in_column(string, column)
    term = 0
    while len(index) > 1:
         index = find_in_column(string.split()[term], column, -1)
         #print(index)
         term +=1 
    return index 

if __name__ == "__main__":
    target = excel_file('target.xlsx', 'Sheet1', {'name' : 'Name:', 'address' : 'Address:', 'zipcode' : 'Zipcode:'})
    ref_address = excel_file('ref_address.xlsx', 'Sheet1', {'name' : 'Name:', 'address' : 'Address:'})
    ref_zipcode = excel_file('ref_zipcode.xlsx', 'Sheet1', {'name' : 'Name:', 'zipcode' : 'Zipcode:'})

    
    index = []
    row = 1
    refs = [ref_address, ref_zipcode]
#for row in range(len(target.df[target.column['name']]))
    for ref in refs:
        for key in ref.column.keys():
            if len(index) == 0: 
                string = target.df.loc[row, target.column[key]]
                # print(string)
                if type(string) is not float:
                    index = find_in_column_recurssive(string, ref.df[ref.column[key]])
                else:
                     #print(type(string))
                     pass
            else:
                 break
           
            #index = find_in_column_recurssive(target.df.loc[7, target.column['name']], ref_address.df[ref_address.column['name']])
        
         
    print(index.values[0])
    index = index.values[0]
    target.df.loc[row, target.column['name']] = ref_address.df.loc[index, ref_address.column['name']] 
    target.df.loc[row, target.column['address']] = ref_address.df.loc[index, ref_address.column['address']] 

    target.df.to_excel("output.xlsx", index=False)