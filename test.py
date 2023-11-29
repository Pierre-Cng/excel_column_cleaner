import pandas as pd 

class excel_file():
    def __init__(self, file_name, sheet_name, column):
        self.file_name = file_name
        self.sheet_name = sheet_name
        self.column = column
        self.df = pd.read_excel(file_name,sheet_name= sheet_name).astype(str)

def find_in_column(string, column, trunc_len = 3):
        if isinstance(string, str) is False:
            string = str(string)
        if trunc_len != -1:
            string = string[:trunc_len]
        is_in_column = column.str.contains(string, case=False)
        index = is_in_column[is_in_column[:] == True].index
        return index

def find_in_column_recurssive(string, column):
    index = find_in_column(string, column)
    term = 0
    while len(index) > 1:
         index = find_in_column(string.split()[term], column, -1)
         term +=1 
    return index 

def row_normalizer(row, refs):
    index = []
    for ref in refs:
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
    with open("failed_list.txt", 'w') as file:
        for item in failed_list:
            file.write(str(item) + '\n\n')
        file.close()

if __name__ == "__main__":
    target = excel_file('target.xlsx', 'Sheet1', {'name' : 'Name:', 'address' : 'Address:', 'zipcode' : 'Zipcode:'})
    ref_address = excel_file('ref_address.xlsx', 'Sheet1', {'name' : 'Name:', 'address' : 'Address:'})
    ref_zipcode = excel_file('ref_zipcode.xlsx', 'Sheet1', {'name' : 'Name:', 'zipcode' : 'Zipcode:'})
    failed_list=[]
    refs = [ref_address, ref_zipcode]
    for row in range(len(target.df)):       
        index = row_normalizer(row, refs)
        if index != -1:   
            print('index', index)
            target.df.loc[row, target.column['name']] = ref_address.df.loc[index, ref_address.column['name']] 
            target.df.loc[row, target.column['address']] = ref_address.df.loc[index, ref_address.column['address']] 
            target.df.loc[row, target.column['zipcode']] = ref_zipcode.df.loc[index, ref_zipcode.column['zipcode']] 
        else:
            print('failed')
            failed_list.append(target.df.loc[row])
            #function list of failed 
            pass
        
    failed_list_to_txt(failed_list)
    target.df.to_excel("output.xlsx", index=False)
    