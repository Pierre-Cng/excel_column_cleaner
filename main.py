'''
Coder: Cristina M*********
Company: Qo***
Date: 11/13/23
Purpose: normalize vendors name in an excel file, based on a reference excel file.
'''
import pandas as pd 

class df_excel_file():
    def __init__(self, dict):
        parameters = dict
        df = self.dict_to_dataframe(dict)
    
    def dict_to_dataframe(self, dict):
        '''
        Return a dataftrame made from excel data referenced in dict.
        '''
        return pd.read_excel(dict['file_name'],sheet_name=dict['sheet_name'])
    

class normalizer():
    def __init__(self, dict_target, dict_ref_address, dict_ref_zipcode):
        target = df_excel_file(dict_target)
        ref_address = df_excel_file(dict_ref_address)
        ref_zipcode = df_excel_file(dict_ref_zipcode)
    
    def find_in_column(self, string, df_ref, column, trunc_len = 3):
        '''
        Return the index position of a string in a dataframe column.
        The string can be truncated to do a pattern research and increase the 
        chance to find the term.
        '''
        string = string[:trunc_len]
        is_in_column = df_ref[column].str.contains(string, case=False)
        index = is_in_column[is_in_column[:] == True].index
        return index
    
    def check_correspondance(self, row, column):
        string  = self.target.df.loc[row, self.target.parameters[column]]
        for ref in [self.ref_address, self.ref_zipcode]:
            # Get the index of the string in the reference files 
            index = self.find_in_column(string, ref.df, ref.parameters[column])
            if len(index.values) > 0:
                

        index_ref_zip = self.find_in_column(string, self.df_ref_zipcode, self.ref_zipcode['column_name'])

    # If company name is present in ref_address:
    if len(index_ref_add.values) > 0:
        df_target.loc[row, target['column_name']] = df_ref_address.loc[index_ref_add.values[0], ref_address['column_name']]

    # Else if company name is present in ref_zipcode:
    elif len(index_ref_zip.values) > 0:
        df_target.loc[row, target['column_name']] = df_ref_zipcode.loc[index_ref_zip.values[0], ref_zipcode['column_name']]

    # Else it is considered as failed and logged in the failed_line_list:
    else: 
        failed_line_list.append(df_target.loc[row].values)

 



def normalize_name(target, ref_address, ref_zipcode):
    '''
    Function to normalize company name using address reference and zipcode reference.
    If the name is not contained in any of the reference file, the name row is logged 
    into 'failed_line_list'.
    '''
    # Declaration of dataframes
    df_target, df_ref_address, df_ref_zipcode = dict_to_dataframe(target), dict_to_dataframe(ref_address), dict_to_dataframe(ref_zipcode)
        
    # Declaration of the list that will contained failed row
    failed_line_list = []
    
    # For each row of the target data:
    for row in range(len(df_target)):

        

    return df_target, failed_line_list





def list_to_txt_file(mylist):
    '''
    Function to log the element of mylist that contains the 
    row of the df_target file that failed to be normalized
    '''
    with open("failed_line_output.txt", 'w') as file:
        for item in mylist:
            file.write(str(item) + '\n')
        file.close()


if __name__ == "__main__":
    # Parameters to define by user:
    target = {'file_name' :  'target.xlsx',
                   'sheet_name' : 'Sheet1',
                   'column_name' : 'Name:', 
                   'column_attibute_1' : 'Address:', 
                   'column_attibute_2' : 'Zipcode:'}
      
    ref_address = {'file_name' : 'ref_address.xlsx',
                        'sheet_name' : 'Sheet1',
                        'column_name' : 'Name:',
                        'column_attribute' : 'Address:'}
    
    ref_zipcode = {'file_name' : 'ref_zipcode.xlsx',
                        'sheet_name' : 'Sheet1',
                        'column_name' : 'Name:',
                        'column_attibute' : 'Zipcode:'}
    
    # Normalizing names
    df_output, failed_line_list = normalize_name(target_file, ref_address_file, ref_zipcode_file)
    
    # Save the failed_line_list log in a txt file named failed_line_output.txt
    list_to_txt_file(failed_line_list)

    # Save the normalized dataframe into a new excel file named output.xlsx
    df_output.to_excel("output.xlsx", index=False)


    
    
            
    
    