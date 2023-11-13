'''
Coder: Cristina M*********
Company: Qo***
Date: 11/13/23
Purpose: normalize vendors name in an excel file, based on a reference excel file.
'''
import pandas as pd 


def find_in_ref(name, df_ref, trunc_len = 3):
    '''
    Function that return the index position of a string called 
    'name' in the column 'Name:' of a 'df_ref' dataframe. The string 
    can be truncated to do a pattern research and increase the chance to 
    find the term even in case of misspelling.
    '''
    name = name[:trunc_len]
    is_in_ref = df_ref['Name:'].str.contains(name, case=False)
    index = is_in_ref[is_in_ref[0:] == True].index
    return index 


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
    # Declaration of the dataframes containing excel file data
    ## Reference files 
    df_ref_address= pd.read_excel('ref_address.xlsx',sheet_name='Sheet1')
    df_ref_zipcode= pd.read_excel('ref_zipcode.xlsx',sheet_name='Sheet1')
    ## Target file to normalize
    df_target= pd.read_excel('target.xlsx',sheet_name='Sheet1')
    
    # Declaration of the list that will contained failed row
    failed_line_list = []
    
    # For each row of the target data:
    for row in range(len(df_target)):

        name  = df_target.loc[row, 'Name:']

        # Get the index of the company name in the reference files 
        index_ref_add = find_in_ref(name, df_ref_address)
        index_ref_zip = find_in_ref(name, df_ref_zipcode)

        # If company name is present in ref_address:
        if len(index_ref_add.values) > 0:
            df_target.loc[row, 'Name:'] = df_ref_address.loc[index_ref_add.values[0], 'Name:']

        # Else if company name is present in ref_zipcode:
        elif len(index_ref_zip.values) > 0:
            df_target.loc[row, 'Name:'] = df_ref_zipcode.loc[index_ref_zip.values[0], 'Name:']

        # Else it is considered as failed and logged in the failed_line_list:
        else: 
            failed_line_list.append(df_target.loc[row].values)
            
    # Save the failed_line_list log in a txt file named failed_line_output.txt
    list_to_txt_file(failed_line_list)
    # Save the normalized dataframe into a new excel file named output.xlsx
    df_target.to_excel("output.xlsx", index=False)