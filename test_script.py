######################################
########## Import Libraries ##########
######################################

# Import Library for Accessing XML File.
import xml.etree.ElementTree as ET
# Deal with Tabular Data using Pandas Library.
import pandas as pd

######################################################
########## Shorten ID Column in Excel File ###########
######################################################

raw_dataframe = pd.read_excel('test_excel_1.xlsx')

# This function shortens an ID from the Excel file.
def shorten_id(original_column) :
    length_to_cut = len('Delete Account ')
    length_total = len(original_column)
    truncated_column = original_column[length_to_cut: length_total]

    return truncated_column

# Apply function to all columns of the raw dataframe.
raw_dataframe['Col1'] = raw_dataframe['Col1'].apply(shorten_id)
# Create a series to contain column information.
id_series = raw_dataframe['Col1']
submission_date_series = raw_dataframe['Col5']
due_date_series = raw_dataframe['Col6']

##############################################
######## Acquire Data from XML File ##########
##############################################

# Deal with XML file content using ElementTree Library.
    # The tree is used to access the entire file,
    # whereas the root is used to access the topmost tag.
tree = ET.parse('test_xml.xml')
root = tree.getroot()

# Create empty lists to append information when iterating through XML file.
role_column = []
organized_role_column = []
xml_ids = []

# This function shortens the role information from the XML file.
def shorten_role(original_column) :
    length_to_cut = len('Role=')
    length_total = len(original_column)
    truncated_column = original_column[length_to_cut: length_total]

    return truncated_column

# Iterate through XML file to filter and organize relevant information.
for account in root.iter('account') :
    id = account.get('id', default = None)

    for i in range(id_series.size) :
        # If ID is found in the ID series, execute code.
        if (id == id_series[i]) :
            # Add this unordered ID to the list.
            xml_ids.append(id)
            ## The following lines are commented out -> used for debugging loop.
            ## print(id)
            ## print(id_series[i])
            ## print('\n')

            # Assign role column to have the role information.
            for attribute in account.iter('attribute') :
                ## The following lines are commented out -> used for debugging loop.
                ## print(attribute.attrib)
                ## print('\n')

                if attribute.get('name') == 'Role' :
                    attributeValueRef_id = str()

                    for attributeValueRef in attribute.iter('attributeValueRef') :
                        ## The following line is commented out -> used for debugging loop.
                        ## print(attributeValueRef.attrib)

                        # Remove the 'Role=' using the custom function before acquiring the information.
                        attributeValueRef_id += shorten_role(attributeValueRef.get('id')) + ' \n'
                    # Append each string to the list if the ID is found in the ID series.
                    role_column.append(attributeValueRef_id)

                    ## The following lines are commented out -> used for debugging loop.
                    ## print('\n')
                    ## print(role_column)
                    ## print('\n')

#####################################################
###### Reorder Role Column To Match ID Series #######
#####################################################

# Organize role column to match the ordered ID series.
for organized_index in range(id_series.size) :
    ## The following lines are commented out -> used for debugging loop.
    ## print(organized_index)
    ## print(id_series[organized_index])
    ## print('\n')

    # Add an empty string to end of list to account for IDs with blank roles.
    organized_role_column.append('')
    for unorganized_index in range(len(xml_ids)) :
        ## The following lines are commented out -> used for debugging loop.
        ## print(unorganized_index)
        ## print(xml_ids[unorganized_index])
        ## print('\n')

        if (id_series[organized_index] == xml_ids[unorganized_index]) :
            # Replace empty string for IDs with roles in the XML file.
            organized_role_column.insert(organized_index, (role_column[unorganized_index]))


#############################################
######## Preparing Final DataFrame ##########
#############################################

# Convert the list to a series.
organized_role_series = pd.DataFrame(organized_role_column)
# Combine the series into a new dataframe.
combined_dataframe = pd.concat([id_series, organized_role_series, submission_date_series, due_date_series], axis = 1)
# Rename the columns in the dataframe for readability.
combined_dataframe.rename(columns = {'Col1' : 'ID', 0 : 'Role Information', 'Col5' : 'Submission Date', 'Col6' : 'Due Date'}, inplace = True)
# View final dataframe before inputting to Excel file.
print(combined_dataframe)

##########################################################
############## Writing to Excel Spreadsheet ##############
##########################################################

# Use Pandas library function ExcelWriter to input the DataFrame.
writer = pd.ExcelWriter('test_excel_2.xlsx', engine = 'xlsxwriter')

# Write content to new spreadsheet.
combined_dataframe.to_excel(writer,
             sheet_name ='organized_sheet')

# Ensure the Excel file isn't open at the time of saving.
writer.save()
writer.close()
