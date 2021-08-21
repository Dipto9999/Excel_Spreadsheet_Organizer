# Excel Spreadsheet Organizer

## Contents
* [Overview](#Overview)
    * [Solution Details](#Solution-Details)

## Overview
This was a short task completed for a confidential corporate organizer where information from an <a href = "test_excel_1.xlsx"><b>Excel Spreadsheet</b></a> was
modified and combined with information extracted from an <b><a href = "test_xml.xml">XML</a></b> file to create a table in another <b><a href = "test_excel_2.xlsx">Excel Spreadsheet</a></b>. This was done through the use of open-source libraries in <b>Python</b> and run in <a href= "test_notebook.ipynb"><b>Jupyter Notebook</b></a>. This allowed me to further explore the <b>Pandas Library</b> and become more familiar working with <b>DataFrames</b>.

<i>Note : All sensitive information has been modified and replaced.</i>

### Solution Details
A few custom functions were built to shorten strings with a common prefixes, as shown here :

```python
# This function shortens an ID from the Excel file.
def shorten_id(original_column) :
    length_to_cut = len('Delete Account ')
    length_total = len(original_column)
    truncated_column = original_column[length_to_cut: length_total]

    return truncated_column
```

```python
# This function shortens the role information from the XML file.
def shorten_role(original_column) :
    length_to_cut = len('Role=')
    length_total = len(original_column)
    truncated_column = original_column[length_to_cut: length_total]

    return truncated_column
```

The XML file was iterated through using <b>ElementTree Library</b> built-in functions, as shown here :

```python
# Iterate through XML file to filter and organize relevant information.
for account in root.iter('account') :
    id = account.get('id', default = None)

    for i in range(id_series.size) :
        # If ID is found in the ID series, execute code.
        if (id == id_series[i]) :
            # Add this unordered ID to the list.
            xml_ids.append(id)

            # Assign role column to have the role information.
            for attribute in account.iter('attribute') :
                if attribute.get('name') == 'Role' :
                    attributeValueRef_id = str()

                    for attributeValueRef in attribute.iter('attributeValueRef') :
                        # Remove the 'Role=' using the custom function before acquiring the information.
                        attributeValueRef_id += shorten_role(attributeValueRef.get('id')) + ' \n'
                    # Append each string to the list if the ID is found in the ID series.
                    role_column.append(attributeValueRef_id)
```

The ID and role columns were matched by iterating through the respective <b>DataFrame</b> and <b>List</b>.

```python
# Organize role column to match the ordered ID series.
for organized_index in range(id_series.size) :
    # Add an empty string to end of list to account for IDs with blank roles.
    organized_role_column.append('')
    for unorganized_index in range(len(xml_ids)) :
        if (id_series[organized_index] == xml_ids[unorganized_index]) :
            # Replace empty string for IDs with roles in the XML file.
            organized_role_column.insert(organized_index, (role_column[unorganized_index]))
```

There is also a <a href = "test_script.py"><b>Python Script</b></a> written with additional comments to further understand the procedure of developing this organizer.

<p align = "center"><i>Note : This Exploration Took a Weekend to Complete, Spanning Approximately 10 Hours Altogether.</i></p>