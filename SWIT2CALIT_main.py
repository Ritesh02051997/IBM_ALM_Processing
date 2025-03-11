import pandas as pd
from os import path
import numpy as np
from datetime import date

class filter_excel :
    """
        Intends to filter the excel data.
        Input: ALM downloaded csv/xlsx file.
        Return: Returns initial data frame of downloaded csv file.
        Methods:
            exceltodf: Generates data frame df from csv/xlsx file.
    """
    def __init__(self, link, sheet=None) :
        self.link = link
        self.sheet = sheet

    def exceltodf(self) :
        """Returns data frame df from csv/xlsx file"""
        if path.exists(self.link) and self.link.endswith('.csv'):
            df = pd.read_csv(self.link, encoding='UTF-16 LE', sep="\t")
        else :
            df = pd.read_excel(self.link, sheet_name=self.sheet)
        return df

class mapping_data :
    """
        Intends to understand inputs from mapping file and process it to the data frame.
        Input: Mapping file.
        Return: Updated df.
        Methods:
            rename: Renames values from columns to desired values
            left: Crops data of Value1 by Value2 characters from left
            concat: Concats column values and strings as mentioned from Value1 to ValueN
            hyperlink: Creates a hyperlink of the baselink at value1 and concat value2
            process_main: Main method of this class. This function calls methods of class mapping_data based on the actions.

    """
    def __init__(self, mapping_file, alm_df) :
        self.mapping_file = mapping_file
        self.alm_df = alm_df

    def rename(self, mapping_df) :
        """Renames values from columns to desired values"""
        for _, row in mapping_df.iterrows():
            # Apply the change to the given column
            self.alm_df[row['Col_name']] = self.alm_df[row['Col_name']].replace(row['Change_from'], row['Change_to'])
        
    def left(self, col_name, value1, value2) :
        """Crops data of Value1 by Value2 characters from left"""
        self.alm_df[col_name] = self.alm_df[value1].str[:value2]  # Crop by taking the first 'Value2' characters
   
    def concat(self, col_name, mapping_df, row) :
        """Concats column values and strings as mentioned from Value1 to ValueN"""
        result = []
        # For each row, concatenate the values based on the mapping
        for _, data_row in self.alm_df.iterrows():
            concat_value = ''            
            # Loop through the Value columns from the mapping sheet
            for i in range(1, len(row) - 1):  # Skip 'Col_name' and 'Action' columns
                value = row[f'Value{i}']
                # If the value is a string (in quotes), add the literal string
                if value.startswith('"') and value.endswith('"'):
                    concat_value += value.strip('"')  # Remove quotes
                else:
                    # Otherwise, treat it as a column name from df
                    concat_value += str(data_row[value])  # Retrieve value from df
            # Append the result for each row
            result.append(concat_value)
        # Add the result as a new column in df
        self.alm_df[col_name] = result

    def hyperlink(self, col_name, value1, value2) :
        """Creates a hyperlink of the baselink at value1 and concat value2"""
        # Add a new column with hyperlinks
        self.alm_df[col_name] = self.alm_df[value2].apply(lambda x: f'=HYPERLINK("{value1}{x}", "{value1}{x}")')

    def process_main(self):
        """Main: This function calls methods of class mapping_data based on the actions."""
        self.mapping_df_rename = pd.read_excel(self.mapping_file, sheet_name="rename")
        self.mapping_add_cl = pd.read_excel(self.mapping_file, sheet_name="add_column")
        #For rename tab tasks
        self.rename(self.mapping_df_rename)  #Please mute this line in case renaming is no more needed.
        #For add_column tab tasks
        for _, row in self.mapping_add_cl.iterrows():
            action = row['Action']
            if action == 'LEFT':
                # Crop the df['Value1'] column by taking the first 'Value2' characters from the left
                self.left(row['Col_name'], row['Value1'], int(row['Value2']))
            elif action == 'CONCAT' :
                self.concat(row['Col_name'], self.mapping_add_cl, row)
            elif action == 'HYPERLINK' :
                self.hyperlink(row['Col_name'], row['Value1'], row['Value2'])
        # self.alm_df.to_excel('output.xlsx', index=False)
        return self.alm_df

class output_excel :
    """
        Intends to generate the final csv based on .
        Input: Mapping file.
        Methods:
            generate_main: Main function of this class. This function calls methods of class output_excel based on the actions.

    """
    def __init__(self, mpf, alm_df, output_dir):
        self.mpf = mpf
        self.alm_df = alm_df
        self.output_dir = output_dir

    def extract_match_keyword(self, summary, keywords):
        """Function to extract relevant keyword from Summary for matching. Extract substrings that likely indicate the work package."""
        summary = summary.replace('.', '')
        for keyword in keywords:
            if keyword in summary:
                return keyword
            elif keyword.split('__')[-1] in summary:
                return keyword.split('__')[-1]
            elif keyword.split('_')[-1] in summary:
                return keyword.split('_')[-1]
        return None

    def sheet(self, col_name, value) :
        """Processing Sheet information."""
        project_map = pd.read_excel(self.mpf, sheet_name="project_map")
        temp_project_map = pd.merge(self.alm_df['Project'], project_map, on='Project', how='left')
        sheet_data = {}
        for sheet in project_map['Sheet_name'] :
            sheet_data[sheet] = pd.read_excel(self.mpf, sheet_name=sheet)
        new_column=[]
        for idx, row in self.output_df.iterrows():
            sheet_name = temp_project_map['Sheet_name'].iloc[idx]
            # Retrieve the corresponding sheet DataFrame
            work_package_data = sheet_data[sheet_name]
            # Extract keyword from the summary for matching
            match_keyword = self.extract_match_keyword(row['Summary'], sheet_data[sheet_name]['WorkPackage'])
            if match_keyword:
                # Find the corresponding WorkPackage that contains the match_keyword
                matched_work_package = work_package_data[work_package_data['WorkPackage'].str.contains(match_keyword)]
                if not matched_work_package.empty:
                    # Get the value for the matched work package
                    owned_by = matched_work_package[value].values[0]
                    new_column.append(owned_by)
                else:
                    # If no match is found, append 'Unassigned'
                    new_column.append('Unassigned')
            else:
                # If no relevant keyword is found in the summary, append 'Unassigned'
                new_column.append('Unassigned')
        self.output_df[col_name] = new_column

    def generate_main(self) :
        """Main: This function calls methods of class output_excel based on the actions."""
        self.mpf_df = pd.read_excel(self.mpf, sheet_name="template_action")
        self.output_df  = pd.DataFrame(index=range(len(self.alm_df)))
        if True :
            #Default values
            output_excel_name = 'Import_'+str(date.today().strftime('%Y%m%d'))
        for _, row in self.mpf_df.iterrows():
            #Row wise action for columns
            action = row['Action']
            if action == 'NAME' :
                if row['Value1'] is not np.nan :
                    output_excel_name = row['Value1']
            elif action == "BLANK" :
                self.output_df[row['Col_name']] = ''
            elif action == "FIXED VALUE" :
                self.output_df[row['Col_name']] = row['Value1']
            elif action == "INPUT CSV" :
                self.output_df[row['Col_name']] = self.alm_df[row['Value1']]
            elif action == "SHEET" :
                self.sheet(row['Col_name'],row['Value1'])
        
        self.output_df.to_csv(path.join(self.output_dir,str(output_excel_name)+str('.csv')) , encoding='UTF-8' ,index=False)
        return str(output_excel_name)+str('.csv')
