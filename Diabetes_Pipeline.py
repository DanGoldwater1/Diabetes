#%%
from fileinput import filename
import gptables as gpt
from numpy import outer
import pandas as pd
from pathlib import Path
from datetime import date




#First we set some parameters:
source_data = 'Diabetes_Source.csv'
output_file_name = 'Diabetes_data_gptables.xlsx'
parent_dir = Path(__file__).parent

#Here we select our theme. This is a simple file and easy to edit, if we want to change the appearance of our Excel file. For example, I have set the title font color to blue. 
theme_to_use = 'diabetes_theme.yml'


#Next we want to import the frontmatter which will be presented on the first sheet of the excel book. This is a lot of text, so we've stored it in 'about.txt'
with open('about.txt', 'r') as file:
    about = file.read()
    about = about.splitlines() #This helps with the formatting


#Now we import our CSV data
diabetes_df = pd.read_csv(source_data)



######### PROCESSING BLOCK
# This is an example, so we're not doing much with the data. But this is where you would call all your data analysis and manipulation functions in. 
##########

# In our output Excel file, we will want two sheets; one for data pertaining to type 1 diabetes, and one for type 2. So let's go ahead and create those. 

diabetes_type_1_df = diabetes_df.loc[diabetes_df['Type']==1]
diabetes_type_1_df = diabetes_type_1_df.drop(columns='Type') #We don't need the info about diabetes type anymore

diabetes_type_2_df = diabetes_df.loc[diabetes_df['Type']==2]
diabetes_type_2_df = diabetes_type_2_df.drop(columns='Type')


# Now we create the metadata for each of our sheets. We'll use a dictionary for this.
type_1_kwargs = {
    "table_name": 'Type_1_Table_Data',
    "title": 'Diabetes Data',
    "subtitles": ['Type 1 Diabetes',],
    "units": None,
    "scope": '',
    "source": 'National Diabetes Study',
    "index_columns": {},
    "annotations": {
            },
    "notes": [
        'This is made as part of an example project, demonstraing gptables.',
        ]
    }

# Let's re-cycle and edit this for our second sheet
type_2_kwargs = type_1_kwargs.copy()
type_2_kwargs.update({
    "table_name": "Type_2_Table_Data",
    "title": "Diabetes Data",
    "subtitles": ['Type 2 Diabetes']
    })



#Next, gptables needs to convert our data from a pandas format to a format excel is happy with. 
diabetes_type_1_df = diabetes_type_1_df.reset_index()
diabetes_type_1_table = gpt.GPTable(table=diabetes_type_1_df, **type_1_kwargs)
diabetes_type_2_df = diabetes_type_2_df.reset_index()
diabetes_type_2_table = gpt.GPTable(table=diabetes_type_2_df, **type_2_kwargs)

##### COVER
#Here we create the cover sheet for the excel file. 
cover = gpt.Cover(
    cover_label="Notes",
    title="National Diabetes Audit (NDA) 2021-22 quarterly report for England, Clinical Commissioning Groups and GP practices - PROVISIONAL",
    intro=[f"Publication Date: {date.today()}", ], #This appears as a sub-heading
    about=about, #This is where we pull in all the information from our text file. 
    contact=["Author Name", "Tel: -----"],
    additional_elements=["subtitles", "scope", "source", "notes"] 
    )


#Now we're ready to write these tables to a new excel file

if __name__=="__main__":
    output_path = parent_dir / output_file_name
    # theme_path = parent_dir / theme_to_use
    gpt.write_workbook(

        filename=output_path,
        sheets={
            #To the left of the colon sets the sheet name. To the right of the colon sets the sheet content. Here we are assigning the tables we created above - not the pandas dataframes!
            "Type 1 Data": diabetes_type_1_table, 
            "Type 2 Data": diabetes_type_2_table
        },
        cover=cover,
        auto_width=True,
        theme=gpt.Theme(theme_to_use) #Selecting our theme from the top
    )
    print(f'Output written to {output_file_name}')






