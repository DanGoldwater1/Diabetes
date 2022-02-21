#%%
from fileinput import filename
import gptables as gpt
from numpy import outer
import pandas as pd
from pathlib import Path


#First we set some parameters:
source_data = 'Book1.csv'
output_file_name = 'Diabetes_data_gptables.xlsx'
theme_to_use = 'diabetes_theme.yml'

with open('about.txt', 'r') as file:
    about = file.read()
    about = about.splitlines()

# about = ['OK']
print('foo')


# about = ['some',
#     {'bold': True},
#     'This is the text'
# ]

# from info import about 

parent_dir = Path(__file__).parent

#Now we import our CSV data
diabetes_df = pd.read_csv(source_data)
diabetes_df.head()


######### PROCESSING BLOCK
# This is an example, so we're not doing much with the data. But this is where you would call all your data analysis and manipulation functions in. 
##########

# In our output Excel file, we will want two sheets; one for data pertaining to type 1 diabetes, and one for type 2. So let's go ahead and create those. 

diabetes_type_1_df = diabetes_df.loc[diabetes_df['Type']==1]
diabetes_type_1_df = diabetes_type_1_df.drop(columns='Type')

diabetes_type_2_df = diabetes_df.loc[diabetes_df['Type']==2]
diabetes_type_2_df = diabetes_type_2_df.drop(columns='Type')


# Now we create the metadata for each of our sheets

type_1_kwargs = {
    "table_name": 'Type_1_Table_Data',
    # 'title':["Mean", {"italic": True}, " Iris", "$$note2$$ sepal dimensions"],
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
diabetes_type_1_df = diabetes_type_1_df.drop(columns='index')

diabetes_type_1_table = gpt.GPTable(table=diabetes_type_1_df, **type_1_kwargs)
diabetes_type_2_df = diabetes_type_2_df.reset_index()
diabetes_type_2_table = gpt.GPTable(table=diabetes_type_2_df, **type_2_kwargs)

##### COVER
cover = gpt.Cover(
    cover_label="Notes",
    title="A Worbook containing good practice tables",
    intro=["This is some introductory information", "And some more"],
    about=about,
    contact=["John Doe", "Tel: 345345345"],
    additional_elements=["subtitles", "scope", "source", "notes"]
    )


#Now we're ready to write these tables to a new excel file

if __name__=="__main__":
    output_path = parent_dir / output_file_name
    # theme_path = parent_dir / theme_to_use
    gpt.write_workbook(

        filename=output_path,
        sheets={
            "Type 1 Data": diabetes_type_1_table, #To the left of the colon sets the sheet name
            "Type 2 Data": diabetes_type_2_table
        },
        cover=cover,
        auto_width=True,
        theme=gpt.Theme(theme_to_use)
    )
    print(f'Output written to {output_file_name}')






