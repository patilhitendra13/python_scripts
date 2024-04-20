import openpyxl # pip install openpyxl
import snowflake.connector # pip install snowflake-connector-python
import pandas # pip install pandas
# pip install "snowflake-connector-python[pandas]"

from openpyxl.styles import Border, Side , PatternFill

connection = snowflake.connector.connect(
    user='patilhitendra13',
    password='Patilhitendra13',
    account='kf94287.ap-south-1.aws',
    warehouse='COMPUTE_WH',
    database='Hiten',
    schema='test'
)

result_sheet_path=r"C:\Users\patil\OneDrive\Documents\Demo_sheet.xlsx"

conn = connection.cursor()

border = Border(top=Side(style='thin'),
                bottom=Side(style='thin'),
                left=Side(style='thin'),
                right=Side(style='thin'))

fill = PatternFill(start_color='40bce6',  # Yellow color
                   fill_type='solid')

# Load the workbook
wb = openpyxl.load_workbook(result_sheet_path)

## write values in cell
# # Select the worksheet
sheet = wb['Sheet2']  # Replace 'Sheet1' with the name of your sheet

# # Write to a specific cell
# sheet['d8'] = 'Hello, New World!'  # Write 'Hello, World!' to cell A1

## to write values of list in same column by adding new rows
# values = ['Value1', 'Value2', 'Value3', 'Value4', 'Value5']

# column = 'D'  # Column D
# start_row = 5

# for value in values:
#     # Find the first empty cell in the specified column
#     cell = sheet[column + str(start_row)]
#     sheet.insert_rows(start_row)
#     sheet[column + str(start_row)] = value
#     start_row += 1


# write function in a cell

# sheet['c5']=10
# sheet['c6']=10
# sheet['c8']='=exact(c5,c6)'

table_name = "hiten.test.dup_test"

query = f"select column_name,data_type,coalesce(character_maximum_length,numeric_precision) character_maximum_length from hiten.information_schema.columns where lower(table_schema)='{table_name.split('.')[1].lower()}' and lower(table_name)='{table_name.split('.')[2].lower()}'"


conn.execute(query)

op = conn.fetch_pandas_all()

columns = op.columns.to_list()

starting_cell = ['d',5]

for n,col in enumerate(columns):
    curr_row = starting_cell[1]
    sheet[starting_cell[0]+str(curr_row)]=col
    sheet[starting_cell[0]+str(curr_row)].border=border
    sheet[starting_cell[0]+str(curr_row)].fill=fill
    curr_row+=1
    if n==0:
        sheet.insert_rows(curr_row)
    
    for val in op[col]:
        sheet[starting_cell[0]+str(curr_row)]=val
        sheet[starting_cell[0]+str(curr_row)].border=border
        curr_row+=1
        if n==0:
            sheet.insert_rows(curr_row)

    starting_cell[0]=chr(ord(starting_cell[0])+1)

# Save the workbook
wb.save(result_sheet_path)
