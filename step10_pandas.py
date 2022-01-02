import pandas

excel_data_df = pandas.read_excel('wine.xlsx', sheet_name='Лист1')

# print whole sheet data
print(excel_data_df.to_dict(orient='records'))
