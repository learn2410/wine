import pandas
from pprint import pprint

g = pandas.read_excel('wine2.xlsx', sheet_name='Лист1').fillna('').groupby('Категория')
wines = {n: g.get_group(n).to_dict(orient='records') for n in g.groups.keys()}
pprint(wines)
