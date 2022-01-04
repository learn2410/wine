import pandas
from pprint import pprint
import collections

df = pandas.read_excel('wine2.xlsx', sheet_name='Лист1').fillna('').to_dict(orient='records')
wines = collections.defaultdict(list)
for w in df:
    wines[w['Категория']] += [w]
pprint(wines)
