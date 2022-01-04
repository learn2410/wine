from http.server import HTTPServer, SimpleHTTPRequestHandler
from jinja2 import Environment, FileSystemLoader, select_autoescape
from datetime import datetime
import pandas
import collections

EXCEL_WINE = 'wine3.xlsx'

def get_firm_age():
    age = str(datetime.now().year - 1920)
    years = 'год' if age[-1] == '1' else 'года' if '2' <= age[-1] <= '4' else 'лет'
    return f'{age} {years}'


def get_group_wines(xlsx):
    df = pandas.read_excel(xlsx, sheet_name='Лист1').fillna('').to_dict(orient='records')
    wines = collections.defaultdict(list)
    grouplist = list(set(x['Категория'] for x in df))
    # move wines to top of list (the wine can be pink, sparkling e.t.c)
    grouplist.sort(key=lambda k: ' ' + k if ' вина' in k else k)
    for category in grouplist:
        wines[category] = []
    for w in df:
        wines[w['Категория']] += [w]
    return wines


env = Environment(
    loader=FileSystemLoader('.'),
    autoescape=select_autoescape(['html', 'xml'])
)

template = env.get_template('template.html')
rendered_page = template.render(
    firm_age=get_firm_age(),
    wines=get_group_wines(EXCEL_WINE),
)

with open('index.html', 'w', encoding="utf8") as file:
    file.write(rendered_page)

server = HTTPServer(('0.0.0.0', 8000), SimpleHTTPRequestHandler)
server.serve_forever()
