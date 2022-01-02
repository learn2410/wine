from http.server import HTTPServer, SimpleHTTPRequestHandler
from jinja2 import Environment, FileSystemLoader, select_autoescape
from datetime import datetime
import pandas

def get_firm_age():
    age=str(datetime.now().year - 1920)
    years='год' if age[-1] == '1' else 'года' if '2' <= age[-1] <= '4' else 'лет'
    return f'{age} {years}'

env = Environment(
    loader=FileSystemLoader('.'),
    autoescape=select_autoescape(['html', 'xml'])
)

template = env.get_template('template.html')
rendered_page = template.render(
    firm_age=get_firm_age(),
    wines=pandas.read_excel('wine.xlsx', sheet_name='Лист1').to_dict(orient='records'),
)

with open('index.html', 'w', encoding="utf8") as file:
    file.write(rendered_page)

server = HTTPServer(('0.0.0.0', 8000), SimpleHTTPRequestHandler)
server.serve_forever()
