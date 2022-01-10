from datetime import datetime
from http.server import HTTPServer, SimpleHTTPRequestHandler

import pandas
from jinja2 import Environment, FileSystemLoader, select_autoescape

WINES_XLSX_PATH = 'wine.xlsx'


def get_firm_age():
    foundation_year = 1920
    age = str(datetime.now().year - foundation_year)
    years = 'год' if age[-1] == '1' else 'года' if '2' <= age[-1] <= '4' else 'лет'
    return f'{age} {years}'


def get_group_wines(xlsx_path):
    wines = pandas.read_excel(xlsx_path, sheet_name='Лист1').fillna('').groupby('Категория')
    # цель сортировки - собрать категории с винами в начало словаря
    kind_wines = sorted(list(wines.groups), key=lambda k: ' ' + k if ' вина' in k else k)
    return {kind: wines.get_group(kind).to_dict(orient='records') for kind in kind_wines}


def main():
    env = Environment(
        loader=FileSystemLoader('.'),
        autoescape=select_autoescape(['html', 'xml'])
    )

    template = env.get_template('template.html')
    rendered_page = template.render(
        firm_age=get_firm_age(),
        wines=get_group_wines(WINES_XLSX_PATH),
    )

    with open('index.html', 'w', encoding="utf8") as file:
        file.write(rendered_page)

    server = HTTPServer(('0.0.0.0', 8000), SimpleHTTPRequestHandler)
    server.serve_forever()


if __name__ == '__main__':
    main()
