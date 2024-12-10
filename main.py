import pandas

from datetime import datetime

from http.server import HTTPServer, SimpleHTTPRequestHandler

from jinja2 import Environment, FileSystemLoader, select_autoescape
from collections import defaultdict


def years_words(year):
    if year % 10 == 1 and year % 100 != 11:
        return 'год'
    elif year % 10 in [2, 3, 4] and year % 100 not in [12, 13, 14]:
        return 'года'
    else:
        return 'лет'


def main():
    env = Environment(
        loader=FileSystemLoader('.'),
        autoescape=select_autoescape(['html', 'xml'])
    )

    how_old = datetime.now().year - 1920

    new_excel = pandas.read_excel('wine.xlsx', keep_default_na=False)
    new_excel_dict = new_excel.to_dict(orient='records')

    wines = defaultdict(list)

    for i in new_excel_dict:
        category = i['Категория']

        wine_dict = {
            'Картинка': i['Картинка'],
            'Категория': category,
            'Название': i['Название'],
            'Сорт': i['Сорт'],
            'Цена': i['Цена'],
            'Акция': i['Акция']

        }

        wines[category].append(wine_dict)

    categories = sorted(wines.keys())

    template = env.get_template('template.html')

    rendered_page = template.render(
        wines=wines,
        categories=categories,
        date_time=f'Уже {how_old} {years_words(how_old)} с вами'
    )

    with open('index.html', 'w', encoding="utf8") as file:
        file.write(rendered_page)

    server = HTTPServer(('127.0.0.1', 8000), SimpleHTTPRequestHandler)
    server.serve_forever()


if __name__ == '__main__':
    main()
