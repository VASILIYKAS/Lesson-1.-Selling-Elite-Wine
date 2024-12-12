import pandas
import argparse

from datetime import datetime

from http.server import HTTPServer, SimpleHTTPRequestHandler

from jinja2 import Environment, FileSystemLoader, select_autoescape
from collections import defaultdict


def get_year_word(year):
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

    parser = argparse.ArgumentParser(description='Программа для работы с файлами excel')
    parser.add_argument('--file_path', default='wine.xlsx', help='Путь к файлу')
    parser.add_argument('--template_path', default='template.html', help='Путь к шаблону HTML')
    parser.add_argument('--output_file', default='index.html', help='Путь куда сохраняется файл')
    args = parser.parse_args()

    new_excel = pandas.read_excel(args.file_path, keep_default_na=False)
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

    template = env.get_template(args.template_path)

    rendered_page = template.render(
        wines=wines,
        categories=categories,
        date_time=f'Уже {how_old} {get_year_word(how_old)} с вами'
    )

    with open(args.output_file, 'w', encoding="utf8") as file:
        file.write(rendered_page)

    server = HTTPServer(('127.0.0.1', 8000), SimpleHTTPRequestHandler)
    server.serve_forever()


if __name__ == '__main__':
    main()
