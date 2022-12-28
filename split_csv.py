import csv


def csv_reader():
    '''
    Функция считывет исходный csv файл, возваращет данные из файла, заголовки файла и словарь с годами
    :return: data: массив с данными из файла,
    name: массив с заголовками файла,
    years: словарь с годами
    '''
    with open(input('Введите название файла: '), encoding='utf-8-sig') as r_file:
        file = csv.reader(r_file, delimiter=",")
        count = 0
        data = []
        name = []
        years = {}
        for line in file:
            if count == 0:
                lenLine = len(line)
                count += 1
                name = line
            else:
                if len(line) != lenLine or '' in line:
                    continue
                data.append(line)
                year = line[-1][0:4]
                if year not in years.keys():
                    years[year] = 1
                else:
                    years[year] += 1
        return data, name, years

data, name, years = csv_reader()

def get_csv_by_years(data, name, years):
    '''
    Функия разбивает исходный файл на мелкие файлы по годам
    :param data: массив с данными из файла
    :param name: массив с заголовками файла
    :param years: словарь с годами
    После работы функции в папку 'split' сохраняются файлы по годам
    '''
    for year in years.keys():
        with open(f'split/data_{year}.csv', 'w', encoding='utf-8-sig') as f:
            writer = csv.writer(f)
            writer.writerow(name)
            for row in data:
                if row[-1][0:4] == year:
                    writer.writerow(row)

get_csv_by_years(data, name, years)