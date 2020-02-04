# -- coding: utf-8
import os, datetime, re
import pyodbc
import xlrd

################################################################################
# SQL OBJECT
class Sql():
    def __init__(self, sql_settings):
        self.connection = pyodbc.connect(**sql_settings)
        self.cursor = self.connection.cursor()
        self.table = '[dbo].[TestSA2]'

    def get_period_id(self):
        query = 'SELECT MAX([Номер периода]) AS mpid FROM %s' % self.table
        try:
            self.cursor.execute(query)
        except pyodbc.Error as err:
            print(err)
            return None
        else:
            period_id = int(self.cursor.fetchone().mpid) + 1
            return period_id

    def insert_data(self, data):
        query = 'INSERT INTO %s VALUES (?,?,?,?,?,?)' % self.table
        try:
            self.cursor.executemany(query, data)
        except pyodbc.Error as err:
            print(err)
        else:
            self.connection.commit()

    def close_sql(self):
        self.cursor.close()
        self.connection.close()


################################################################################
# VALIDATION OBJECT
class Validation():
    def __init__(self):
        pass

    def zero_check(self, isbn):
        if isbn == 0 or isbn == '0' or isbn == '' or isbn is None:
            return False
        return True

    def equal_check(self, isbn, isbn_p):
        if isbn == isbn_p:
            return True
        return False

    def symbol_check(self, isbn, isbn_p):
        counter = 0
        for isbn_index, isbn_p_index in zip(isbn, isbn_p):
            if isbn_index != isbn_p_index:
                counter += 1
        return ((len(isbn) - counter) / len(isbn)) > 0.7

    def comparison_check(self, isbn, dict):
        new_isbn = isbn
        if isbn in dict.keys():
            new_isbn = dict[isbn]
        return new_isbn

    def ignore_check(self, isbn, ignore):
        if isbn in ignore:
            return True
        return False


################################################################################
# EXCEL OBJECT
def open_excel():
    workbook = xlrd.open_workbook(os.path.join(os.getcwd(), 'files', 'report.xlsx'),
        encoding_override='utf8')
    worksheet = workbook.sheet_by_index(0)
    return worksheet


def set_columns(worksheet):
    """Define column number by its name and return dict {name: number}"""
    # Определеяет номера колонок по их имени и вовзращает словрь {имя: номер}
    columns = dict(dict.fromkeys(['Title', 'ISBN', 'ISBN_P', 'Sales']))
    for column in range(worksheet.ncols):
        if worksheet.cell(0, column).value in ('Название', 'Позиция'):
            columns['Title'] = column
        elif worksheet.cell(0, column).value in ('ISBN',):
            columns['ISBN_P'] = column
        elif worksheet.cell(0, column).value in ('IDext', 'Код',):
            columns['ISBN'] = column
        elif worksheet.cell(0, column).value in ('Кол-во',):
            columns['Sales'] = column
    if None in columns.values():
        print('Check new column name in customer report file!')
        return None
    return columns


def get_row_data(worksheet, columns):
    """Get data from report row by row and return row data list"""
    # Собирает данные построчно и возрващает список строк
    data_list = []
    for row in range(1, worksheet.nrows):
        title = worksheet.cell(row, columns['Title']).value
        isbn = worksheet.cell(row, columns['ISBN']).value
        isbn_p = worksheet.cell(row, columns['ISBN_P']).value
        sales = worksheet.cell(row, columns['Sales']).value
        if not sales:
            sales = 0
        data_list.append(
            {'title': title, 'isbn': isbn, 'isbn_p': isbn_p,
             'sales': sales}
        )
    return data_list


def get_report_data():
    worksheet = open_excel()
    columns = set_columns(worksheet)
    data = get_row_data(worksheet, columns)
    return data


################################################################################
# FILE OBJECT
def create_ignore_list():
    """Return isbn list from ignore.txt file"""
    # Возвращает список isbn из файла ignore.txt
    with open(os.path.join(os.getcwd(), 'files', 'ignore.txt'), 'r') as file:
        ignore_list = file.read().splitlines()
        return ignore_list


def create_comparison_dict():
    """Return isbn pairs dict from comparison.txt file"""
    # Возвращает словарь пар isbn из файла comparison.txt
    with open(os.path.join(os.getcwd(), 'files', 'comparison.txt'), 'r') as file:
        comparison_list = file.read().splitlines()
        for i in range(len(comparison_list)):
            comparison_list[i] = comparison_list[i].split(':')
        return dict(comparison_list)


def open_research_file():
    """Return file object to write data to research file"""
    # Возвращает объект файла research.txt для записи данных
    file = open(os.path.join(os.getcwd(), 'files', 'research.txt'), 'w')
    return file


################################################################################
# DATE OBJECT
def date(date_start, date_end):
    periods = ['day', 'month', 'year']
    # create date list >>> input '01/01/2020' >>> output ['01', '01', '2020']
    ds_list = re.sub('[\.,\\/:;=-]', ',', date_start).split(',')
    de_list = re.sub('[\.,\\/:;=-]', ',', date_end).split(',')
    # zip two lists (periods and date) and convert str(date elem) to int(date elem)
    # >>> input ['day', 'month', 'year'], ['01', '01', '2020']
    # >>> output {'day': 1, 'month': 1, 'year': 2020}
    ds_dict = dict(zip(periods, map(lambda x: int(x), ds_list)))
    de_dict = dict(zip(periods, map(lambda x: int(x), de_list)))
    # make date from dict >>> input {'day': 1, 'month': 1, 'year': 2020} >>> output '01.01.2020'
    ds_str_date = datetime.date(**ds_dict).strftime('%d.%m.%Y')
    de_str_date = datetime.date(**de_dict).strftime('%d.%m.%Y')
    # create final string for sql >>> input '01.01.2020', '10.01.2020' >>> output '01.01.2020 - 10.01.2020'
    sql_final_string = '{0} - {1}'.format(ds_str_date, de_str_date)
    return sql_final_string
