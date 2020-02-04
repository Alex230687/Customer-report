# -- coding: utf-8
import re
from sub_modul import (Sql, Validation, get_report_data as report,
    create_ignore_list as ignore, create_comparison_dict as comparison,
    open_research_file as research, date)


class Program():
    def __init__(self, date_start, date_end, sql_settings):
        self.sql = Sql(sql_settings)
        self.report_data = report()
        self.ignore = ignore()
        self.comparison = comparison()
        self.research = research()
        self.valid = Validation()
        self.date_string = date(date_start, date_end)
        self.sql_load_data = []

    def run_program(self):
        period_id = self.sql.get_period_id()
        for elem in self.report_data:
            isbn = re.sub('\D', '',
                self.valid.comparison_check(elem['isbn'], self.comparison))
            isbn_p = re.sub('\D', '', elem['isbn_p'])
            if isbn and isbn != '0' and len(isbn) >= 13:
                if isbn_p and isbn_p != '0':
                    if isbn_p == isbn:
                        pass
                    else:
                        if isbn in self.ignore:
                            isbn_p == isbn
                        else:
                            if self.valid.symbol_check(isbn, isbn_p):
                                self.research.write(isbn + '\n')
                            isbn_p = isbn
                isbn_p == isbn
            elif 10 <= len(isbn) < 13:
                self.research.write(isbn + '\n')
            self.sql_load_data.append((elem['title'], elem['sales'],
                                isbn, isbn_p, self.date_string, period_id))

    def update_sql_table(self):
        self.sql.insert_data(self.sql_load_data)

    def close_program(self):
        self.sql.close_sql()
        self.research.close()
