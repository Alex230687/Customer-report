# -- coding: utf-8
import time
from program_object import Program
from settings import sql_settings


def timeit(func):
    def wrapper(*args, **kwargs):
        ts = time.perf_counter()
        func(*args, **kwargs)
        print(time.perf_counter() - ts)
    return wrapper


@timeit
def main():
    # date_start = input('Дача начала отчета (дд.мм.гггг): ')
    # date_end = input('Дата окончания отчета (дд.мм.гггг): ')
    date_start = '01.01.2020'
    date_end = '06.01.2020'

    prog = Program(date_start, date_end, sql_settings)
    prog.run_program()
    prog.update_sql_table()
    prog.close_program()


main()
