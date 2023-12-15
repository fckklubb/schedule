from datetime import datetime, date, time
from dataclasses import dataclass
import pandas as PD
import numpy as NP
import openpyxl
from openpyxl import load_workbook
from openpyxl.workbook import Workbook

SCHEDULE_FILE = 'files\schedule_spb_2024.xlsx'

schedule5_2 = [5, 2]
schedule2_2_3 = [2, 2, 3]
schedule2_2 = [2, 2]

long_week = [True, True, False, False, True, True, True]
short_week = [False, False, True, True, False, False, False]
five_work_day = [True, True, True, True, True, False, False]

YY = 2024

@dataclass
class shift:
    day_one: date
    schedule: str
    workers_office: []
    workers_tech: []

holydays24 = {1: [date(2024, 1, i) for i in range(1,9)],
              2: [date(2024, 2, 23)],
              3: [date(2024, 3, 8)],
              4: [date(2024, 4, 29), date(2024, 4, 30)],
              5: [date(2024, 5, 1), date(2024, 5, 9), date(2024, 5, 10)],
              6: [date(2024, 6, 12)],
              11: [date(2024, 11, 4)],
              12: [date(2024, 12, 30), date(2024, 12, 31)]}
exceptions24 = [date(2024,4, 27), date(2024, 11, 2), date(2024, 12, 28)]

shift0 = shift(date(2024, 1, 9), 'five', ['Жуков Роман'], ['Филатов Алексей'])
shift1 = shift(date(2024, 1, 1), 'shift_work_1', ['Березкина Александра', 'Бобарыкин Глеб'], ['Церетели Кирилл'])
shift2 = shift(date(2024, 1, 3), 'shift_work_2', ['Смирнов Андрей', 'Ильченко Андрей'], ['Кочурин Роман'])

months = ['Январь', 'Февраль', 'Март', 'Апрель', 'Май', 'Июнь', 'Июль', 'Август', 'Сентябрь', 'Октябрь', 'Ноябрь', 'Декабрь']
m_days = [31, 29, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]

sheet_prefix = 'График/табель работы'

days = ['пн', 'вт', 'ср', 'чт', 'пт', 'сб', 'вс']

plan_day_symb = 'п'
sick_leave = 'б'
off_days = 'о'
off_day = 'г'

time_five_day_work = {'start': '9:00:00', 'end': '18:00:00', 'short_end': '17:00:00'}
time_shift_work = {'start': '9:00:00', 'end': '21:00:00', 'short_end': '21:00:00'}

def five_day_return(d):
    wd = d.isoweekday()
    try:
        holydays = holydays24.get(d.month)
        # print(holydays)
        if holydays != None and d in holydays: return ('', '', '')
    except KeyError:
        print('Error with dict')
    if wd in [1, 2, 3, 4, 5] or d in exceptions24:
        return (plan_day_symb, time_five_day_work.get('start'), time_five_day_work.get('end'))
    return ('', '', '')

def shift_day_return(d, shift_no):
    wd = d.isoweekday()
    s = d.isocalendar().week % 2
    if shift_no == 1:
        if s == 1: w = long_week
        else: w = short_week
    if shift_no == 2:
        if s == 1: w = short_week
        else: w = long_week
    if w[wd-1]: return (plan_day_symb, time_shift_work.get('start'), time_shift_work.get('end'))
    else: return ('', '', '')

@dataclass
class Person:
    name: str
    shift: int
    vocation: [date] = []
    str_plan: int = 0
    str_f_start: int = 0
    str_f_end: int = 0

def getVocationList (start_d: date, days: int) -> [date]:
    if days > 28: return None
    pass

def set_schedule (m, person, sched_type):
    pass

def find_person_lines (ws: openpyxl.worksheet.worksheet.Worksheet, p: Person, m: int):
    for i in range(1, ws.max_row+1):
        pass