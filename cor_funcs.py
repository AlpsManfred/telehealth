import datetime
import sqlite3
import pandas as pd
import openpyxl


def get_working_days_from_today():
    today = datetime.date.today()
    weekday = today.weekday()

    monday = today - datetime.timedelta(days=weekday)

    working_days = []
    for i in range(weekday, 5):
        working_days.append(monday + datetime.timedelta(days=i))

    for i in range(5):
        working_days.append(monday + datetime.timedelta(days=i + 7))

    return working_days


def get_tg_and_phone_number():
    workbook = openpyxl.load_workbook('telehealth.xlsx')
    worksheet = workbook["Лист1"]
    phone_number = worksheet["G2"].value
    tg = worksheet["H2"].value
    return tg, phone_number
