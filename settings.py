# coding=utf-8
import os
import pytz
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
PATH_DB = os.path.join(BASE_DIR, '/timeguardian/db/TG.FDB')
PATH_BOOK = os.path.join(BASE_DIR, '/timeguardian/report/RMEN.xls')
TIME_ZONE = pytz.timezone('America/Chicago')
HOST = '127.0.0.1'
USER = 'SYSDBA'
PASS = "masterkey"