# db_config.py
import pymysql
from pymysql.err import OperationalError
import os

# 数据库配置
DB_CONFIG = {
    'host': '192.168.1.140',
    'port': 3306,
    'user': 'printer_user',
    'password': 'printer_user',
    'database': 'printer_logs',
    'charset': 'utf8mb4'
}

TABLE_NAME = 'printer_logs'