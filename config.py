import os
import json

_defaults = {
    'EXCEL_PATH': 'c:/ayumi/ayumi.xlsm',
    'SHEET_NAME': 'Sheet2',
    'JSON_PATH': 'c:/ayumi/data.json',
    'PROGRESS_PATH': 'c:/ayumi/progress.json',
    'EXCEL_OUTPUT_PATH': 'c:/ayumi/updated_data.xlsx',
    'CONSOLIDATED_OUTPUT_PATH': 'c:/ayumi/consolidated_data.xlsx',
    'MYSQL': {
        'host': '192.168.0.139',
        'database': 'test',
        'user': 'root',
        'password': 'testpass',
    },
}

config_file = os.getenv('AYUMI_CONFIG')
if config_file and os.path.exists(config_file):
    with open(config_file, 'r', encoding='utf-8') as f:
        file_conf = json.load(f)
    for k, v in file_conf.items():
        if k == 'MYSQL':
            _defaults['MYSQL'].update(v)
        else:
            _defaults[k] = v

EXCEL_PATH = os.getenv('EXCEL_PATH', _defaults['EXCEL_PATH'])
SHEET_NAME = os.getenv('SHEET_NAME', _defaults['SHEET_NAME'])
JSON_PATH = os.getenv('JSON_PATH', _defaults['JSON_PATH'])
PROGRESS_PATH = os.getenv('PROGRESS_PATH', _defaults['PROGRESS_PATH'])
EXCEL_OUTPUT_PATH = os.getenv('EXCEL_OUTPUT_PATH', _defaults['EXCEL_OUTPUT_PATH'])
CONSOLIDATED_OUTPUT_PATH = os.getenv('CONSOLIDATED_OUTPUT_PATH', _defaults['CONSOLIDATED_OUTPUT_PATH'])
MYSQL = {
    'host': os.getenv('MYSQL_HOST', _defaults['MYSQL']['host']),
    'database': os.getenv('MYSQL_DATABASE', _defaults['MYSQL']['database']),
    'user': os.getenv('MYSQL_USER', _defaults['MYSQL']['user']),
    'password': os.getenv('MYSQL_PASSWORD', _defaults['MYSQL']['password']),
}
