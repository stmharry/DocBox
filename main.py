# -*- coding: utf-8 -*-

import os
import pyodbc
import pandas as pd

from configobj import ConfigObj
from datetime import datetime
from itertools import zip_longest
from mailmerge import MailMerge

class Record(dict):
    @staticmethod
    def ROC_date(year, month=None, day=None):
        return '中華民國{}年{}月{}日'.format(
            year - 1911,
            month or '  ',
            day or '  ',
        )

    def __init__(self, *args, **kwargs):
        super(Record, self).__init__(*args, **kwargs)

        if self['發文日期'] is None:
            year = datetime.now().year
            month = None
            day = None
        else:
            year = self['發文日期'].year
            month = self['發文日期'].month
            day = self['發文日期'].day

        self['發文年份'] = year - 1911
        self['民國發文日期'] = Record.ROC_date(
            year=year,
            month=month,
            day=day,
        )

        self['民國說明日期'] = Record.ROC_date(
            year=self['說明日期'].year,
            month=self['說明日期'].month,
            day=self['說明日期'].day,
        )
        self['說明文件號'] = self['說明文件號'] or ''

        #
        if self['說明文件'] in ['簽文']:
            self.header = '創 先簽後稿'
        else:
            self.header = '創 以稿代簽'

        self.archive_code = self.format('{發文年份:d}/020411')
        self.archive_years = '3'
        self.doc_date = self.format('{民國發文日期:s}')

        if self['發文號'] == 0:
            self.doc_number = self.format('保七三大人字第{發文年份:d}000     號')
        else:
            self.doc_number = self.format('保七三大人字第{發文年份:d}{發文號:07d}號')

        self.title = self.format('{姓名:s}（{身分證字號:s}）')
        self.position = self.format('現職：{單位:s}（{單位代碼:s}），{職稱:s}（{職稱代碼:s}），{官等:s}。')
        self.result = self.format('獎懲：{結果:s}（{結果代碼:s}）。')
        self.subject = self.format('獎懲事由：{事由:s}（{事由代碼:s}）。')
        self.rule = self.format('法令依據：{法令:s}。')
        self.other = '其他事項：※'
        self.note = self.format('依據{說明單位:s}{民國說明日期:s}{說明文件號:s}{說明文件:s}辦理。')

    def format(self, s):
        return s.format(**self)


def get_field_dict(records):
    field_dicts = []
    for (num, (record0, record1)) in enumerate(records):
        if record1 is None:
            representation = record0['姓名'] + '1員'
            field_dict = {
                'FIELD_0': [
                    {'FIELD_0': record0.title},
                ],
                'FIELD_10': [
                    {'FIELD_10': '一、', 'FIELD_11': record0.position},
                    {'FIELD_10': '二、', 'FIELD_11': record0.result},
                    {'FIELD_10': '三、', 'FIELD_11': record0.subject},
                    {'FIELD_10': '四、', 'FIELD_11': record0.rule},
                    {'FIELD_10': '五、', 'FIELD_11': record0.other},
                ],
                'FIELD_20': [],
                'FIELD_30': [],
                'FIELD_40': [],
            }
        else:
            if record0['姓名'] == record1['姓名']:
                representation = record0['姓名'] + '1員'
            else:
                representation = record0['姓名'] + '等2員'

            field_dict = {
                'FIELD_0': [],
                'FIELD_10': [
                    {'FIELD_10': '一、', 'FIELD_11': record0.title},
                ],
                'FIELD_20': [
                    {'FIELD_20': '(一)', 'FIELD_21': record0.position},
                    {'FIELD_20': '(二)', 'FIELD_21': record0.result},
                    {'FIELD_20': '(三)', 'FIELD_21': record0.subject},
                    {'FIELD_20': '(四)', 'FIELD_21': record0.rule},
                    {'FIELD_20': '(五)', 'FIELD_21': record0.other},
                ],
                'FIELD_30': [
                    {'FIELD_30': '二、', 'FIELD_31': record1.title},
                ],
                'FIELD_40': [
                    {'FIELD_40': '(一)', 'FIELD_41': record1.position},
                    {'FIELD_40': '(二)', 'FIELD_41': record1.result},
                    {'FIELD_40': '(三)', 'FIELD_41': record1.subject},
                    {'FIELD_40': '(四)', 'FIELD_41': record1.rule},
                    {'FIELD_40': '(五)', 'FIELD_41': record1.other},
                ],
            }

        field_dict.update({
            'HEADER': record0.header,
            'ARCHIVE_CODE': record0.archive_code,
            'ARCHIVE_YEARS': record0.archive_years,
            'DOC_DATE': record0.doc_date,
            'DOC_NUMBER': record0.doc_number,
            'REPRESENTATION': representation,
            'NOTE': record0.note,
            'NUM_PAGE': num,
        })

        if num == len(records) - 1:
            field_dict.update({
                'FOOTER_0': [
                    {'FOOTER_0': '第一層決行', 'FOOTER_1': '', 'FOOTER_2': ''},
                    {'FOOTER_0': '承辦單位', 'FOOTER_1': '核稿', 'FOOTER_2': '批示'},
                    {'FOOTER_0': '擬：稿擬發。', 'FOOTER_1': '', 'FOOTER_2': ''},
                ],
            })
        else:
            field_dict.update({
                'FOOTER_0': [],
            })

        field_dicts.append(field_dict)

    return field_dicts

config = ConfigObj('config.cfg')
if not os.path.isdir(config['輸出資料夾']):
    os.makedirs(config['輸出資料夾'])

connection = pyodbc.connect(
    driver='{Microsoft Access Driver (*.mdb, *.accdb)}',
    dbq=config['資料庫'],
)

df = pd.read_sql(
    'select * from 列印查詢',
    con=connection,
    index_col='識別碼',
    parse_dates=True,
)

cases = df.groupby('案件編號').groups
for (case_key, case_indices) in cases.items():
    case_df = df.loc[case_indices]
    case_records = [Record(row) for (index, row) in case_df.iterrows()]

    doc_records = list(zip_longest(*[iter(case_records)] * 2))
    field_dicts = get_field_dict(doc_records)

    document = MailMerge('doc/template.docx')
    document.merge_pages(field_dicts)
    document_path = os.path.join(
        config['輸出資料夾'],
        '{:d}_{:s}.docx'.format(
            case_key,
            case_records[0]['事由'],
        ),
    )
    document.write(document_path)

    print(document_path)
