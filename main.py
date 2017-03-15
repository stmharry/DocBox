# -*- coding: utf-8 -*-

import itertools
import mailmerge
import pyodbc
import pandas as pd


class Record(dict):
    @staticmethod
    def ROC_date(date):
        return '中華民國{:d}年{:d}月{:d}日'.format(
            date.year - 1911,
            date.month,
            date.day,
        )

    def __init__(self, *args, **kwargs):
        super(Record, self).__init__(*args, **kwargs)

        self['發文年份'] = self['發文日期'].year
        self['民國發文日期'] = Record.ROC_date(self['發文日期'])
        self['民國說明日期'] = Record.ROC_date(self['說明日期'])
        self['說明文件號'] = self['說明文件號'] or ''

        self.archive_code = self.format('{發文年份:d}/020411')
        self.archive_year = '3'
        self.doc_date = self.format('{民國發文日期:s}')
        self.doc_number = self.format('保七三大人字第{發文年份:d}{發文號:07d}號')
        self.title = self.format('{姓名:s}（{身分證字號:s}）')
        self.position = self.format('現職：{單位:s}（{單位代碼:s}），{職稱:s}（{職稱代碼:s}），{官等:s}。')
        self.result = self.format('獎懲：{結果:s}（{結果代碼:s}）。')
        self.subject = self.format('獎懲事由：{事由:s}（{事由代碼:s}）。')
        self.rule = self.format('法令依據：警察人員獎懲標準第{條:s}條第{款:s}款。')
        self.other = '其他事項：※'
        self.footer = self.format('依據{說明單位:s}{民國說明日期:s}{說明文件號:s}{說明文件:s}辦理')

    def format(self, s):
        return s.format(**self)


def get_field_dict(record0, record1=None):
    if record1 is None:
        header = record0['姓名'] + '1員'
        field_dict = {
            'FIELD_0': [
                {'FIELD_0': record0['姓名']},
            ],
            'FIELD_10': [
                {'FIELD_10': '一、', 'FIELD_11': record0.position},
                {'FIELD_10': '二、', 'FIELD_11': record0.result},
                {'FIELD_10': '三、', 'FIELD_11': record0.subject},
                {'FIELD_10': '四、', 'FIELD_11': record0.rule},
                {'FIELD_10': '五、', 'FIELD_11': record0.other},
            ],
        }
    else:
        if record0['姓名'] == record1['姓名']:
            header = record0['姓名'] + '1員'
        else:
            header = record0['姓名'] + '等2員'

        field_dict = {
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
        'ARCHIVE_CODE': record0.archive_code,
        'ARCHIVE_YEAR': record0.archive_year,
        'DOC_DATE': record0.doc_date,
        'DOC_NUMBER': record0.doc_number,
        'HEADER': header,
        'FOOTER': record0.footer,
    })
    return field_dict


connection = pyodbc.connect(
    driver='{Microsoft Access Driver (*.mdb, *.accdb)}',
    dbq=r'\\隊本部收發\收發\公文附件\銘進\獎懲輸入\獎懲.accdb',
)
document = mailmerge.MailMerge('doc/doc_2+.docx')
print(document.get_merge_fields())

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

    doc_records = itertools.zip_longest(*[iter(case_records)] * 2)
    for (doc_record0, doc_record1) in doc_records:
        field_dict = get_field_dict(doc_record0, doc_record1)
        document.merge(**field_dict)
        document.write('doc/doc_2_out.docx')
