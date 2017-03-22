# -*- coding: utf-8 -*-

import os
import pyodbc
import pandas as pd

from comtypes.client import CreateObject
from configobj import ConfigObj
from datetime import datetime
from itertools import zip_longest
from mailmerge import MailMerge


class Record(dict):
    @staticmethod
    def ROC_date(year, month=None, day=None):
        return '{}年{}月{}日'.format(
            year - 1911,
            month or '   ',
            day or '   ',
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
        self.doc_serial = self.format('#{案件編號:d}')
        if self['說明文件'] in ['簽文']:
            self.header = '創 先簽後稿'
        else:
            self.header = '創 以稿代簽'

        self.archive_code = self.format('{發文年份:d}/020411')
        self.archive_years = '3'
        self.doc_date = self.format('中華民國{民國發文日期:s}')

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


class WordMerge(object):
    DRAFT = 0
    FORMAL = 1

    def __init__(self, template_path, output_dir):
        self.template_path = template_path
        self.output_dir = output_dir

    def merge_records(self, records, format):
        record_batches = list(zip_longest(*[iter(records)] * 2))
        num_batches = len(record_batches)

        mergefields = []
        for (num_batch, record_batch) in enumerate(record_batches):
            if record_batch[1] is None:
                num_people = 1
                mergefield_base = {
                    'FIELD_0': [
                        {'FIELD_0': record_batch[0].title},
                    ],
                    'FIELD_10': [
                        {'FIELD_10': '一、', 'FIELD_11': record_batch[0].position},
                        {'FIELD_10': '二、', 'FIELD_11': record_batch[0].result},
                        {'FIELD_10': '三、', 'FIELD_11': record_batch[0].subject},
                        {'FIELD_10': '四、', 'FIELD_11': record_batch[0].rule},
                        {'FIELD_10': '五、', 'FIELD_11': record_batch[0].other},
                    ],
                    'FIELD_20': [],
                    'FIELD_30': [],
                    'FIELD_40': [],
                }
            else:
                if record_batch[0]['姓名'] == record_batch[1]['姓名']:
                    num_people = 1
                else:
                    num_people = 2

                mergefield_base = {
                    'FIELD_0': [],
                    'FIELD_10': [
                        {'FIELD_10': '一、', 'FIELD_11': record_batch[0].title},
                    ],
                    'FIELD_20': [
                        {'FIELD_20': '(一)', 'FIELD_21': record_batch[0].position},
                        {'FIELD_20': '(二)', 'FIELD_21': record_batch[0].result},
                        {'FIELD_20': '(三)', 'FIELD_21': record_batch[0].subject},
                        {'FIELD_20': '(四)', 'FIELD_21': record_batch[0].rule},
                        {'FIELD_20': '(五)', 'FIELD_21': record_batch[0].other},
                    ],
                    'FIELD_30': [
                        {'FIELD_30': '二、', 'FIELD_31': record_batch[1].title},
                    ],
                    'FIELD_40': [
                        {'FIELD_40': '(一)', 'FIELD_41': record_batch[1].position},
                        {'FIELD_40': '(二)', 'FIELD_41': record_batch[1].result},
                        {'FIELD_40': '(三)', 'FIELD_41': record_batch[1].subject},
                        {'FIELD_40': '(四)', 'FIELD_41': record_batch[1].rule},
                        {'FIELD_40': '(五)', 'FIELD_41': record_batch[1].other},
                    ],
                }

            if num_people == 1:
                representation = record_batch[0]['姓名'] + '1員'
            elif num_people == 2:
                representation = record_batch[0]['姓名'] + '等2員'
            else:
                representation = None

            mergefield_base.update({
                'DOC_DATE': record_batch[0].doc_date,
                'DOC_NUMBER': record_batch[0].doc_number,
                'REPRESENTATION': representation,
                'NOTE': record_batch[0].note,
                'NUM_PAGE': str(num_batch + 1),
                'NUM_PAGES': str(num_batches),
            })

            if format == WordMerge.DRAFT:
                mergefield = mergefield_base.copy()
                mergefield.update({
                    'DOC_SERIAL': record_batch[0].doc_serial,
                    'HEADER': record_batch[0].header,
                    'ARCHIVE_CODE': record_batch[0].archive_code,
                    'ARCHIVE_YEARS': record_batch[0].archive_years,
                    'DRAFT': '（稿）',
                    'RECIPIENT': '如正本',
                })

                if num_batch == len(record_batches) - 1:
                    mergefield.update({
                        'FOOTER_0': '大隊長 李ＯＯ',
                        'FOOTER_10': [
                            {
                                'FOOTER_10': '第一層決行',
                                'FOOTER_11': '',
                                'FOOTER_12': '',
                            },
                            {
                                'FOOTER_10': '承辦單位',
                                'FOOTER_11': '核稿',
                                'FOOTER_12': '批示',
                            },
                            {
                                'FOOTER_10': '擬：稿擬發。',
                                'FOOTER_11': '',
                                'FOOTER_12': '',
                            },
                        ],
                    })
                else:
                    mergefield.update({
                        'FOOTER_0': [],
                        'FOOTER_10': [],
                    })

                mergefields.append(mergefield)

            elif format == WordMerge.FORMAL:
                for num_person in range(num_people):
                    record = record_batch[num_person]
                    mergefield = mergefield_base.copy()
                    mergefield.update({
                        'HEADER': '',
                        'ARCHIVE_CODE': '',
                        'ARCHIVE_YEARS': '',
                        'DRAFT': '',
                        'RECIPIENT': record['姓名'],
                    })
                    mergefields.append(mergefield)

        path_prefix = os.path.join(
            self.output_dir,
            '{:d}_{:s}'.format(
                records[0]['案件編號'],
                records[0]['事由'],
            ),
        )
        word_path = '{:s}.docx'.format(path_prefix)
        pdf_path = '{:s}.pdf'.format(path_prefix)

        word_template = MailMerge(self.template_path)
        word_template.merge_pages(mergefields)
        word_template.write(word_path)

        word_document = word_app.Documents.open(word_path)
        word_document.SaveAs(pdf_path, FileFormat=17)  # magic 17

        print(path_prefix)


if __name__ == '__main__':
    config = ConfigObj('config.cfg')
    if not os.path.isdir(config['輸出資料夾']):
        os.makedirs(config['輸出資料夾'])
    word_app = CreateObject('Word.Application')

    connection = pyodbc.connect(
        driver='{Microsoft Access Driver (*.mdb, *.accdb)}',
        dbq=config['資料庫'],
    )

    df = pd.read_sql(
        (
            'SELECT * FROM 列印查詢 '
            'WHERE 案件編號 BETWEEN ? AND ?'
        ),
        params=[
            config['案件編號'][0],
            config['案件編號'][1],
        ],
        con=connection,
        index_col='識別碼',
        parse_dates=True,
    )

    cases = df.groupby('案件編號').groups
    for (case_key, case_indices) in cases.items():
        case_df = df.loc[case_indices]
        case_records = [Record(row) for (index, row) in case_df.iterrows()]

        document = WordMerge(
            template_path=config['模板'],
            output_dir=config['輸出資料夾'],
        )
        document.merge_records(case_records, WordMerge.DRAFT)

    word_app.Quit()