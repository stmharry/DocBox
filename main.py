# -*- coding: utf-8 -*-

import os
import pyodbc
import pandas as pd
import PyPDF2

from comtypes.client import CreateObject
from configobj import ConfigObj
from datetime import datetime
from itertools import zip_longest
from mailmerge import MailMerge
from natsort import natsorted


class Record(object):
    @staticmethod
    def ROC_date(year, month=None, day=None):
        return '{}年{}月{}日'.format(
            year - 1911,
            month or '   ',
            day or '   ',
        )

    def __init__(self, series):
        self.dict_ = dict(zip(series.index, series))

        if self.dict_['發文日期'] is pd.NaT or self.dict_['發文日期'] is None:
            year = datetime.now().year
            month = None
            day = None
        else:
            year = self.dict_['發文日期'].year
            month = self.dict_['發文日期'].month
            day = self.dict_['發文日期'].day

        self.dict_['發文年份'] = year - 1911
        self.dict_['民國發文日期'] = Record.ROC_date(
            year=year,
            month=month,
            day=day,
        )

        self.dict_['民國說明日期'] = Record.ROC_date(
            year=self.dict_['說明日期'].year,
            month=self.dict_['說明日期'].month,
            day=self.dict_['說明日期'].day,
        )
        self.dict_['說明文件號'] = self.dict_['說明文件號'] or ''

        #
        self.doc_serial = self.format('#{案件編號:d}')
        if self.dict_['說明文件'] in ['簽文']:
            self.header = '創 先簽後稿'
        else:
            self.header = '創 以稿代簽'

        self.archive_code = self.format('{發文年份:d}/020411')
        self.archive_years = '3'
        self.doc_date = self.format('中華民國{民國發文日期:s}')

        if self.dict_['發文號'] == 0:
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
        return s.format(**self.dict_)


class WordMerge(object):
    DRAFT = 0
    FORMAL = 1

    def __init__(self, template_path):
        self.template_path = template_path

    @staticmethod
    def get_mergefields(records, format):
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
                if record_batch[0].dict_['姓名'] == record_batch[1].dict_['姓名']:
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
                representation = record_batch[0].dict_['姓名'] + '1員'
            else:
                representation = record_batch[0].dict_['姓名'] + '等2員'

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
                recipients = set()
                for num_person in range(num_people):
                    record = record_batch[num_person]
                    # DEBUG # assert record.dict_['發文日期'] is not None and record.dict_['發文號'] != 0, '有未輸入發文資訊之紀錄！'

                    recipients.add(record.dict_['姓名'])
                    recipients.add(record.dict_['中隊'])

                for recipient in recipients:
                    mergefield = mergefield_base.copy()
                    mergefield.update({
                        'DOC_SERIAL_ALT': record_batch[0].doc_serial,
                        'HEADER': '',
                        'ARCHIVE_CODE': '',
                        'ARCHIVE_YEARS': '',
                        'DRAFT': '',
                        'RECIPIENT': recipient,
                    })
                    mergefields.append(mergefield)

        return pd.DataFrame(mergefields)

    def merge(self, mergefields, word_path):
        print('WordMerge.merge')
        print(word_path)  # DEBUG

        word_template = MailMerge(self.template_path)
        word_template.merge_pages([mergefield._asdict() for mergefield in mergefields.itertuples()])
        word_template.write(word_path)


class Manager(object):
    def __init__(self, config_path):
        self.config = ConfigObj(config_path)
        self.connection = pyodbc.connect(
            driver='{Microsoft Access Driver (*.mdb, *.accdb)}',
            dbq=self.config['資料庫'],
        )
        self.word_merge = WordMerge(
            template_path=self.config['模板'],
        )
        self.double_sided = DoubleSided()

        self.output_dir = self.config['輸出資料夾']
        if not os.path.isdir(self.output_dir):
            os.makedirs(self.output_dir)

        if self.config['格式'] == '草稿':
            self.format = WordMerge.DRAFT
        elif self.config['格式'] == '正本':
            self.format = WordMerge.FORMAL

    def merge(self):
        df = pd.read_sql(
            (
                'SELECT * FROM 列印查詢 '
                'WHERE 案件編號 BETWEEN ? AND ?'
            ),
            params=[
                self.config['案件編號範圍'][0],
                self.config['案件編號範圍'][1],
            ],
            con=self.connection,
            index_col='識別碼',
            parse_dates=True,
        )

        cases = df.groupby('案件編號').groups

        if self.format == WordMerge.DRAFT:
            filepaths = []
            for (case_key, case_indices) in cases.items():
                case_records = [
                    Record(series=df.loc[case_index])
                    for case_index in case_indices
                ]

                mergefields = WordMerge.get_mergefields(case_records, format=self.format)
                filepath = os.path.join(
                    self.output_dir,
                    '{:d}_{:s}.docx'.format(
                        case_records[0].dict_['案件編號'],
                        case_records[0].dict_['事由'],
                    ),
                )
                self.word_merge.merge(mergefields, word_path=filepath)
                filepaths.append(filepath)

            self.double_sided.merge(
                natsorted(filepaths),
                pdf_path=os.path.join(self.output_dir, '{:d}-{:d}.pdf'.format(min(cases.keys()), max(cases.keys()))),
            )

        elif self.format == WordMerge.FORMAL:
            all_merge_fields = pd.DataFrame()
            for (case_key, case_indices) in cases.items():
                case_records = [
                    Record(series=df.loc[case_index])
                    for case_index in case_indices
                ]

                mergefields = WordMerge.get_mergefields(case_records, format=self.format)
                all_merge_fields = all_merge_fields.append(mergefields, ignore_index=True)

            df_recipient = df[['姓名', '中隊']].drop_duplicates()
            indices_by_name = all_merge_fields.groupby('RECIPIENT').groups

            # tidying

            for team in df_recipient['中隊'].unique():
                mergefields = pd.DataFrame()
                for name in df_recipient.ix[df_recipient['中隊'] == team, '姓名']:
                    mergefields = mergefields.append(all_merge_fields.loc[indices_by_name[name]])

                filename = os.path.join(
                    self.output_dir,
                    '{:s}_個人'.format(team),
                )
                self.word_merge.merge(mergefields, word_path=filename + '.docx')
                self.double_sided.merge(
                    [filename + '.docx'],
                    pdf_path=filename + '.pdf',
                )

                filepaths = []
                team_merge_fields = all_merge_fields.loc[indices_by_name[team]]
                cases = team_merge_fields.groupby('DOC_SERIAL_ALT').groups
                for (case_key, case_indices) in cases.items():
                    mergefields = team_merge_fields.loc[case_indices]

                    filepath = os.path.join(
                        self.output_dir,
                        '{:s}_總表_{:s}.docx'.format(team, case_key),
                    )
                    self.word_merge.merge(mergefields, word_path=filepath)
                    filepaths.append(filepath)

                self.double_sided.merge(
                    natsorted(filepaths),
                    pdf_path=os.path.join(self.output_dir, '{:s}_總表.pdf'.format(team)),
                )


class DoubleSided(object):
    def __init__(self):
        self.word_app = CreateObject('word.application')

    def merge(self, word_paths, pdf_path):
        print('DoubleSided.merge')
        print(word_paths)
        print(pdf_path)

        if len(word_paths) == 1:
            word_path = word_paths[0]
            (basepath, _) = os.path.splitext(word_path)

            word_document = self.word_app.documents.open(word_path)
            word_document.saveas(pdf_path, FileFormat=17)  # magic 17 as pdf
            word_document.close()
            os.remove(word_path)
        else:
            writer = PyPDF2.PdfFileWriter()

            fs = {}
            for (num_path, word_path) in enumerate(word_paths):
                (basepath, _) = os.path.splitext(word_path)
                temp_pdf_path = '{:s}_temp.pdf'.format(basepath)

                word_document = self.word_app.documents.open(word_path)
                word_document.saveas(temp_pdf_path, FileFormat=17)  # magic 17 as pdf
                word_document.close()
                os.remove(word_path)

                f = open(temp_pdf_path, 'rb')
                reader = PyPDF2.PdfFileReader(f)
                writer.appendPagesFromReader(reader)

                fs[temp_pdf_path] = f

                if (reader.numPages % 2 == 1) and (num_path != len(word_paths) - 1):
                    writer.addBlankPage()

            with open(pdf_path, 'wb') as f:
                writer.write(f)

            for (path, f) in fs.items():
                f.close()
                os.remove(path)

if __name__ == '__main__':
    os.system('taskkill /f /im WINWORD.exe')

    manager = Manager(config_path='config.cfg')
    manager.merge()