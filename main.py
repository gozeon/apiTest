# This is a sample Python script.
# Press Shift+F10 to execute it or replace it with your code.
# Press Double Shift to search everywhere for classes, files, tool windows, actions, and settings.
import configparser
import glob
import json
import logging
import os
import re
from datetime import datetime

import requests
import pandas as pd
import numpy as np
from cerberus import Validator


class Report():
    def __init__(self, source, excel_file='report.xlsx', sheet_name='Sheet1'):
        self.source = source
        self.excel_file = excel_file
        self.sheet_name = sheet_name

    def data(self):
        data = []
        for k in self.source:
            data.extend(k['data'])
        return data

    def pass_data(self):
        def valid_filter(target):
            return target['valid']

        return list(filter(valid_filter, self.data()))

    def write(self):
        df = pd.DataFrame(self.data())

        writer = pd.ExcelWriter(self.excel_file, engine='xlsxwriter')

        # doc: https://jike.in/?qa=960238/python-using-pandas-to-format-excel-cell
        def f(x):
            col = "valid"
            r = 'background-color: red'
            g = 'background-color: green'
            c = np.where(x[col] == True, g, r)
            y = pd.DataFrame('', index=x.index, columns=x.columns)
            y[col] = c
            return y

        df = df.style.apply(f, axis=None)
        df.to_excel(writer, sheet_name=self.sheet_name, index=False, startrow=20)

        # Access the XlsxWriter workbook and worksheet objects from the dataframe.
        workbook = writer.book
        worksheet = writer.sheets[self.sheet_name]

        worksheet.write('A1', '通过')
        worksheet.write('A2', '未通过')
        worksheet.write('B1', len(self.pass_data()))
        worksheet.write('B2', len(self.data()) - len(self.pass_data()))

        # Create a chart object.
        chart = workbook.add_chart({'type': 'pie'})

        # Configure the chart from the dataframe data. Configuring the segment
        # colours is optional. Without the 'points' option you will get Excel's
        # default colours.
        chart.add_series({
            'categories': '=Sheet1!A1:A2',
            'values': '=Sheet1!B1:B2',
            'points': [
                {'fill': {'color': 'green'}},
                {'fill': {'color': 'red'}},
            ],
            'data_labels': {'percentage': True},
        })
        chart.set_title({"name": "测试通过率"})

        worksheet.insert_chart('A4', chart)

        writer.save()


class Config():
    def __init__(self, cwd, config_file):
        self.cwd = cwd
        self.config_file = config_file

    def get_config_full_path(self):
        return os.path.join(self.cwd, self.config_file)

    def get_config_base_folder(self):
        return os.path.dirname(self.get_config_full_path())

    def get_config_base_name(self):
        return os.path.basename(self.get_config_base_folder())

    def get_config_data_path(self):
        file_path = os.path.join(self.get_config_base_folder(), 'data.xlsx')
        if os.path.isfile(file_path) is True:
            return file_path
        else:
            return False

    def data_config(self):
        if self.get_config_data_path():
            data = pd.read_excel(self.get_config_data_path())
            return data.to_dict("records")

        return None

    def http_config(self):
        config = configparser.ConfigParser()
        config.read(self.get_config_full_path())
        return config


# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    logging.basicConfig(format=u'%(asctime)s: %(levelname)s: %(message)s', level=logging.INFO)

    cwd = os.getcwd()
    logging.info(cwd)

    config_files = glob.glob("**/*/http.ini")
    logging.info(config_files)

    report = []
    if len(config_files) > 0:
        for config_file in config_files:
            report_item = {}
            config = Config(cwd, config_file)

            report_item['name'] = config.get_config_base_name()

            logging.info("working in " + config.get_config_base_folder())

            report_item_data = []
            if config.data_config() is not None:
                for row in config.data_config():
                    errMsg = None
                    headers = {
                        "Content-Type": "application/json",
                        "Accept": "application/json"
                    }

                    param = {}
                    param_r = re.compile("param.*")
                    for param_key in list(filter(param_r.match, row.keys())):
                        param[param_key.replace("param:", '')] = row[param_key]

                    query = {}
                    query_r = re.compile("query.*")
                    for query_key in list(filter(query_r.match, row.keys())):
                        query[query_key.replace("query:", '')] = row[query_key]

                    body = {}
                    body_r = re.compile("body.*")
                    for body_key in list(filter(body_r.match, row.keys())):
                        body[body_key.replace("body:", '')] = row[body_key]

                    if "schema" in row:
                        try:
                            response = requests.request(method=config.http_config()['Base']['Method'],
                                                        params=query,
                                                        url=str(config.http_config()['Base']['Url']).format(**param),
                                                        data=json.dumps(body),
                                                        headers=headers)
                            v = Validator(json.loads(row['schema']), purge_unknown=True)
                            response_valid = v.validate(json.loads(response.text))
                            if v.errors:
                                errMsg = v.errors


                        except Exception as err:
                            errMsg = repr(err)
                        finally:
                            method = config.http_config()['Base']['Method']
                            valid = response_valid if errMsg is None else False
                            logging.info("[{code}] [{valid}] {method} {url} {errors}".format(code=response.status_code,
                                                                                             method=method,
                                                                                             url=response.url,
                                                                                             valid=valid,
                                                                                             errors=errMsg))
                            report_item_data.append({
                                "name": config.get_config_base_name(),
                                "code": response.status_code,
                                "method": method,
                                "url": response.url,
                                "errors": errMsg,
                                "param": param,
                                "query": query,
                                "body": body,
                                "valid": valid,
                            })
                    else:
                        logging.info('no find "schema" in data.xlsx!')
            else:
                logging.info('no find data.xlsx!')

            report_item['data'] = report_item_data
            report.append(report_item)
    else:
        logging.info('no find config!')

    logging.info('report: ' + str(report))

    report_file = "{}.xlsx".format(datetime.today().strftime("%Y-%m-%d-%H-%M-%S"))
    rep = Report(source=report, excel_file=os.path.join(cwd, report_file))
    rep.write()
