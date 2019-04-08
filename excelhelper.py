# coding=utf-8

from xlrd import open_workbook
from xlutils.copy import copy as copy
from xlwt import easyxf, Style
from datetime import datetime


class ExcelHelper:
    ROBOT_LIBRARY_SCOPE = 'GLOBAL'

    def __init__(self):
        self.fileName = None
        self.reader = None
        self.writer = None

    def open_excel(self, file_name):
        """
        打开excel
        :param file_name: 文件名
        """

        self.fileName = file_name
        # 生成读取对象 并拷贝写入对象
        self.reader = open_workbook(self.fileName)
        self.writer = copy(self.reader)

    def refresh_excel(self):
        """
        刷新对象，把写入对象写入到磁盘中，并保持写入和读取对象一致
        """

        self.save_excel()
        self.open_excel(self.fileName)

    def save_excel(self):
        """
        保存更改，把写入对象写入到磁盘中
        """

        self.writer.save(self.fileName)

    def save_as_excel(self, file_name):
        """
        另存为，把写入对象另存
        :param file_name:另存文件名
        """

        self.writer.save(file_name)

    def get_number_of_sheets(self):
        """
        获取表个数
        :return:表个数
        """

        return self.reader.nsheets

    def get_sheet_names(self):
        """
        获取所有表名
        :return:所有表名
        """

        return self.reader.sheet_names()

    def get_column_count(self, sheet_index=0):
        """
        获取第sheet_index表列数
        :param sheet_index:表索引（默认为0）
        :return:第sheet_index表列数
        """

        return self.reader.sheet_by_index(sheet_index).ncols

    def get_row_count(self, sheet_index=0):
        """
        获取第sheet_index表行数
        :param sheet_index:表索引（默认为0）
        :return:第sheet_index表行数
        """

        return self.reader.sheet_by_index(sheet_index).nrows

    def get_column_values(self, column_index, sheet_index=0):
        """
        获取第sheet_index表第column_index列数据
        :param column_index:列索引
        :param sheet_index:表索引（默认为0）
        :return:第sheet_index表第column_index列数据
        """

        sheet = self.reader.sheet_by_index(sheet_index)
        data = []

        # 便利行，取出数据
        for row_index in range(sheet.nrows):
            data.append(sheet.cell(row_index, int(column_index)).value)

        return data

    def get_row_values(self, row_index, sheet_index=0):
        """
        获取第sheet_index表第row_index行数据
        :param row_index:行索引
        :param sheet_index:表索引（默认为0）
        :return:第sheet_index表第row_index行数据
        """

        sheet = self.reader.sheet_by_index(sheet_index)
        data = []

        # 便利列取出数据
        for column_index in range(sheet.ncols):
            data.append(sheet.cell(int(row_index), column_index).value)

        return data

    def get_values(self, sheet_index=0):
        """
        获取第sheet_index表中数据
        :param sheet_index:表索引（默认为0）
        :return:第sheet_index表中数据
        """
        sheet = self.reader.sheet_by_index(sheet_index)
        data = []

        # 遍历每一个数据，存入数组中
        for row_index in range(sheet.nrows):
            row_date = []
            for column_index in range(sheet.ncols):
                row_date.append(sheet.cell(row_index, column_index).value)
            # 保存为二维数据
            data.append(row_date)

        return data

    def get_workbook_values(self):
        """
        获取所有数据
        :return:所有数据
        """

        data = []
        # 遍历每张表每条数据，存为三维数组
        for sheet_index in range(self.reader.nsheets):
            data.append(self.get_values(sheet_index))

        return data

    def get_value(self, row_index, column_index, sheet_index=0):
        """
        获取具体值
        :param row_index:行索引
        :param column_index:列索引
        :param sheet_index:表索引（默认为0）
        :return:具体值
        """

        return self.reader.sheet_by_index(sheet_index).cell(row_index, column_index)

    def set_string_value(self, value, row_index, column_index, sheet_index=0):
        """
        写入字符串数据
        :param value:值
        :param row_index:行索引
        :param column_index:列索引
        :param sheet_index:表索引（默认为0）
        """

        style = easyxf('')
        self.writer.get_sheet(sheet_index).write(row_index, column_index, str(value), style)

    def set_number_value(self, value, row_index, column_index, sheet_index=0):
        """
        写入值数据
        :param value:值
        :param row_index:行索引
        :param column_index:列索引
        :param sheet_index:表索引（默认为0）
        """

        style = easyxf('')
        self.writer.get_sheet(sheet_index).write(row_index, column_index, float(value), style)

    def set_date_value(self, value, row_index, column_index, sheet_index=0):
        """
        写入日期数据 （格式：2019/01/01）
        :param value:值
        :param row_index:行索引
        :param column_index:列索引
        :param sheet_index:表索引（默认为0）
        """

        value = datetime.strptime(value, '%Y/%m/%d')
        style = easyxf('', num_format_str='yyyy/MM/dd')
        self.writer.get_sheet(sheet_index).write(row_index, column_index, value, style)

    def set_time_value(self, value, row_index, column_index, sheet_index=0):
        """
        写入时间数据 （格式：01:01:01）
        :param value:值
        :param row_index:行索引
        :param column_index:列索引
        :param sheet_index:表索引（默认为0）
        """

        value = datetime.strptime(value, '%H:%M:%S')
        style = easyxf('', num_format_str='HH:mm:ss')
        self.writer.get_sheet(sheet_index).write(row_index, column_index, value, style)

    def set_datetime_value(self, value, row_index, column_index, sheet_index=0):
        """
        写入时间日期数据 （格式：2019/01/01 01:01:01）
        :param value:值
        :param row_index:行索引
        :param column_index:列索引
        :param sheet_index:表索引（默认为0）
        """

        value = datetime.strptime(value, '%Y/%m/%d %H:%M:%S')
        style = easyxf('', num_format_str='yyyy/MM/dd HH:mm:ss')
        self.writer.get_sheet(sheet_index).write(row_index, column_index, value, style)
