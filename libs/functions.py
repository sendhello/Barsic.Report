# /usr/bin/python3
# -*- coding: utf-8 -*-

from decimal import Decimal


def convert_to_dict(func):
    def itog_report_convert_to_dict(*args, **kwargs):
        """
        Преобразует список кортежей отчета в словарь
        :param report: list - Итоговый отчет в формате списка картежей полученный из функции full_report
        :return: dict - Словарь услуг и их значений
        """
        report = func(*args, **kwargs)
        result = {}
        for row in report:
            result[row[4]] = (row[1], row[0])
        return result
    return itog_report_convert_to_dict


def add_sum(func):
    def add_sum_wrapper(*args, **kwargs):
        """
        Расчитывает и добавляет к словарю-отчету 1 элемент: Итого
        :param report: dict - словарь-отчет
        :return: dict - словарь-отчет
        """
        report = func(*args, **kwargs)
        sum_service = Decimal(0)
        sum_many = Decimal(0)
        for line in report:
            if not (report[line][0] is None or report[line][0] is None):
                if line != 'Депозит':
                    sum_service += report[line][0]
                sum_many += report[line][1]
        report['Итого по отчету'] = (sum_service, sum_many)
        return report
    return add_sum_wrapper


def to_googleshet(func):
    def decimal_to_googlesheet(*args, **kwargs):
        """
        Преобразует суммы Decimal в float
        """
        dict = func(*args, **kwargs)
        new_dict = {}
        for key in dict:
            if type(dict[key][0]) is Decimal:
                new_dict[key] = (int(dict[key][0]), float(dict[key][1]))
            else:
                new_dict[key] = (dict[key][0], dict[key][1])
        return new_dict
    return decimal_to_googlesheet