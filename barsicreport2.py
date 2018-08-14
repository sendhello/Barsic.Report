# -*- coding: utf-8 -*-
#
# This file created with KivyCreatorProject
# <https://github.com/HeaTTheatR/KivyCreatorProgect
#
# Copyright © 2017 Easy
#
# For suggestions and questions:
# <kivydevelopment@gmail.com>
# 
# LICENSE: MIT

import os
import sys
from ast import literal_eval
import logging
from datetime import datetime, timedelta
import pyodbc
from decimal import Decimal

from kivy.app import App
from kivy.uix.modalview import ModalView
from kivy.lang import Builder
from kivy.core.window import Window
from kivy.config import ConfigParser
from kivy.clock import Clock
from kivy.utils import get_color_from_hex, get_hex_from_color
from kivy.metrics import dp
from kivy.properties import ObjectProperty, StringProperty
from kivymd.dialog import MDDialog
from kivymd.bottomsheet import MDListBottomSheet, MDGridBottomSheet

from main import __version__
from libs.translation import Translation
from libs.uix.baseclass.startscreen import StartScreen
from libs.uix.lists import Lists
from libs.utils.showplugins import ShowPlugins

from kivymd.theming import ThemeManager
from kivymd.label import MDLabel
from kivymd.time_picker import MDTimePicker
from kivymd.date_picker import MDDatePicker

from toast import toast
from dialogs import card


class BarsicReport2(App):
    """
    Функционал программы.
    """

    title = 'Барсик.Отчеты'
    icon = 'icon.png'
    nav_drawer = ObjectProperty()
    theme_cls = ThemeManager()
    theme_cls.primary_palette = 'Purple'
    theme_cls.theme_style = 'Light'
    lang = StringProperty('ru')

    previous_date_from = ObjectProperty()
    previous_date_to = ObjectProperty()

    def __init__(self, **kvargs):
        super(BarsicReport2, self).__init__(**kvargs)
        Window.bind(on_keyboard=self.events_program)
        Window.soft_input_mode = 'below_target'

        self.list_previous_screens = ['base']
        self.window = Window
        self.plugin = ShowPlugins(self)
        self.config = ConfigParser()
        self.manager = None
        self.window_language = None
        self.exit_interval = False
        self.dict_language = literal_eval(
            open(
                os.path.join(self.directory, 'data', 'locales', 'locales.txt')).read()
        )
        self.translation = Translation(
            self.lang, 'Ttest', os.path.join(self.directory, 'data', 'locales')
        )
        self.date_from = datetime.strptime(datetime.now().strftime('%Y-%m-%d'), '%Y-%m-%d')
        self.date_to = self.date_from + timedelta(1)

        self.org1 = ''
        self.org2 = ''

        self.count_sql_error = 0

    def get_application_config(self):
        return super(BarsicReport2, self).get_application_config(
            '{}/%(appname)s.ini'.format(self.directory))

    def build_config(self, config):
        '''Создаёт файл настроек приложения barsicreport2.ini.'''

        config.adddefaultsection('General')
        config.setdefault('General', 'language', 'ru')
        config.adddefaultsection('MSSQL')
        config.setdefault('MSSQL', 'driver', '{SQL Server}')
        config.setdefault('MSSQL', 'server', '127.0.0.1\\SQLEXPRESS')
        config.setdefault('MSSQL', 'user', 'sa')
        config.setdefault('MSSQL', 'pwd', 'password')
        config.setdefault('MSSQL', 'database1', 'database')
        config.setdefault('MSSQL', 'database2', 'database')
        config.setdefault('MSSQL', 'database_bitrix', 'database')

    def set_value_from_config(self):
        '''Устанавливает значения переменных из файла настроек barsicreport2.ini.'''

        self.config.read(os.path.join(self.directory, 'barsicreport2.ini'))
        self.lang = self.config.get('General', 'language')
        self.driver = self.config.get('MSSQL', 'driver')
        self.server = self.config.get('MSSQL', 'server')
        self.user = self.config.get('MSSQL', 'user')
        self.pwd = self.config.get('MSSQL', 'pwd')
        self.database1 = self.config.get('MSSQL', 'database1')
        self.database2 = self.config.get('MSSQL', 'database2')
        self.database_bitrix = self.config.get('MSSQL', 'database_bitrix')

    def build(self):
        self.set_value_from_config()
        self.load_all_kv_files(os.path.join(self.directory, 'libs', 'uix', 'kv'))
        self.screen = StartScreen()  # главный экран программы
        self.manager = self.screen.ids.manager
        self.nav_drawer = self.screen.ids.nav_drawer
        return self.screen

    def load_all_kv_files(self, directory_kv_files):
        for kv_file in os.listdir(directory_kv_files):
            kv_file = os.path.join(directory_kv_files, kv_file)
            if os.path.isfile(kv_file):
                with open(kv_file, encoding='utf-8') as kv:
                    Builder.load_string(kv.read())

    def events_program(self, instance, keyboard, keycode, text, modifiers):
        '''Вызывается при нажатии кнопки Меню или Back Key
        на мобильном устройстве.'''

        if keyboard in (1001, 27):
            if self.nav_drawer.state == 'open':
                self.nav_drawer.toggle_nav_drawer()
            self.back_screen(event=keyboard)
        elif keyboard in (282, 319):
            pass

        return True

    def back_screen(self, event=None):
        '''Менеджер экранов. Вызывается при нажатии Back Key
        и шеврона "Назад" в ToolBar.'''

        # Нажата BackKey.
        if event in (1001, 27):
            if self.manager.current == 'base':
                self.dialog_exit()
                return
            try:
                self.manager.current = self.list_previous_screens.pop()
            except:
                self.manager.current = 'base'
            self.screen.ids.action_bar.title = self.title
            self.screen.ids.action_bar.left_action_items = \
                [['menu', lambda x: self.nav_drawer._toggle()]]

    def show_plugins(self, *args):
        '''Выводит на экран список плагинов.'''

        self.plugin.show_plugins()

    def show_about(self, *args):
        self.nav_drawer.toggle_nav_drawer()
        self.screen.ids.about.ids.label.text = \
            self.translation._(
                u'[size=20][b]BarsicReport2[/b][/size]\n\n'
                u'[b]Version:[/b] {version}\n'
                u'[b]License:[/b] MIT\n\n'
                u'[size=20][b]Developer[/b][/size]\n\n'
                u'[ref=SITE_PROJECT]'
                u'[color={link_color}]NAME_AUTHOR[/color][/ref]\n\n'
                u'[b]Source code:[/b] '
                u'[ref=REPO_PROJECT]'
                u'[color={link_color}]GitHub[/color][/ref]').format(
                version=__version__,
                link_color=get_hex_from_color(self.theme_cls.primary_color)
            )
        self.manager.current = 'about'
        self.screen.ids.action_bar.left_action_items = \
            [['chevron-left', lambda x: self.back_screen(27)]]

    def show_reports(self, *args):
        self.nav_drawer.toggle_nav_drawer()
        self.manager.current = 'report'
        self.screen.ids.action_bar.left_action_items = \
            [['chevron-left', lambda x: self.back_screen(27)]]

    def show_license(self, *args):
        self.screen.ids.license.ids.text_license.text = \
            self.translation._('%s') % open(
                os.path.join(self.directory, 'LICENSE'), encoding='utf-8').read()
        self.nav_drawer._toggle()
        self.manager.current = 'license'
        self.screen.ids.action_bar.left_action_items = \
            [['chevron-left', lambda x: self.back_screen(27)]]
        self.screen.ids.action_bar.title = \
            self.translation._('MIT LICENSE')

    def select_locale(self, *args):
        '''Выводит окно со списком имеющихся языковых локализаций для
        установки языка приложения.'''

        def select_locale(name_locale):
            '''Устанавливает выбранную локализацию.'''

            for locale in self.dict_language.keys():
                if name_locale == self.dict_language[locale]:
                    self.lang = locale
                    self.config.set('General', 'language', self.lang)
                    self.config.write()

        dict_info_locales = {}
        for locale in self.dict_language.keys():
            dict_info_locales[self.dict_language[locale]] = \
                ['locale', locale == self.lang]

        if not self.window_language:
            self.window_language = card(
                Lists(
                    dict_items=dict_info_locales,
                    events_callback=select_locale, flag='one_select_check'
                ),
                size=(.85, .55)
            )
        self.window_language.open()

    def dialog_exit(self):
        def check_interval_press(interval):
            self.exit_interval += interval
            if self.exit_interval > 5:
                self.exit_interval = False
                Clock.unschedule(check_interval_press)

        if self.exit_interval:
            sys.exit(0)

        Clock.schedule_interval(check_interval_press, 1)
        toast(self.translation._('Press Back to Exit'))

    def show_dialog(self, title, text):
        content = MDLabel(font_style='Body1',
                          theme_text_color='Secondary',
                          text=text,
                          size_hint_y=None,
                          valign='top')
        content.bind(texture_size=content.setter('size'))
        dialog = MDDialog(title=title,
                               content=content,
                               size_hint=(.8, None),
                               height=dp(200),
                               auto_dismiss=False)

        dialog.add_action_button("Закрыть", action=lambda *x: dialog.dismiss())
        dialog.open()

    def show_dialog_sqlerror(self, title, text):
        if self.count_sql_error > 1:
            self.show_dialog(title, text)
            self.count_sql_error = 0
        else:
            self.count_sql_error += 1

    def on_lang(self, instance, lang):
        self.translation.switch_lang(lang)

    def get_time_picker_data(self, instance, time):
        self.root.ids.time_picker_label.text = str(time)
        self.previous_time = time

    def show_time_picker(self):
        self.time_dialog = MDTimePicker()
        self.time_dialog.bind(time=self.get_time_picker_data)
        if self.root.ids.time_picker_use_previous_time.active:
            try:
                self.time_dialog.set_time(self.previous_time)
            except AttributeError:
                pass
        self.time_dialog.open()

    def set_date_from(self, date_obj):
        self.previous_date_from = date_obj
        self.date_from = datetime.strptime(str(date_obj), '%Y-%m-%d')
        self.root.ids.report.ids.date_from.text = str(date_obj)
        if self.date_to <= self.date_from:
            self.root.ids.report.ids.date_to.text = self.show_next_day()
        logging.info(f'{str(datetime.now()):25}:    Установка периода отчета на {self.date_from} - {self.date_to}')

    def show_date_from(self):
        pd = self.previous_date_from
        try:
            MDDatePicker(self.set_date_from,
                         pd.year, pd.month, pd.day).open()
        except AttributeError:
            MDDatePicker(self.set_date_from).open()

    def set_date_to(self, date_obj):
        self.previous_date_to = date_obj
        self.date_to = datetime.strptime(str(date_obj), '%Y-%m-%d')
        self.root.ids.report.ids.date_to.text = str(date_obj)
        if self.date_to <= self.date_from:
            self.root.ids.report.ids.date_from.text = self.show_pre_day()
        logging.info(f'{str(datetime.now()):25}:    Установка периода отчета на {self.date_from} - {self.date_to}')

    def show_date_to(self):
        pd = self.previous_date_to
        try:
            MDDatePicker(self.set_date_to,
                         pd.year, pd.month, pd.day).open()
        except AttributeError:
            MDDatePicker(self.set_date_to).open()

    def click_date_switch(self):
        if self.root.ids.report.ids.date_switch.active:
            self.root.ids.report.ids.label_date.text = 'Дата:'
            self.root.ids.report.ids.date_to.text = self.show_next_day()
        else:
            self.root.ids.report.ids.label_date.text = 'Период:'

    def show_today(self):
        return self.date_from.strftime("%Y-%m-%d")

    def show_next_day(self):
        try:
            self.date_to = datetime.strptime(self.root.ids.report.ids.date_from.text, "%Y-%m-%d") + timedelta(1)
        except AttributeError:
            self.date_to = datetime.strptime(datetime.now().strftime('%Y-%m-%d'), '%Y-%m-%d') + timedelta(1)
        logging.info(f'{str(datetime.now()):25}:    Установка периода отчета на {self.date_from} - {self.date_to}')
        return self.date_to.strftime("%Y-%m-%d")

    def show_pre_day(self):
        try:
            self.date_from = datetime.strptime(self.root.ids.report.ids.date_to.text, "%Y-%m-%d") - timedelta(1)
        except AttributeError:
            self.date_from = datetime.strptime(datetime.now().strftime('%Y-%m-%d'), '%Y-%m-%d')
        return self.date_from.strftime("%Y-%m-%d")

    def count_clients(
            self,
            driver,
            server,
            database,
            uid,
            pwd,
    ):
        """
        Количество человек в зоне
        :return: Количество человек в зоне
        """

        logging.info(f'{str(datetime.now()):25}:    Выполнение функции "count_clients"')

        result = []

        try:
            logging.info(f'{str(datetime.now()):25}:    Попытка соединения с {server}')

            cnxn = pyodbc.connect(
                f'DRIVER={driver};SERVER={server};DATABASE={database};UID={uid};PWD={pwd}')
            cursor = cnxn.cursor()

            cursor.execute("""
                            SELECT
                                [gr].[c1] as [c11],
                                [gr].[StockCategory_Id] as [StockCategory_Id1],
                                [c].[Name],
                                [c].[NN]
                            FROM
                                (
                                    SELECT
                                        [_].[CategoryId] as [StockCategory_Id],
                                        Count(*) as [c1]
                                    FROM
                                        [AccountStock] [_]
                                            INNER JOIN [SuperAccount] [t1] ON [_].[SuperAccountId] = [t1].[SuperAccountId]
                                    WHERE
                                        [_].[StockType] = 41 AND
                                        [t1].[Type] = 0 AND
                                        [_].[Amount] > 0 AND
                                        NOT ([t1].[IsStuff] = 1)
                                    GROUP BY
                                        [_].[CategoryId]
                                ) [gr]
                                    INNER JOIN [Category] [c] ON [gr].[StockCategory_Id] = [c].[CategoryId]
                           """)
            while True:
                row = cursor.fetchone()
                if row:
                    result.append(row)
                else:
                    break
            logging.info(f'{str(datetime.now()):25}:    Результат функции "count_clients": {result}')
            if not result:
                result.append(('Пусто', 488, '', '0003'))

        except pyodbc.OperationalError as e:
            logging.error(f'{str(datetime.now()):25}:    Ошибка {repr(e)}')
            result.append(('Нет данных', 488, 'Ошибка соединения', repr(e)))
            self.show_dialog(f'Ошибка соединения с {server}: {database}', repr(e))
        except pyodbc.ProgrammingError as e:
            logging.error(f'{str(datetime.now()):25}:    Ошибка {repr(e)}')
            result.append(('Нет данных', 488, 'Ошибка соединения', repr(e)))
            self.show_dialog(f'Невозможно открыть {database}', repr(e))
        return result

    def count_clients_print(self):
        count_clients = self.count_clients(
            driver=self.driver,
            server=self.server,
            database=self.database1,
            uid=self.user,
            pwd=self.pwd,
        )
        self.screen.ids.base.ids.name_zone.text = str(count_clients[len(count_clients) - 1][2])
        self.screen.ids.base.ids.count.text = str(count_clients[len(count_clients) - 1][0])

    def select_org1(self):
        org_list = self.list_organisation(
            server=self.server,
            database=self.database1,
            uid=self.user,
            pwd=self.pwd,
            driver=self.driver,
        )
        if org_list:
            bs = MDListBottomSheet()
            for org in org_list:
                bs.add_item(org[2], lambda x: self.click_select_org(org[0], org[2], self.database1), icon='nfc')
            bs.open()

    def select_org2(self):
        org_list = self.list_organisation(
            server=self.server,
            database=self.database2,
            uid=self.user,
            pwd=self.pwd,
            driver=self.driver,
        )
        if org_list:
            bs = MDListBottomSheet()
            for org in org_list:
                bs.add_item(org[2], lambda x: self.click_select_org(org[0], org[2], self.database2), icon='nfc')
            bs.open()

    def click_select_org(self, id, name, database):
        if database == self.database1:
            self.org1 = (id, name)
            self.screen.ids.report.ids.org1.text = name
        elif database == self.database2:
            self.org2 = (id, name)
            self.screen.ids.report.ids.org2.text = name


    def list_organisation(self,
                          server,
                          database,
                          uid,
                          pwd,
                          driver,
                          ):
        """
        Функция делает запрос в базу Барс и возвращает список заведенных в базе организаций в виде списка кортежей
        :param server: str - Путь до MS-SQL сервера, например 'SQLEXPRESS\BASE'
        :param database: str - Имя базы данных Барса, например 'SkiBars2'
        :param uid: str - Имя пользователя базы данных, например 'sa'
        :param pwd: str - Пароль пользователя базы данных, например 'pass'
        :return: list = Список организаций, каджая из которых - кортеж с параметрами организации
        """
        result = []
        try:
            logging.info(f'{str(datetime.now()):25}:    Попытка соединения с {server}')
            cnxn = pyodbc.connect(f'DRIVER={driver};SERVER={server};DATABASE={database};UID={uid};PWD={pwd}')
            cursor = cnxn.cursor()

            id_type = 1
            cursor.execute(
                f"""
                SELECT
                    SuperAccountId, Type, Descr, CanRegister, CanPass, IsStuff, IsBlocked, BlockReason, DenyReturn, 
                    ClientCategoryId, DiscountCard, PersonalInfoId, Address, Inn, ExternalId, RegisterTime,LastTransactionTime, 
                    LegalEntityRelationTypeId, SellServicePointId, DepositServicePointId, AllowIgnoreStoredPledge, Email, 
                    Latitude, Longitude, Phone, WebSite, TNG_ProfileId
                FROM
                    SuperAccount
                WHERE
                    Type={id_type}
                """)
            while True:
                row = cursor.fetchone()
                if row:
                    result.append(row)
                else:
                    break
        except pyodbc.OperationalError as e:
            logging.error(f'{str(datetime.now()):25}:    Ошибка {repr(e)}')
            self.show_dialog(f'Ошибка соединения с {server}: {database}', repr(e))
        except pyodbc.ProgrammingError as e:
            logging.error(f'{str(datetime.now()):25}:    Ошибка {repr(e)}')
            self.show_dialog(f'Невозможно открыть {database}', repr(e))
        return result

    def itog_report(
            self,
            server,
            database,
            driver,
            user,
            pwd,
            org,
            date_from,
            date_to,
            hide_zeroes='0',
            hide_internal='1',
    ):
        """
        Делает запрос в базу Барс и возвращает итоговый отчет за запрашиваемый период
        :param server: str - Путь до MS-SQL сервера, например 'SQLEXPRESS\BASE'
        :param database: str - Имя базы данных Барса, например 'SkiBars2'
        :param uid: str - Имя пользователя базы данных, например 'sa'
        :param pwd: str - Пароль пользователя базы данных, например 'pass'
        :param sa: str - Id организации в Барсе
        :param date_from: str - Начало отчетного периода в формате: 'YYYYMMDD 00:00:00'
        :param date_to:  str - Конец отчетного периода в формате: 'YYYYMMDD 00:00:00'
        :param hide_zeroes: 0 or 1 - Скрывать нулевые позиции?
        :param hide_internal: 0 or 1 - Скрывать внутренние точки обслуживания?
        :return: Итоговый отчет
        """
        cnxn = pyodbc.connect(f'DRIVER={driver};SERVER={server};DATABASE={database};UID={user};PWD={pwd}')
        date_from = date_from.strftime('%Y%m%d 00:00:00')
        date_to = date_to.strftime('%Y%m%d 00:00:00')
        cursor = cnxn.cursor()
        cursor.execute(
            f"exec sp_reportOrganizationTotals_v2 @sa={org},@from='{date_from}',@to='{date_to}',@hideZeroes={hide_zeroes},"
            f"@hideInternal={hide_internal}")
        report = []
        while True:
            row = cursor.fetchone()
            if row:
                report.append(row)
            else:
                break
        if len(report) > 1:
            logging.info(f'{str(datetime.now())[:-7]}: Отчет сформирован ID организации = {org}, '
                         f'Период: {date_from[:8]}-{date_to[:8]}, Скрывать нули = {hide_zeroes}, .'
                         f'Скрывать внутренние точки обслуживания: {hide_internal})')
        return report

    def itog_report_convert_to_dict(self, report):
        """
        Преобразует список кортежей отчета в словарь
        :param report: list - Итоговый отчет в формате списка картежей полученный из функции full_report
        :return: dict - Словарь услуг и их значений
        """
        result = {}
        for row in report:
            result[row[4]] = (row[1], row[0])
        return result

    def add_sum(self, report):
        """
        Расчитывает и добавляет к словарю-отчету 1 элемент: Итого
        :param report: dict - словарь-отчет
        :return: dict - словарь-отчет
        """
        sum_service = Decimal(0)
        sum_many = Decimal(0)
        for line in report:
            if not (report[line][0] is None or report[line][0] is None):
                if line != 'Депозит':
                    sum_service += report[line][0]
                sum_many += report[line][1]
        report['Итого по отчету'] = (sum_service, sum_many)
        return report

    def decimal_to_googlesheet(self, dict):
        """
        Преобразует суммы Decimal в float
        :param dict:
        :return:
        """
        new_dict = {}
        for key in dict:
            if type(dict[key][0]) is Decimal:
                new_dict[key] = (int(dict[key][0]), float(dict[key][1]))
            else:
                new_dict[key] = (dict[key][0], dict[key][1])
        return new_dict

    def read_bitrix_base(self,
                         server,
                         database,
                         user,
                         pwd,
                         driver,
                         date_from,
                         date_to,
                         ):
        """
        Функция делает запрос в базу и возвращает список продаж за аказанную дату
        :param server: str - Путь до MS-SQL сервера, например 'SQLEXPRESS\BASE'
        :param database: str - Имя базы данных Барса, например 'SkiBars2'
        :param uid: str - Имя пользователя базы данных, например 'sa'
        :param pwd: str - Пароль пользователя базы данных, например 'pass'
        :param date_from: str - Начало отчетного периода в формате: 'YYYYMMDD 00:00:00'
        :param date_to:  str - Конец отчетного периода в формате: 'YYYYMMDD 00:00:00'
        :return: list = Список организаций, каджая из которых - кортеж с параметрами организации
        """
        date_from = (date_from - timedelta(1)).strftime("%Y%m%d") + " 19:00:00"
        date_to = (date_to - timedelta(1)).strftime("%Y%m%d") + " 19:00:00"

        cnxn = pyodbc.connect(f'DRIVER={driver};SERVER={server};DATABASE={database};UID={user};PWD={pwd}')
        cursor = cnxn.cursor()

        cursor.execute(
            f"""
            SELECT 
                Id, OrderNumber, ProductId, ProductName, OrderDate, PayDate, Sum, Pay, Status, Client
            FROM 
                Transactions
            WHERE
                (OrderDate >= '{date_from}')and(OrderDate < '{date_to}')
            """)
        orders = []
        while True:
            row = cursor.fetchone()
            if row:
                orders.append(row)
            else:
                break

        summ = 0
        for order in orders:
            summ += float(order[6])
        return len(orders), summ

    def run_report(self):
        """
        Выполнить отчеты
        :return:
        """
        itog_report_org1 = None
        itog_report_org2 = None
        report_bitrix = None

        if self.org1:
            itog_report_org1 = self.itog_report(
                server=self.server,
                database=self.database1,
                driver=self.driver,
                user=self.user,
                pwd=self.pwd,
                org=self.org1[0],
                date_from=self.date_from,
                date_to=self.date_to,
                hide_zeroes='0',
                hide_internal='1',
            )
        if self.org2:
            itog_report_org2 = self.itog_report(
                server=self.server,
                database=self.database2,
                driver=self.driver,
                user=self.user,
                pwd=self.pwd,
                org=self.org2[0],
                date_from=self.date_from,
                date_to=self.date_to,
                hide_zeroes='0',
                hide_internal='1',
            )

        # Преобразование в словарь
        itog_report_org1 = self.itog_report_convert_to_dict(itog_report_org1)
        itog_report_org2 = self.itog_report_convert_to_dict(itog_report_org2)

        # Добавление ИТОГО
        itog_report_org1 = self.add_sum(itog_report_org1)
        itog_report_org2 = self.add_sum(itog_report_org2)

        # Decimal to float
        itog_report_org1 = self.decimal_to_googlesheet(itog_report_org1)
        itog_report_org2 = self.decimal_to_googlesheet(itog_report_org2)

        report_bitrix = self.read_bitrix_base(
            server=self.server,
            database=self.database_bitrix,
            user=self.user,
            pwd=self.pwd,
            driver=self.driver,
            date_from=self.date_from,
            date_to=self.date_to,
        )



        print(itog_report_org1)
        print(itog_report_org2)
        print(report_bitrix)


if __name__ == '__main__':
    pass
