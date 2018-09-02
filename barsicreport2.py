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
from datetime import datetime, timedelta, date
import pyodbc
from decimal import Decimal
from lxml import etree, objectify
import csv

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
from kivymd.textfields import MDTextField

from main import __version__
from libs.translation import Translation
from libs.uix.baseclass.startscreen import StartScreen
from libs.uix.lists import Lists
from libs.utils.showplugins import ShowPlugins

from libs import functions, to_google_sheets

from kivymd.theming import ThemeManager
from kivymd.label import MDLabel
from kivymd.time_picker import MDTimePicker
from kivymd.date_picker import MDDatePicker

from toast import toast
from dialogs import card
import yadisk
import xlwt
import itertools

import webbrowser
import httplib2
import apiclient.discovery
from oauth2client.service_account import ServiceAccountCredentials
import telepot


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
        self.org_for_finreport = {}
        self.new_service = []
        self.orgs = []
        self.new_agentservice = []
        self.agentorgs = []

    def get_application_config(self):
        return super(BarsicReport2, self).get_application_config(
            '{}/%(appname)s.ini'.format(self.directory))

    def build_config(self, config):
        """Создаёт файл настроек приложения barsicreport2.ini."""
        config.adddefaultsection('General')
        config.setdefault('General', 'language', 'ru')
        config.setdefault('General', 'finreport_xls', 'False')
        config.setdefault('General', 'finreport_google', 'False')
        config.setdefault('General', 'finreport_telegram', 'False')
        config.setdefault('General', 'agentreport_xls', 'False')
        config.setdefault('General', 'split_by_days', 'False')
        config.setdefault('General', 'date_switch', 'True')
        config.setdefault('General', 'use_yadisk', 'False')
        config.adddefaultsection('MSSQL')
        config.setdefault('MSSQL', 'driver', '{SQL Server}')
        config.setdefault('MSSQL', 'server', '127.0.0.1\\SQLEXPRESS')
        config.setdefault('MSSQL', 'user', 'sa')
        config.setdefault('MSSQL', 'pwd', 'password')
        config.setdefault('MSSQL', 'database1', 'database')
        config.setdefault('MSSQL', 'database2', 'database')
        config.setdefault('MSSQL', 'database_bitrix', 'database')
        config.adddefaultsection('PATH')
        config.setdefault('PATH', 'reportXML', 'data/org_for_report.xml')
        config.setdefault('PATH', 'agentXML', 'data/org_plat_agent.xml')
        config.setdefault('PATH', 'local_folder', 'report')
        config.setdefault('PATH', 'path_aquapark', 'report')
        config.setdefault('PATH', 'path_beach', 'report')
        config.setdefault('PATH', 'CREDENTIALS_FILE', 'data/1720aecc5640.json')
        config.setdefault('PATH', 'list_google_docs', 'data/list_google_docs.csv')
        config.adddefaultsection('Yadisk')
        config.setdefault('Yadisk', 'yadisk_token', 'token')
        config.adddefaultsection('Telegram')
        config.setdefault('Telegram', 'telegram_token', '111111111:aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa')
        config.setdefault('Telegram', 'telegram_chanel_id', '111111111111')
        config.setdefault('Telegram', 'telegram_proxy', 'http://10.10.10.10:80')
        config.setdefault('Telegram', 'telegram_basic_auth', '("login", "password")')
        config.adddefaultsection('GoogleShets')
        config.setdefault('GoogleShets', 'google_all_read', 'False')
        config.setdefault('GoogleShets', 'google_reader_list', '')
        config.setdefault('GoogleShets', 'google_writer_list', '')

    def set_value_from_config(self):
        '''Устанавливает значения переменных из файла настроек barsicreport2.ini.'''
        self.config.read(os.path.join(self.directory, 'barsicreport2.ini'))
        self.lang = self.config.get('General', 'language')
        self.finreport_xls = functions.to_bool(self.config.get('General', 'finreport_xls'))
        self.finreport_google = functions.to_bool(self.config.get('General', 'finreport_google'))
        self.finreport_telegram = functions.to_bool(self.config.get('General', 'finreport_telegram'))
        self.agentreport_xls = functions.to_bool(self.config.get('General', 'agentreport_xls'))
        # self.split_by_days = functions.to_bool(self.config.get('General', 'split_by_days'))
        self.split_by_days = False
        self.date_switch = functions.to_bool(self.config.get('General', 'date_switch'))
        self.use_yadisk = functions.to_bool(self.config.get('General', 'use_yadisk'))
        self.driver = self.config.get('MSSQL', 'driver')
        self.server = self.config.get('MSSQL', 'server')
        self.user = self.config.get('MSSQL', 'user')
        self.pwd = self.config.get('MSSQL', 'pwd')
        self.database1 = self.config.get('MSSQL', 'database1')
        self.database2 = self.config.get('MSSQL', 'database2')
        self.database_bitrix = self.config.get('MSSQL', 'database_bitrix')
        self.reportXML = self.config.get('PATH', 'reportXML')
        self.agentXML = self.config.get('PATH', 'agentXML')
        self.local_folder = self.config.get('PATH', 'local_folder')
        self.path_aquapark = self.config.get('PATH', 'path_aquapark')
        self.path_beach = self.config.get('PATH', 'path_beach')
        self.CREDENTIALS_FILE = self.config.get('PATH', 'CREDENTIALS_FILE')
        self.list_google_docs = self.config.get('PATH', 'list_google_docs')
        self.yadisk_token = self.config.get('Yadisk', 'yadisk_token')
        self.telegram_token = self.config.get('Telegram', 'telegram_token')
        self.telegram_chanel_id = self.config.get('Telegram', 'telegram_chanel_id') # '215624388'
        self.telegram_proxy = self.config.get('Telegram', 'telegram_proxy')
        self.telegram_basic_auth = self.config.get('Telegram', 'telegram_basic_auth')
        self.google_all_read = functions.to_bool(self.config.get('GoogleShets', 'google_all_read'))
        self.google_reader_list = self.config.get('GoogleShets', 'google_reader_list')
        self.google_writer_list = self.config.get('GoogleShets', 'google_writer_list')

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
        """
        Переход на экран ОТЧЕТЫ
        :param args:
        :return:
        """
        self.nav_drawer.toggle_nav_drawer()
        self.manager.current = 'report'
        self.screen.ids.action_bar.left_action_items = \
            [['chevron-left', lambda x: self.back_screen(27)]]
        # Загрузка параметров из INI-файла
        self.load_checkbox()
        self.set_date_from(datetime.now().date())

    def show_license(self, *args):
        """
        Переход на экран ЛИЦЕНЗИЯ
        :param args:
        :return:
        """
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
        """
        Выводит окно со списком имеющихся языковых локализаций для
        установки языка приложения.
        :param args:
        :return:
        """

        def select_locale(name_locale):
            """
            Устанавливает выбранную локализацию.
            :param name_locale:
            :return:
            """

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

    def show_dialog(self, title, text, func=functions.func_pass, *args, **kwargs):
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

        dialog.add_action_button("Закрыть", action=lambda *x: (dialog.dismiss(), func(*args, **kwargs)))
        dialog.open()

    def show_dialog_variant(self, title, text, func=functions.func_pass, *args, **kwargs):
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

        dialog.add_action_button("ДА", action=lambda *x: (dialog.dismiss(), func(*args, **kwargs)))
        dialog.add_action_button("Нет", action=lambda *x: (dialog.dismiss(), False))
        dialog.open()

    def show_dialog_variant2(self, title, text, func_yes=functions.func_pass, func_no=functions.func_pass):
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

        dialog.add_action_button("ДА", action=lambda *x: (dialog.dismiss(), func_yes))
        dialog.add_action_button("Нет", action=lambda *x: (dialog.dismiss(), func_no))
        dialog.open()

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
        if self.date_to <= self.date_from or self.root.ids.report.ids.date_switch.active:
            self.set_date_to(date_obj + timedelta(1))
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
            self.set_date_from(date_obj - timedelta(1))
        logging.info(f'{str(datetime.now()):25}:    Установка периода отчета на {self.date_from} - {self.date_to}')

    def show_date_to(self):
        if self.root.ids.report.ids.date_switch.active:
            pass
        else:
            pd = self.previous_date_to
            try:
                MDDatePicker(self.set_date_to,
                             pd.year, pd.month, pd.day).open()
            except AttributeError:
                MDDatePicker(self.set_date_to).open()

    def click_date_switch(self):
        if self.root.ids.report.ids.date_switch.active:
            self.date_switch = True
            self.root.ids.report.ids.label_date.text = 'Дата:'
            self.set_date_to(self.date_from.date() + timedelta(1))
            self.root.ids.report.ids.date_to.theme_text_color = 'Secondary'
            self.root.ids.report.ids.split_by_days.active = False
            self.root.ids.report.ids.split_by_days.disabled = True
            self.root.ids.report.ids.split_by_days_text.theme_text_color = 'Secondary'
            self.change_checkbox('split_by_days', False)
            self.root.ids.report.ids.finreport_google_text.disabled = False
            self.root.ids.report.ids.finreport_google.disabled = False
        else:
            self.date_switch = False
            self.root.ids.report.ids.label_date.text = 'Период:'
            self.root.ids.report.ids.date_to.theme_text_color = 'Primary'
            self.root.ids.report.ids.split_by_days.disabled = False
            self.root.ids.report.ids.split_by_days.active = True
            self.root.ids.report.ids.split_by_days_text.theme_text_color = 'Primary'
            self.change_checkbox('split_by_days', True)

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
        in_zone = self.count_clients(
            driver=self.driver,
            server=self.server,
            database=self.database1,
            uid=self.user,
            pwd=self.pwd,
        )
        self.click_select_org()
        try:
            count_clients = int(self.itog_report(
                    server=self.server,
                    database=self.database1,
                    driver=self.driver,
                    user=self.user,
                    pwd=self.pwd,
                    org=self.org1[0],
                    date_from=datetime.now(),
                    date_to=datetime.now()+timedelta(1),
                    hide_zeroes='0',
                    hide_internal='1',
                )['Аквазона'][0])
        except KeyError:
            count_clients = 0

        self.screen.ids.base.ids.count_clients.text = str(count_clients)
        self.screen.ids.base.ids.name_zone.text = str(in_zone[len(in_zone) - 1][2])
        self.screen.ids.base.ids.count.text = str(in_zone[len(in_zone) - 1][0])


    # -------------------------------Кнопки вывода списка организаций для выбора----------------------------------------

    # def select_org1(self):
    #     """
    #     Вывод списка организаций
    #     :return:
    #     """
    #     org_list = self.list_organisation(
    #         server=self.server,
    #         database=self.database1,
    #         uid=self.user,
    #         pwd=self.pwd,
    #         driver=self.driver,
    #     )
    #     if org_list:
    #         bs = MDListBottomSheet()
    #         for org in org_list:
    #             bs.add_item(org[2], lambda x: self.click_select_org(org[0], org[2], self.database1), icon='nfc')
    #         bs.open()
    #
    # def select_org2(self):
    #     """
    #     Вывод списка организаций
    #     :return:
    #     """
    #     org_list = self.list_organisation(
    #         server=self.server,
    #         database=self.database2,
    #         uid=self.user,
    #         pwd=self.pwd,
    #         driver=self.driver,
    #     )
    #     if org_list:
    #         bs = MDListBottomSheet()
    #         for org in org_list:
    #             bs.add_item(org[2], lambda x: self.click_select_org(org[0], org[2], self.database2), icon='nfc')
    #         bs.open()
    #
    # def click_select_org(self, id, name, database):
    #     """
    #     Выбор организации из списка и запись ее в переменную
    #     :param id:
    #     :param name:
    #     :param database:
    #     :return:
    #     """
    #     if database == self.database1:
    #         self.org1 = (id, name)
    #         self.screen.ids.report.ids.org1.text = name
    #     elif database == self.database2:
    #         self.org2 = (id, name)
    #         self.screen.ids.report.ids.org2.text = name

    # ---------- Выбор первой организации из списка организаций (Замена кнопкам выбора организаций) --------------------

    def click_select_org(self):
        """
        Выбор первой организации из списка организаций
        """
        org_list1 = self.list_organisation(
                server=self.server,
                database=self.database1,
                uid=self.user,
                pwd=self.pwd,
                driver=self.driver,
            )
        org_list2 = self.list_organisation(
            server=self.server,
            database=self.database2,
            uid=self.user,
            pwd=self.pwd,
            driver=self.driver,
        )
        self.org1 = (org_list1[0][0], org_list1[0][2])
        self.org2 = (org_list2[0][0], org_list2[0][2])
        logging.info(f'{str(datetime.now()):25}:    Выбраны организации {org_list1[0][2]} и {org_list2[0][2]}')

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
        logging.info(f'{str(datetime.now()):25}:    Поиск организаций...')
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

    @functions.to_googleshet
    @functions.add_date
    @functions.add_sum
    @functions.convert_to_dict
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
        logging.info(f'{str(datetime.now()):25}:    Чтение online-продаж...')
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

    def read_reportgroup(self, XML):
        """
        Чтение XML с привязкой групп услуг к услугам
        :param path:
        :return:
        """
        with open(XML, encoding='utf-8') as f:
            xml = f.read()

        root = objectify.fromstring(xml)
        orgs_dict = {}

        for org in root.UrFace:
            orgs_dict[org.get('Name')] = []
            for serv in org.Services.Service:
                if serv.get('Name') != 'Пустая обязательная категория':
                    orgs_dict[org.get('Name')].append(serv.get('Name'))

        return orgs_dict

    def find_new_service(self, service_dict, orgs_dict):
        """
        Поиск новых услуг и организаций из XML
        :param service_dict: Итоговый отчет
        :param orgs_dict: словарь из XML-файла
        :return:
        """
        servise_set = set()

        for key in orgs_dict:
            for s in orgs_dict[key]:
                servise_set.add(s)

        for org in service_dict:
            if org not in servise_set and org not in self.new_service:
                self.new_service.append(org)
                servise_set.add(org)

        for key in orgs_dict:
            if key not in self.orgs:
                self.orgs.append(key)

    def distibution_service(self):
        """
        Извлекает услугу из списка нвых услуг и вызывает список групп до тех пор пока есть новые услуги,
        затем передает управление следующей функции
        """
        if self.new_service:
            service = self.new_service.pop()
            self.viev_orgs(service)
        else:
            self.agentservice()

    def viev_orgs(self, service):
        """
        Выводит всплывающий список групп, при клике на одну из которых услуга указанная в заголовке добавляется в нее
        """
        bs = MDListBottomSheet()
        bs.add_item(f'К какой группе отчета относится услуга "{service}"? (1 из {len(self.new_service) + 1})',
                    lambda x: x)
        for i in range(len(self.orgs)):
            if self.orgs[i] != 'ИТОГО' and self.orgs[i] != 'Депозит' and self.orgs[i] != 'Дата':
                bs.add_item(self.orgs[i], lambda x: self.select_org(service, x.text), icon='nfc')
        bs.add_item(f'Добавить новую группу отчета...',
                    lambda x: self.show_dialog_add_org("Новая группа", "Название новой группы", service))
        bs.open()

    def show_dialog_add_org(self, title, text, service):
        """
        Выводит диалоговое окно с возможность ввода имени новой группы и двумя кнопками
        """
        content = MDTextField(hint_text="Persistent helper text222",
                              helper_text="Text is always here111",
                              helper_text_mode="persistent",
                              text=text,
                              )
        dialog = MDDialog(title=title,
                          content=content,
                          size_hint=(.8, None),
                          height=dp(200),
                          auto_dismiss=False)
        dialog.add_action_button("Отмена", action=lambda *x: (dialog.dismiss(), self.readd_org(service)))
        dialog.add_action_button("Добавить",
                                 action=lambda *x: (dialog.dismiss(), self.create_new_org(dialog.content.text, service)))
        dialog.open()

    def create_new_org(self, name, service):
        """
        Добавляет новую организацию в список организаций self.orgs, словарь self.orgs_dict и XML конфигурацию.
        Возвращает изьятую ранее услугу в список новых услуг с помощью функции self.readd_org
        """
        logging.info(f'{str(datetime.now()):25}:    Добавление новой группы - {name}')
        logging.info(f'{str(datetime.now()):25}:    Добавление услуги {service} в группу {name}')
        self.orgs.append(name)
        self.orgs_dict[name] = []
        self.readd_org(service)
        with open(self.reportXML, encoding='utf-8') as f:
            xml = f.read()
        root = objectify.fromstring(xml)
        #Добавляем новые организации
        new_org = objectify.SubElement(root, "UrFace")
        new_org.set('Name', name)
        new_servs = objectify.SubElement(new_org, 'Services')
        new_serv = objectify.SubElement(new_servs, 'Service')
        new_serv.set('Name', 'Пустая обязательная категория')
        # удаляем аннотации.
        objectify.deannotate(root)
        etree.cleanup_namespaces(root)
        obj_xml = etree.tostring(root,
                                 encoding='utf-8',
                                 pretty_print=True,
                                 xml_declaration=True
                                 )
        # сохраняем данные в файл.
        try:
            with open(self.reportXML, "w", encoding='utf_8_sig') as xml_writer:
                xml_writer.write(obj_xml.decode('utf-8'))
        except IOError:
            pass


    def readd_org(self, service):
        """
        Возвращает изьятую ранее услугу в список новых услуг, затем вызывает функцию распределения
        """
        self.new_service.append(service)
        self.distibution_service()


    def select_org(self, service, org):
        """
        Добавляет услугу в список услуг и вызывает функцию распределения для других услуг
        """
        logging.info(f'{str(datetime.now()):25}:    Добавление услуги {service} в группу {org}')
        self.orgs_dict[org].append(service)
        #Запись новой услуги в XML
        with open(self.reportXML, encoding='utf-8') as f:
            xml = f.read()
        root = objectify.fromstring(xml)
        # Перечисляем существующие организации в файле и добавляем новые строчки
        for x in root.UrFace:
            if x.get('Name') == org:
                Service = objectify.SubElement(x.Services, "Service")
                Service.set("Name", service)
        # удаляем аннотации.
        objectify.deannotate(root)
        etree.cleanup_namespaces(root)
        obj_xml = etree.tostring(root, encoding='utf-8', pretty_print=True, xml_declaration=True)
        # сохраняем данные в файл.
        try:
            with open(self.reportXML, "w", encoding='utf_8_sig') as xml_writer:
                xml_writer.write(obj_xml.decode('utf-8'))
        except IOError:
            pass
        self.distibution_service()

    def agentservice(self):
        self.agent_dict = self.read_reportgroup(self.agentXML)
        self.find_new_agentservice(self.itog_report_org1, self.agent_dict)
        self.distibution_agentservice()

    def distibution_agentservice(self):
        if self.new_agentservice:
            service = self.new_agentservice.pop()
            self.viev_agentorgs(service)
        else:
            self.save_reports()

    def find_new_agentservice(self, service_dict, orgs_dict):
        """
        Поиск новых услуг и организаций из XML
        :param service_dict: Итоговый отчет
        :param orgs_dict: словарь из XML-файла
        :return:
        """
        servise_set = set()

        for key in orgs_dict:
            for s in orgs_dict[key]:
                servise_set.add(s)

        for org in service_dict:
            if org not in servise_set and org not in self.new_agentservice:
                self.new_agentservice.append(org)
                servise_set.add(org)

        for key in orgs_dict:
            if key not in self.agentorgs:
                self.agentorgs.append(key)

    def viev_agentorgs(self, service):
        """
        Выводит всплывающий список организаций,
        при клике на одну из которых услуга указанная в заголовке добавляется в нее
        """
        bs = MDListBottomSheet()
        bs.add_item(f'К какой организации относится услуга "{service}"? (1 из {len(self.new_agentservice) + 1})',
                    lambda x: x)
        for i in range(len(self.agentorgs)):
            if self.agentorgs[i] != 'ИТОГО' and self.agentorgs[i] != 'Депозит' and self.agentorgs[i] != 'Дата':
                bs.add_item(self.agentorgs[i], lambda x: self.select_agentorg(service, x.text), icon='nfc')
        bs.add_item(f'Добавить новую организацию...',
                    lambda x: self.show_dialog_add_agentorg('Наименование организации', 'ООО Рога и Копыта', service))
        bs.open()

    def select_agentorg(self, service, org):
        """
        Добавляет услугу в список услуг и вызывает функцию распределения для других услуг
        """
        self.agent_dict[org].append(service)
        #Запись новой услуги в XML
        with open(self.agentXML, encoding='utf-8') as f:
            xml = f.read()
        root = objectify.fromstring(xml)
        # Перечисляем существующие организации в файле и добавляем новые строчки
        for x in root.UrFace:
            if x.get('Name') == org:
                Service = objectify.SubElement(x.Services, "Service")
                Service.set("Name", service)
        # удаляем аннотации.
        objectify.deannotate(root)
        etree.cleanup_namespaces(root)
        obj_xml = etree.tostring(root, encoding='utf-8', pretty_print=True, xml_declaration=True)
        # сохраняем данные в файл.
        try:
            with open(self.agentXML, "w", encoding='utf_8_sig') as xml_writer:
                xml_writer.write(obj_xml.decode('utf-8'))
        except IOError:
            pass
        self.distibution_agentservice()

    def show_dialog_add_agentorg(self, title, text, service):
        """
        Выводит диалоговое окно с возможность ввода имени новой организыции и двумя кнопками
        """
        content = MDTextField(hint_text="Persistent helper text222",
                              helper_text="Text is always here111",
                              helper_text_mode="persistent",
                              text=text,
                              )
        dialog = MDDialog(title=title,
                          content=content,
                          size_hint=(.8, None),
                          height=dp(200),
                          auto_dismiss=False)
        dialog.add_action_button("Отмена", action=lambda *x: (dialog.dismiss(), self.readd_agentorg(service)))
        dialog.add_action_button("Добавить",
                                 action=lambda *x: (dialog.dismiss(), self.create_new_agentorg(dialog.content.text, service)))
        dialog.open()

    def create_new_agentorg(self, name, service):
        """
        Добавляет новую организацию в список организаций self.orgs, словарь self.orgs_dict и XML конфигурацию.
        Возвращает изьятую ранее услугу в список новых услуг с помощью функции self.readd_org
        """
        self.agentorgs.append(name)
        self.agent_dict[name] = []
        self.readd_agentorg(service)
        with open(self.agentXML, encoding='utf-8') as f:
            xml = f.read()
        root = objectify.fromstring(xml)
        #Добавляем новые организации
        new_org = objectify.SubElement(root, "UrFace")
        new_org.set('Name', name)
        new_servs = objectify.SubElement(new_org, 'Services')
        new_serv = objectify.SubElement(new_servs, 'Service')
        new_serv.set('Name', 'Пустая обязательная категория')
        # удаляем аннотации.
        objectify.deannotate(root)
        etree.cleanup_namespaces(root)
        obj_xml = etree.tostring(root,
                                 encoding='utf-8',
                                 pretty_print=True,
                                 xml_declaration=True
                                 )
        # сохраняем данные в файл.
        try:
            with open(self.agentXML, "w", encoding='utf_8_sig') as xml_writer:
                xml_writer.write(obj_xml.decode('utf-8'))
        except IOError:
            pass


    def readd_agentorg(self, service):
        """
        Возвращает изьятую ранее услугу в список новых услуг, затем вызывает функцию распределения
        """
        self.new_agentservice.append(service)
        self.distibution_agentservice()

    def fin_report(self):
        """
        Форминует финансовый отчет в установленном формате
        :return - dict
        """
        logging.info(f'{str(datetime.now()):25}:    Формирование финансового отчета')
        self.finreport_dict = {}
        is_aquazona = None
        for key in self.orgs_dict:
            if key != 'Не учитывать':
                self.finreport_dict[key] = [0, 0.00]
                for serv in self.orgs_dict[key]:
                    try:
                        if key == 'Нулевые':
                            self.finreport_dict[key][0] += self.itog_report_org1[serv][0]
                            self.finreport_dict[key][1] += self.itog_report_org1[serv][1]
                        elif key == 'Дата':
                            self.finreport_dict[key][0] = self.itog_report_org1[serv][0]
                            self.finreport_dict[key][1] = self.itog_report_org1[serv][1]
                        elif serv == 'Депозит':
                            self.finreport_dict[key][1] += self.itog_report_org1[serv][1]
                        elif serv == 'Аквазона':
                            self.finreport_dict['Кол-во проходов'] = [self.itog_report_org1[serv][0], 0]
                            self.finreport_dict[key][1] += self.itog_report_org1[serv][1]
                            is_aquazona = True
                        else:
                            self.finreport_dict[key][0] += self.itog_report_org1[serv][0]
                            self.finreport_dict[key][1] += self.itog_report_org1[serv][1]
                    except KeyError:
                        pass
                    except TypeError:
                        pass
        if not is_aquazona:
            self.finreport_dict['Кол-во проходов'] = [0, 0.00]
        self.finreport_dict['Online Продажи'] = list(self.report_bitrix)
        # self.finreport_dict['ИТОГО'][1] -= self.finreport_dict['Депозит'][1]

    def agent_report(self):
        """
        Форминует отчет платежного агента в установленном формате
        :return - dict
        """
        logging.info(f'{str(datetime.now()):25}:    Формирование отчета платежного агента')
        self.agentreport_dict = {}
        for key in self.agent_dict:
            if key != 'Не учитывать':
                self.agentreport_dict[key] = [0, 0]
                for serv in self.agent_dict[key]:
                    try:
                        if key == 'Дата':
                            self.agentreport_dict[key][0] = self.itog_report_org1[serv][0]
                            self.agentreport_dict[key][1] = self.itog_report_org1[serv][1]
                        elif serv == 'Депозит':
                            self.agentreport_dict[key][1] += self.itog_report_org1[serv][1]
                        elif serv == 'Аквазона':
                            self.agentreport_dict[key][1] += self.itog_report_org1[serv][1]
                        else:
                            self.agentreport_dict[key][0] += self.itog_report_org1[serv][0]
                            self.agentreport_dict[key][1] += self.itog_report_org1[serv][1]
                    except KeyError:
                        pass
                    except TypeError:
                        pass

    def export_fin_report(self):
        """
        Сохраняет Финансовый отчет в виде Excel-файла в локальную директорию
        """
        font0 = xlwt.Font()
        font0.name = 'Arial'
        font0.colour_index = 0
        font0.bold = True

        style0 = xlwt.XFStyle()
        style0.font = font0

        style1 = xlwt.XFStyle()
        style1.num_format_str = '0'

        style2 = xlwt.XFStyle()
        style2.num_format_str = '0.00'

        wb = xlwt.Workbook()
        ws = wb.add_sheet('Фин отчет')
        col_width = 200 * 20

        try:
            for i in itertools.count():
                ws.col(i).width = col_width
        except ValueError:
            pass

        first_col = ws.col(0)
        first_col.width = 200 * 20
        second_col = ws.col(1)
        second_col.width = 200 * 20

        default_book_style = wb.default_style
        default_book_style.font.height = 20 * 36  # 36pt

        ws.write(0, 0, 'Дата', style0)
        ws.write(0, 1, 'Кол-во проходов', style0)
        ws.write(0, 2, 'Общая сумма', style0)
        ws.write(0, 3, 'Сумма KPI', style0)
        ws.write(0, 4, 'Билеты Кол-во', style0)
        ws.write(0, 5, 'Билеты Сумма', style0)
        ws.write(0, 6, 'Билеты Средний чек', style0)
        ws.write(0, 7, 'Общепит Кол-во', style0)
        ws.write(0, 8, 'Общепит Сумма', style0)
        ws.write(0, 9, 'Общепит Средний чек', style0)
        ws.write(0, 10, 'Прочее Кол-во', style0)
        ws.write(0, 11, 'Прочее Сумма', style0)
        ws.write(0, 12, 'Online Продажи Кол-во', style0)
        ws.write(0, 13, 'Online Продажи Сумма', style0)
        ws.write(0, 14, 'Online Продажи Средний чек', style0)
        ws.write(0, 15, 'Сумма безнал', style0)
        if self.finreport_dict['Дата'][0] + timedelta(1) == self.finreport_dict['Дата'][1]:
            date_ = datetime.strftime(self.finreport_dict["Дата"][0], "%Y-%m-%d")
        else:
            date_ = f'{datetime.strftime(self.finreport_dict["Дата"][0], "%Y-%m-%d")} - ' \
                    f'{datetime.strftime(self.finreport_dict["Дата"][1] - timedelta(1), "%Y-%m-%d")}'
        ws.write(1, 0, date_, style1)
        ws.write(1, 1, self.finreport_dict['Кол-во проходов'][0], style1)
        ws.write(1, 2, self.finreport_dict['ИТОГО'][1], style2)
        ws.write(1, 3, '=C2-L2+N2+P2+Q2', style2)
        ws.write(1, 4, self.finreport_dict['Билеты аквапарка'][0], style1)
        ws.write(1, 5, self.finreport_dict['Билеты аквапарка'][1], style2)
        ws.write(1, 6, '=ЕСЛИОШИБКА(F2/E2;0)', style2)
        ws.write(1, 7, self.finreport_dict['Общепит'][0], style1)
        ws.write(1, 8, self.finreport_dict['Общепит'][1], style2)
        ws.write(1, 9, '=ЕСЛИОШИБКА(I2/H2;0)', style2)
        ws.write(1, 10, self.finreport_dict['Прочее'][0], style1)
        ws.write(1, 11, self.finreport_dict['Прочее'][1], style2)
        ws.write(1, 12, self.finreport_dict['Online Продажи'][0], style1)
        ws.write(1, 13, self.finreport_dict['Online Продажи'][1], style2)
        ws.write(1, 14, '=ЕСЛИОШИБКА(N2/M2;0)', style2)
        path = self.local_folder + self.path_aquapark + date_ + '_Финансовый_отчет' + ".xls"
        logging.info(f'{str(datetime.now()):25}:    Сохранение финансового отчета в {path}')
        path = self.create_path(path)
        self.save_file(path, wb)
        return path

    def export_agent_report(self):
        """
        Сохраняет отчет платежного агента в виде Excel-файла в локальную директорию
        """
        font0 = xlwt.Font()
        font0.name = 'Arial'
        font0.colour_index = 0
        font0.bold = True

        style0 = xlwt.XFStyle()
        style0.font = font0

        style1 = xlwt.XFStyle()
        style1.num_format_str = '0'
        style1.borders.bottom = 1
        style1.borders.left = 1
        style1.borders.right = 1
        style1.borders.top = 1

        style2 = xlwt.XFStyle()
        style2.num_format_str = '0.00'
        style2.borders.bottom = 1
        style2.borders.left = 1
        style2.borders.right = 1
        style2.borders.top = 1

        style3 = xlwt.XFStyle()
        style3.font = font0
        style3.num_format_str = '0'
        style3.borders.bottom = 1
        style3.borders.left = 1
        style3.borders.right = 1
        style3.borders.top = 1

        style4 = xlwt.XFStyle()
        style4.font = font0
        style4.num_format_str = '0.00'
        style4.borders.bottom = 1
        style4.borders.left = 1
        style4.borders.right = 1
        style4.borders.top = 1

        wb = xlwt.Workbook()
        ws = wb.add_sheet('Отчет платежного агента')
        col_width = 200 * 20

        try:
            for i in itertools.count():
                ws.col(i).width = col_width
        except ValueError:
            pass

        first_col = ws.col(0)
        first_col.width = 700 * 20
        second_col = ws.col(1)
        second_col.width = 300 * 20

        default_book_style = wb.default_style
        default_book_style.font.height = 20 * 44  # 36pt

        if self.agentreport_dict['Дата'][0] + timedelta(1) == self.agentreport_dict["Дата"][1]:
            date_ = datetime.strftime(self.agentreport_dict["Дата"][0], "%Y-%m-%d")
            head = f'ОТЧЕТ ПЛАТЕЖНОГО АГЕНТА ПО ПРИЕМУ ДЕНЕЖНЫХ СРЕДСТВ ЗА ' \
                   f'{datetime.strftime(self.agentreport_dict["Дата"][0], "%d.%m.%Y")}г.'
        else:
            date_ = f'{datetime.strftime(agentreport_dict["Дата"][0], "%Y-%m-%d")} - ' \
                    f'{datetime.strftime(agentreport_dict["Дата"][1] - timedelta(1), "%Y-%m-%d")}'
            head = f'ОТЧЕТ ПЛАТЕЖНОГО АГЕНТА ПО ПРИЕМУ ДЕНЕЖНЫХ СРЕДСТВ ЗА ' \
                   f'{datetime.strftime(agentreport_dict["Дата"][0], "%d.%m.%Y")} - ' \
                   f'{datetime.strftime(agentreport_dict["Дата"][1] - timedelta(1), "%d.%m.%Y")}г.'
        ws.write(0, 0, head, style0)
        # ws.write(0, 1, date_, style1)
        ws.write(2, 0, 'Наименование поставщика услуг', style3)
        ws.write(2, 1, 'Сумма', style3)

        i = 3
        for key in self.agentreport_dict:
            if key != 'Дата':
                if key == 'ИТОГО':
                    ws.write(i, 0, key, style3)
                    ws.write(i, 1, self.agentreport_dict[key][1], style4)
                    i += 1
                else:
                    ws.write(i, 0, key, style1)
                    ws.write(i, 1, self.agentreport_dict[key][1], style2)
                    i += 1

        path = self.local_folder + self.path_aquapark + date_ + '_Отчет_платежного_агента' + ".xls"
        logging.info(f'{str(datetime.now()):25}:    Сохранение отчета платежного агента в {path}')
        path = self.create_path(path)
        self.save_file(path, wb)
        return path

    def create_path(self, path):
        """
        Проверяет наличие указанного пути. В случае отсутствия каких-либо папок создает их
        """
        logging.info(f'{str(datetime.now()):25}:    Проверка локальных путей сохранения файлов...')
        list_path = path.split('/')
        path = ''
        end_path = ''
        if list_path[-1][-4:] == '.xls' or list_path[-1]:
            end_path = list_path.pop()
        list_path.append(self.date_from.strftime('%Y'))
        list_path.append(self.date_from.strftime('%m') + '-' + self.date_from.strftime('%B'))
        directory = os.getcwd()
        for folder in list_path:
            if folder not in os.listdir():
                os.mkdir(folder)
                logging.warning(f'{str(datetime.now()):25}:    В директории "{os.getcwd()}" создана папка "{folder}"')
                os.chdir(folder)
            else:
                os.chdir(folder)
            path += folder + '/'
        path += end_path
        os.chdir(directory)
        return path

    def save_file(self, path, file):
        """
        Проверяет не занят ли файл другим процессом и если нет, то перезаписывает его, в противном
        случае выводит диалоговое окно с предложением закрыть файл и продолжить
        """
        try:
            file.save(path)
        except PermissionError as e:
            logging.error(f'{str(datetime.now()):25}:    Файл "{path}" занят другим процессом.\n{repr(e)}')
            self.show_dialog(f'Ошибка записи файла',
                             f'Файл "{path}" занят другим процессом.\nДля повтора попытки закройте это сообщение',
                             func=self.save_file, path=path, file=file)

    def sync_to_yadisk(self, local_path, token):
        """
        Копирует локальные файлы в Яндекс Диск
        """
        logging.info(f'{str(datetime.now()):25}:    Копирование отчета {local_path} в Яндекс.Диск...')
        if self.use_yadisk:
            logging.info(f'{str(datetime.now()):25}:    Соединение с YaDisk...')
            self.yadisk = yadisk.YaDisk(token=token)
            if self.yadisk.check_token():
                path = '' + self.path_aquapark
                remote_folder = self.create_path_yadisk(path)
                remote_path = remote_folder + local_path.split('/')[-1]
                logging.info(f'{str(datetime.now()):25}:    Отправка файла "{local_path.split("/")[-1]}" в YaDisk...')
                files_list_yandex = list(self.yadisk.listdir(remote_folder))
                files_list = []
                for key in files_list_yandex:
                    if key['file']:
                        files_list.append(remote_folder + key['name'])
                if remote_path not in files_list:
                    self.yadisk.upload(local_path, remote_path)
                    logging.info(
                        f'{str(datetime.now()):25}:    '
                        f'Файл "{local_path.split("/")[-1]}" отправлен в "{remote_folder}" YaDisk...')
                else:
                    logging.warning(
                        f'{str(datetime.now()):25}:    '
                        f'Файл "{local_path.split("/")[-1]}" уже существует в "{remote_folder}"')
                    def rewrite_file():
                        self.yadisk.remove(remote_path, permanently=True)
                        self.yadisk.upload(local_path, remote_path)
                        logging.warning(
                            f'{str(datetime.now()):25}:    '
                            f'Файл "{local_path.split("/")[-1]}" успешно обновлен')
                    if self.root.ids.report.ids.split_by_days.active:
                        rewrite_file()
                    else:
                        self.show_dialog_variant('Файл уже существует',
                                             f'Файл "{local_path.split("/")[-1]}" уже существует в "{remote_folder}"'
                                             f'\nЗаменить?',
                                             rewrite_file
                                             )
            else:
                logging.error(f'{str(datetime.now()):25}:    Ошибка YaDisk: token не валиден')
                self.show_dialog('Ошибка соединения с Yandex.Disc',
                                 f'\nОтчеты сохранены в папке {self.local_folder} '
                                 f'и не будут отправлены на Yandex.Disc.')

    def create_path_yadisk(self, path):
        """
        Проверяет наличие указанного пути в Яндекс Диске. В случае отсутствия каких-либо папок создает их
        :param path:
        :return:
        """
        logging.info(f'{str(datetime.now()):25}:    Проверка путей сохранения файлов на Яндекс.Диске...')
        list_path = path.split('/')
        path = ''
        end_path = ''
        if list_path[-1][-4:] == '.xls' or list_path[-1] == '':
            end_path = list_path.pop()
        list_path.append(self.date_from.strftime('%Y'))
        list_path.append(self.date_from.strftime('%m') + '-' + self.date_from.strftime('%B'))
        directory = '/'
        list_path_yandex = []
        for folder in list_path:
            folder = directory + folder
            directory = folder + '/'
            list_path_yandex.append(folder)
        directory = '/'
        for folder in list_path_yandex:
            folders_list = []
            folders_list_yandex = list(self.yadisk.listdir(directory))
            for key in folders_list_yandex:
                if not key['file']:
                    folders_list.append(directory + key['name'])
            if folder not in folders_list:
                self.yadisk.mkdir(folder)
                logging.info(f'{str(datetime.now()):25}:    Создание новой папки в YandexDisk - "{folder}"')
                directory = folder + '/'
            else:
                directory = folder + '/'
        path = list_path_yandex[-1] + '/'
        return path

    def export_to_google_sheet(self):
        """
        Формирование и заполнение google-таблицы
        """
        logging.info(f'{str(datetime.now()):25}:    Сохранение Финансового отчета в Google-таблицах...')

        #self.CREDENTIALS_FILE # имя файла с закрытым ключом

        self.sheet_width = 17
        height = 35

        credentials = ServiceAccountCredentials.from_json_keyfile_name(self.CREDENTIALS_FILE,
                                                                       ['https://www.googleapis.com/auth/spreadsheets',
                                                                        'https://www.googleapis.com/auth/drive'])
        httpAuth = credentials.authorize(httplib2.Http())
        self.googleservice = apiclient.discovery.build('sheets', 'v4', http=httpAuth)

        data_report = datetime.strftime(self.finreport_dict['Дата'][0], '%m')
        month = [
            '',
            'Январь',
            'Ферраль',
            'Март',
            'Апрель',
            'Май',
            'Июнь',
            'Июль',
            'Август',
            'Сентябрь',
            'Октябрь',
            'Ноябрь',
            'Декабрь',
        ]
        data_report = month[int(data_report)]

        doc_name = f"Итоговый отчет по Аквапарку - {data_report} {datetime.strftime(self.finreport_dict['Дата'][0], '%Y')}"

        if self.finreport_dict['Дата'][0] + timedelta(1) != self.finreport_dict['Дата'][1]:
            logging.info(f'{str(datetime.now()):25}:    Экспорт отчета в Google Sheet за несколько дней невозможен!')
            self.show_dialog('Ошибка экспорта в Google.Sheet', 'Экспорт отчета в Google Sheet за несколько дней невозможен!')
        else:
            with open(self.list_google_docs, 'r', encoding='utf-8') as f:
                links = csv.reader(f, delimiter=';')
                self.google_links = {}
                for line in links:
                    self.google_links[line[0]] = line[1]
            if self.date_from.strftime('%Y-%m') in self.google_links:
                self.google_doc = (self.date_from.strftime('%Y-%m'), self.google_links[self.date_from.strftime('%Y-%m')])
            else:
                self.google_doc = None
                # Создание документа
                self.spreadsheet = self.googleservice.spreadsheets().create(body={
                    'properties': {'title': doc_name, 'locale': 'ru_RU'},
                    'sheets': [{'properties': {'sheetType': 'GRID',
                                               'sheetId': 0,
                                               'title': 'Сводный',
                                               'gridProperties': {'rowCount': height, 'columnCount': self.sheet_width}}},
                               {'properties': {'sheetType': 'GRID',
                                               'sheetId': 1,
                                               'title': 'Расширенный',
                                               'gridProperties': {'rowCount': height, 'columnCount': 100}}}
                               ]
                }).execute()

                # Доступы к документу
                self.google_reader_list = self.google_reader_list.split(',')
                self.google_writer_list = self.google_writer_list.split(',')
                driveService = apiclient.discovery.build('drive', 'v3', http=httpAuth)
                if self.google_all_read:
                    shareRes = driveService.permissions().create(
                        fileId=self.spreadsheet['spreadsheetId'],
                        body={'type': 'anyone', 'role': 'reader'},  # доступ на чтение кому угодно
                        fields='id'
                    ).execute()
                # Возможные значения writer, commenter, reader
                # доступ на Чтение определенным пользователоям
                for adress in self.google_reader_list:
                    shareRes = driveService.permissions().create(
                        fileId=self.spreadsheet['spreadsheetId'],
                        body={'type': 'user', 'role': 'reader', 'emailAddress': adress},
                        fields='id'
                    ).execute()
                # доступ на Запись определенным пользователоям
                for adress in self.google_writer_list:
                    shareRes = driveService.permissions().create(
                        fileId=self.spreadsheet['spreadsheetId'],
                        body={'type': 'user', 'role': 'writer', 'emailAddress': adress},
                        fields='id'
                    ).execute()

                sheetId = 0

                # Ширина столбцов
                ss = to_google_sheets.Spreadsheet(self.spreadsheet['spreadsheetId'], sheetId, self.googleservice,
                                                  self.spreadsheet['sheets'][sheetId]['properties']['title'])
                ss.prepare_setColumnWidth(0, 100)
                ss.prepare_setColumnWidth(1, 120)
                ss.prepare_setColumnsWidth(2, 3, 120)
                ss.prepare_setColumnWidth(4, 65)
                ss.prepare_setColumnWidth(5, 120)
                ss.prepare_setColumnWidth(6, 100)
                ss.prepare_setColumnWidth(7, 65)
                ss.prepare_setColumnWidth(8, 120)
                ss.prepare_setColumnWidth(9, 100)
                ss.prepare_setColumnWidth(10, 65)
                ss.prepare_setColumnWidth(11, 120)
                ss.prepare_setColumnWidth(12, 65)
                ss.prepare_setColumnWidth(13, 120)
                ss.prepare_setColumnWidth(14, 100)
                ss.prepare_setColumnWidth(15, 120)
                ss.prepare_setColumnWidth(16, 120)

                # Объединение ячеек
                ss.prepare_mergeCells("A1:A2")
                ss.prepare_mergeCells("B1:B2")
                ss.prepare_mergeCells("C1:C2")
                ss.prepare_mergeCells("D1:D2")
                ss.prepare_mergeCells("E1:G1")
                ss.prepare_mergeCells("H1:J1")
                ss.prepare_mergeCells("K1:L1")
                ss.prepare_mergeCells("M1:O1")
                ss.prepare_mergeCells("P1:P2")
                ss.prepare_mergeCells("Q1:Q2")

                # Задание параметров группе ячеек
                # Жирный, по центру
                ss.prepare_setCellsFormat('A1:Q2', {'horizontalAlignment': 'CENTER', 'textFormat': {'bold': True}})
                # ss.prepare_setCellsFormat('E4:E8', {'numberFormat': {'pattern': '[h]:mm:ss', 'type': 'TIME'}},
                #                           fields='userEnteredFormat.numberFormat')

                # Заполнение таблицы
                ss.prepare_setValues("A1:Q2", [["Дата", "Кол-во проходов", "Общая сумма", "Сумма KPI", "Билеты", "",
                                                "", "Общепит", "", "", "Прочее", "", "Online Продажи", "", "",
                                                "Сумма безнал", "Сумма Biglion"],
                                               ["", "", "", "", "Кол-во", "Сумма", "Средний чек", "Кол-во", "Сумма",
                                                "Средний чек", "Кол-во", "Сумма", "Кол-во", "Сумма", "Средний чек", "",
                                                ""]],
                                     "ROWS")
                # ss.prepare_setValues("D5:E6", [["This is D5", "This is D6"], ["This is E5", "=5+5"]], "COLUMNS")

                # Цвет фона ячеек
                ss.prepare_setCellsFormat("A1:Q2", {"backgroundColor": functions.htmlColorToJSON("#f7cb4d")},
                                          fields="userEnteredFormat.backgroundColor")

                # Бордер
                for i in range(2):
                    for j in range(self.sheet_width):
                        ss.requests.append({"updateBorders": {
                            "range": {"sheetId": ss.sheetId, "startRowIndex": i, "endRowIndex": i + 1,
                                      "startColumnIndex": j,
                                      "endColumnIndex": j + 1},
                            "top": {"style": "SOLID", "width": 1, "color": {"red": 0, "green": 0, "blue": 0}}}})
                        ss.requests.append({"updateBorders": {
                            "range": {"sheetId": ss.sheetId, "startRowIndex": i, "endRowIndex": i + 1,
                                      "startColumnIndex": j,
                                      "endColumnIndex": j + 1},
                            "right": {"style": "SOLID", "width": 1,
                                      "color": {"red": 0, "green": 0, "blue": 0, "alpha": 1.0}}}})
                        ss.requests.append({"updateBorders": {
                            "range": {"sheetId": ss.sheetId, "startRowIndex": i, "endRowIndex": i + 1,
                                      "startColumnIndex": j,
                                      "endColumnIndex": j + 1},
                            "left": {"style": "SOLID", "width": 1,
                                      "color": {"red": 0, "green": 0, "blue": 0, "alpha": 1.0}}}})

                ss.runPrepared()

                self.google_doc = (self.date_from.strftime('%Y-%m'), self.spreadsheet['spreadsheetId'])
                self.google_links[self.google_doc[0]] = self.google_doc[1]
                links = []
                for docid in self.google_links:
                    links.append([docid, self.google_links[docid]])
                with open(self.list_google_docs, 'w', newline='', encoding='utf-8') as f:
                    file = csv.writer(f, delimiter=';')
                    for link in links:
                        file.writerow(link)
                logging.info(
                    f'{str(datetime.now()):25}:    Создана новая таблица с Id: {self.spreadsheet["spreadsheetId"]}')

            self.spreadsheet = self.googleservice.spreadsheets().get(spreadsheetId=self.google_doc[1], ranges=[],
                                                     includeGridData=True).execute()

            # -------------------------------- ЗАПОЛНЕНИЕ ДАННЫМИ ------------------------------------------------

            # Печать таблицы в консоль
            # s = ''
            # for line_table in spreadsheet['sheets'][0]['data'][0]['rowData']:
            #     for cell in line_table['values']:
            #         try:
            #             s += cell['formattedValue'] + " | "
            #         except KeyError:
            #             pass
            #     print(s)
            #     s = ''

            # Проверка нет ли текущей даты в таблице
            self.nex_line = 1
            self.reprint = 2

            for line_table in self.spreadsheet['sheets'][0]['data'][0]['rowData']:
                try:
                    if line_table['values'][0]['formattedValue'] == datetime.strftime(self.finreport_dict['Дата'][0],
                                                                                      '%d.%m.%Y'):
                        if self.root.ids.report.ids.split_by_days.active:
                            self.rewrite_google_sheet()
                        else:
                            self.show_dialog_variant(f'Перезаписать эту строку?',
                                                     f'Строка за '
                                                     f'{datetime.strftime(self.finreport_dict["Дата"][0], "%d.%m.%Y")} '
                                                     f'уже существует в таблице!',
                                                     self.rewrite_google_sheet,
                                                     )
                        self.reprint = 0
                        break
                    elif line_table['values'][0]['formattedValue'] == "ИТОГО":
                        break
                    else:
                        self.nex_line += 1
                except KeyError:
                    self.nex_line += 1
            if self.reprint:
                self.write_google_sheet()

            # width_table = len(self.spreadsheet['sheets'][0]['data'][0]['rowData'][0]['values'])

    def rewrite_google_sheet(self):
        """
        Заполнение google-таблицы в случае, если данные уже существуют
        """
        logging.warning(f'{str(datetime.now()):25}:    Перезапись уже существующей строки...')
        self.reprint = 1
        self.write_google_sheet()

    def write_google_sheet(self):
        """
        Заполнение google-таблицы
        """
        sheetId = 0
        ss = to_google_sheets.Spreadsheet(self.spreadsheet['spreadsheetId'], sheetId, self.googleservice,
                                          self.spreadsheet['sheets'][sheetId]['properties']['title'])

        # Заполнение строки с данными
        ss.prepare_setValues(f"A{self.nex_line}:Q{self.nex_line}",
                             [[datetime.strftime(self.finreport_dict['Дата'][0], '%d.%m.%Y'),
                               f"{self.finreport_dict['Кол-во проходов'][0]}",
                               self.finreport_dict['ИТОГО'][1],
                               f"=C{self.nex_line}-L{self.nex_line}+N{self.nex_line}+P{self.nex_line}+Q{self.nex_line}",
                               self.finreport_dict['Билеты аквапарка'][0],
                               self.finreport_dict['Билеты аквапарка'][1],
                               f"=IFERROR(F{self.nex_line}/E{self.nex_line};0)",
                               self.finreport_dict['Общепит'][0],
                               self.finreport_dict['Общепит'][1],
                               f"=IFERROR(I{self.nex_line}/H{self.nex_line};0)",
                               self.finreport_dict['Прочее'][0],
                               self.finreport_dict['Прочее'][1],
                               self.finreport_dict['Online Продажи'][0],
                               self.finreport_dict['Online Продажи'][1],
                               f"=IFERROR(N{self.nex_line}/M{self.nex_line};0)",
                               0,
                               0,
                               ]],
                             "ROWS")

        # Задание форматы вывода строки
        ss.prepare_setCellsFormats(f"A{self.nex_line}:Q{self.nex_line}", [[{'numberFormat': {'type': 'DATE',
                                                                                   'pattern': 'dd.mm.yyyy'}},
                                                                 {'numberFormat': {}},
                                                                 {'numberFormat': {'type': 'CURRENCY',
                                                                                   'pattern': '#,##0.00[$ ₽]'}},
                                                                 {'numberFormat': {'type': 'CURRENCY',
                                                                                   'pattern': '#,##0.00[$ ₽]'}},
                                                                 {'numberFormat': {}},
                                                                 {'numberFormat': {'type': 'CURRENCY',
                                                                                   'pattern': '#,##0.00[$ ₽]'}},
                                                                 {'numberFormat': {'type': 'CURRENCY',
                                                                                   'pattern': '#,##0.00[$ ₽]'}},
                                                                 {'numberFormat': {}},
                                                                 {'numberFormat': {'type': 'CURRENCY',
                                                                                   'pattern': '#,##0.00[$ ₽]'}},
                                                                 {'numberFormat': {'type': 'CURRENCY',
                                                                                   'pattern': '#,##0.00[$ ₽]'}},
                                                                 {'numberFormat': {}},
                                                                 {'numberFormat': {'type': 'CURRENCY',
                                                                                   'pattern': '#,##0.00[$ ₽]'}},
                                                                 {'numberFormat': {}},
                                                                 {'numberFormat': {'type': 'CURRENCY',
                                                                                   'pattern': '#,##0.00[$ ₽]'}},
                                                                 {'numberFormat': {'type': 'CURRENCY',
                                                                                   'pattern': '#,##0.00[$ ₽]'}},
                                                                 {'numberFormat': {'type': 'CURRENCY',
                                                                                   'pattern': '#,##0.00[$ ₽]'}},
                                                                 {'numberFormat': {'type': 'CURRENCY',
                                                                                   'pattern': '#,##0.00[$ ₽]'}}]])

        # Цвет фона ячеек
        if self.nex_line % 2 != 0:
            ss.prepare_setCellsFormat(f"A{self.nex_line}:Q{self.nex_line}",
                                      {"backgroundColor": functions.htmlColorToJSON("#fef8e3")},
                                      fields="userEnteredFormat.backgroundColor")

        # Бордер
        for j in range(self.sheet_width):
            ss.requests.append({"updateBorders": {
                "range": {"sheetId": ss.sheetId, "startRowIndex": self.nex_line - 1, "endRowIndex": self.nex_line,
                          "startColumnIndex": j,
                          "endColumnIndex": j + 1},
                "top": {"style": "SOLID", "width": 1, "color": {"red": 0, "green": 0, "blue": 0}}}})
            ss.requests.append({"updateBorders": {
                "range": {"sheetId": ss.sheetId, "startRowIndex": self.nex_line - 1, "endRowIndex": self.nex_line,
                          "startColumnIndex": j,
                          "endColumnIndex": j + 1},
                "right": {"style": "SOLID", "width": 1,
                          "color": {"red": 0, "green": 0, "blue": 0, "alpha": 1.0}}}})
            ss.requests.append({"updateBorders": {
                "range": {"sheetId": ss.sheetId, "startRowIndex": self.nex_line - 1, "endRowIndex": self.nex_line,
                          "startColumnIndex": j,
                          "endColumnIndex": j + 1},
                "left": {"style": "SOLID", "width": 1,
                         "color": {"red": 0, "green": 0, "blue": 0, "alpha": 1.0}}}})
            ss.requests.append({"updateBorders": {
                "range": {"sheetId": ss.sheetId, "startRowIndex": self.nex_line - 1, "endRowIndex": self.nex_line,
                          "startColumnIndex": j,
                          "endColumnIndex": j + 1},
                "bottom": {"style": "SOLID", "width": 1,
                           "color": {"red": 0, "green": 0, "blue": 0, "alpha": 1.0}}}})

        ss.runPrepared()

        # ------------------------------------------- Заполнение ИТОГО --------------------------------------

        # Вычисление последней строки в таблице
        for i, line_table in enumerate(self.spreadsheet['sheets'][0]['data'][0]['rowData']):
            try:
                if line_table['values'][0]['formattedValue'] == "ИТОГО":
                    # Если строка переписывается - итого на 1 поз вниз, если новая - на 2 поз
                    height_table = i + self.reprint
                    break
                else:
                    height_table = 4
            except KeyError:
                pass

        #
        ss.prepare_setValues(f"A{height_table}:Q{height_table}",
                             [[f'ИТОГО',
                               f"=SUM(B3:B{height_table - 1})",
                               f"=SUM(C3:C{height_table - 1})",
                               f"=SUM(D3:D{height_table - 1})",
                               f"=SUM(E3:E{height_table - 1})",
                               f"=SUM(F3:F{height_table - 1})",
                               f"=IFERROR(F{height_table}/E{height_table};0)",
                               f"=SUM(H3:H{height_table - 1})",
                               f"=SUM(I3:i{height_table - 1})",
                               f"=IFERROR(I{height_table}/H{height_table};0)",
                               f"=SUM(K3:K{height_table - 1})",
                               f"=SUM(L3:L{height_table - 1})",
                               f"=SUM(M3:M{height_table - 1})",
                               f"=SUM(N3:N{height_table - 1})",
                               f"=IFERROR(N{height_table}/M{height_table};0)",
                               f"=SUM(P3:P{height_table - 1})",
                               f"=SUM(Q3:Q{height_table - 1})",
                               ]],
                             "ROWS")

        # Задание форматы вывода строки
        ss.prepare_setCellsFormats(f"A{height_table}:Q{height_table}", [[{'numberFormat': {}},
                                                                         {'numberFormat': {}},
                                                                         {'numberFormat': {'type': 'CURRENCY',
                                                                                           'pattern': '#,##0.00[$ ₽]'}},
                                                                         {'numberFormat': {'type': 'CURRENCY',
                                                                                           'pattern': '#,##0.00[$ ₽]'}},
                                                                         {'numberFormat': {}},
                                                                         {'numberFormat': {'type': 'CURRENCY',
                                                                                           'pattern': '#,##0.00[$ ₽]'}},
                                                                         {'numberFormat': {'type': 'CURRENCY',
                                                                                           'pattern': '#,##0.00[$ ₽]'}},
                                                                         {'numberFormat': {}},
                                                                         {'numberFormat': {'type': 'CURRENCY',
                                                                                           'pattern': '#,##0.00[$ ₽]'}},
                                                                         {'numberFormat': {'type': 'CURRENCY',
                                                                                           'pattern': '#,##0.00[$ ₽]'}},
                                                                         {'numberFormat': {}},
                                                                         {'numberFormat': {'type': 'CURRENCY',
                                                                                           'pattern': '#,##0.00[$ ₽]'}},
                                                                         {'numberFormat': {}},
                                                                         {'numberFormat': {'type': 'CURRENCY',
                                                                                           'pattern': '#,##0.00[$ ₽]'}},
                                                                         {'numberFormat': {'type': 'CURRENCY',
                                                                                           'pattern': '#,##0.00[$ ₽]'}},
                                                                         {'numberFormat': {'type': 'CURRENCY',
                                                                                           'pattern': '#,##0.00[$ ₽]'}},
                                                                         {'numberFormat': {'type': 'CURRENCY',
                                                                                           'pattern': '#,##0.00[$ ₽]'}}]])

        ss.prepare_setCellsFormat(f"A{height_table}:Q{height_table}",
                                  {'horizontalAlignment': 'RIGHT', 'textFormat': {'bold': True}})

        # Цвет фона ячеек
        ss.prepare_setCellsFormat(f"A{height_table}:Q{height_table}",
                                  {"backgroundColor": functions.htmlColorToJSON("#fce8b2")},
                                  fields="userEnteredFormat.backgroundColor")

        # Бордер
        for j in range(self.sheet_width):
            ss.requests.append({"updateBorders": {
                "range": {"sheetId": ss.sheetId, "startRowIndex": height_table - 1, "endRowIndex": height_table,
                          "startColumnIndex": j,
                          "endColumnIndex": j + 1},
                "top": {"style": "SOLID", "width": 1, "color": {"red": 0, "green": 0, "blue": 0}}}})
            ss.requests.append({"updateBorders": {
                "range": {"sheetId": ss.sheetId, "startRowIndex": height_table - 1, "endRowIndex": height_table,
                          "startColumnIndex": j,
                          "endColumnIndex": j + 1},
                "right": {"style": "SOLID", "width": 1,
                          "color": {"red": 0, "green": 0, "blue": 0, "alpha": 1.0}}}})
            ss.requests.append({"updateBorders": {
                "range": {"sheetId": ss.sheetId, "startRowIndex": height_table - 1, "endRowIndex": height_table,
                          "startColumnIndex": j,
                          "endColumnIndex": j + 1},
                "left": {"style": "SOLID", "width": 1,
                         "color": {"red": 0, "green": 0, "blue": 0, "alpha": 1.0}}}})
            ss.requests.append({"updateBorders": {
                "range": {"sheetId": ss.sheetId, "startRowIndex": height_table - 1, "endRowIndex": height_table,
                          "startColumnIndex": j,
                          "endColumnIndex": j + 1},
                "bottom": {"style": "SOLID", "width": 1,
                           "color": {"red": 0, "green": 0, "blue": 0, "alpha": 1.0}}}})

        ss.runPrepared()

    def open_googlesheet(self):
        """
        Открывает брацзев с текущей гугл-таблицей
        """
        if not self.open_browser:
            logging.info(f'{str(datetime.now()):25}:    Открытие файла-отчета в браузере...')
            self.show_dialog_variant(f'Открыть Google-отчет?',
                                     'Открыть Google-отчет?',
                                     webbrowser.open,
                                     self.spreadsheet['spreadsheetUrl']
                                     )
            self.open_browser = True

    def sms_report(self):
        """
        Составляет текстовую версию финансового отчета
        :return: str
        """
        logging.info(f'{str(datetime.now()):25}:    Составление SMS-отчета...')
        resporse = 'Отчет по аквапарку за '
        if self.finreport_dict['Дата'][0] + timedelta(1) == self.finreport_dict['Дата'][1]:
            resporse += f'{datetime.strftime(self.finreport_dict["Дата"][0], "%d.%m.%Y")}:\n'
        else:
            resporse += f'{datetime.strftime(self.finreport_dict["Дата"][0], "%d.%m.%Y")} - {datetime.strftime(aqua_report["Дата"][1], "%d.%m.%Y")}:\n'
        if self.finreport_dict['ИТОГО'][1]:
            resporse += f'Люди - {self.finreport_dict["Кол-во проходов"][0]};\n'
            resporse += f'По аквапарку - {self.finreport_dict["Билеты аквапарка"][1]:.2f} ₽;\n'
            resporse += f'По общепиту - {self.finreport_dict["Общепит"][1]:.2f} ₽;\n'
            resporse += f'Термозона - {self.finreport_dict["Термозона"][1]:.2f} ₽;\n'
            resporse += f'Прочее - {self.finreport_dict["Прочее"][1]:.2f} ₽;\n'
            resporse += f'Общая по БАРСу - {self.finreport_dict["ИТОГО"][1]:.2f} ₽;\n'
            resporse += f'ONLINE продажи - {self.finreport_dict["Online Продажи"][1]:.2f} ₽;\n'
        if self.itog_report_org2['Итого по отчету'][1]:
            # resporse += 'Отчет по пляжу за '
            # if beach_report['Дата'][0] + timedelta(1) == beach_report['Дата'][1]:
            #     resporse += f'{datetime.strftime(beach_report["Дата"][0], "%d.%m.%Y")}:\n'
            # else:
            #     resporse += f'{datetime.strftime(beach_report["Дата"][0], "%d.%m.%Y")} - {datetime.strftime(beach_report["Дата"][0], "%d.%m.%Y")}:\n'
            resporse += f'Люди (пляж) - {self.itog_report_org2["Летняя зона | БЕЗЛИМИТ | 1 проход"][0]};\n'
            resporse += f'Итого по пляжу - {self.itog_report_org2["Итого по отчету"][1]:.2f} ₽;\n'
        resporse += f'Без ЧП.'
        return resporse

    def send_message_to_telegram(self):
        """
        Отправка отчета в telegram
        """
        logging.info(f'{str(datetime.now()):25}:    Отправка SMS-отчета в Telegram-канал...')
        SetProxy = telepot.api.set_proxy(self.telegram_proxy, basic_auth=self.telegram_basic_auth)
        bot = telepot.Bot(self.telegram_token)
        bot.sendMessage(self.telegram_chanel_id, self.sms_report())

    def load_checkbox(self):
        """
        Установка чекбоксов в соответствии с настройками INI-файла
        """
        logging.info(f'{str(datetime.now()):25}:    Загрузка настроек...')
        self.root.ids.report.ids.split_by_days.active = self.split_by_days
        self.root.ids.report.ids.finreport_xls.active = self.finreport_xls
        self.root.ids.report.ids.agentreport_xls.active = self.agentreport_xls
        self.root.ids.report.ids.use_yadisk.active = self.use_yadisk
        self.root.ids.report.ids.finreport_google.active = self.finreport_google
        self.root.ids.report.ids.finreport_telegram.active = self.finreport_telegram

    def change_checkbox(self, name, checkbox):
        """
        Изменяет состояние элемента конфигурации и записывает в INI-файл
        :param name: Имя чекбокса
        :param checkbox: Состояние active чекбокса
        """
        self.config.set('General', name, str(checkbox))
        setattr(self, name, checkbox)
        self.config.write()
        logging.info(f'{str(datetime.now()):25}:    Параметр {name} изменен на значение {checkbox}')
        if name == 'split_by_days' and not checkbox and not self.root.ids.report.ids.date_switch.active:
            self.root.ids.report.ids.finreport_google.active = False
            self.change_checkbox('finreport_google', False)
            self.root.ids.report.ids.finreport_google.disabled = True
            self.root.ids.report.ids.finreport_google_text.disabled = True
        elif name == 'split_by_days' and checkbox:
            self.root.ids.report.ids.finreport_google_text.disabled = False
            self.root.ids.report.ids.finreport_google.disabled = False

    def save_reports(self):
        """
        Функция управления
        """
        self.fin_report()
        self.agent_report()
        if self.finreport_xls:
            fin_report_path = self.export_fin_report()
            if self.use_yadisk:
                self.sync_to_yadisk(fin_report_path, self.yadisk_token)
        if self.agentreport_xls:
            agent_report_path = self.export_agent_report()
            if self.use_yadisk:
                self.sync_to_yadisk(agent_report_path, self.yadisk_token)
        if self.finreport_google:
            self.export_to_google_sheet()
            self.open_googlesheet()
        if self.finreport_telegram:
            self.send_message_to_telegram()

    def load_report(self):
        """
        Выполнить отчеты
        """
        self.itog_report_org1 = None
        self.itog_report_org2 = None
        self.report_bitrix = None

        self.click_select_org()

        if self.org1:
            self.itog_report_org1 = self.itog_report(
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
            self.itog_report_org2 = self.itog_report(
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

        self.report_bitrix = self.read_bitrix_base(
            server=self.server,
            database=self.database_bitrix,
            user=self.user,
            pwd=self.pwd,
            driver=self.driver,
            date_from=self.date_from,
            date_to=self.date_to,
        )
        # Чтение XML с привязкой групп услуг к услугам
        self.orgs_dict = self.read_reportgroup(self.reportXML)
        # Поиск новых услуг
        self.find_new_service(self.itog_report_org1, self.orgs_dict)
        self.distibution_service()

    def run_report(self):
        self.open_browser = False

        if self.date_switch:
            self.load_report()
            pass
        else:
            if self.split_by_days:
                period = []
                while True:
                    period.append(self.date_from)
                    if self.date_from + timedelta(1) == self.date_to:
                        break
                    else:
                        self.date_from = self.date_from + timedelta(1)
                for date in period:
                    self.date_from = date
                    self.date_to = date + timedelta(1)
                    self.load_report()
                self.date_from = datetime.strptime(self.root.ids.report.ids.date_from.text, "%Y-%m-%d")
            else:
                self.load_report()

if __name__ == '__main__':
    pass
