# /usr/bin/python3
# -*- coding: utf-8 -*-

# This file created with KivyCreatorProject
# <https://github.com/HeaTTheatR/KivyCreatorProgect
#
# Copyright © 2017 Easy
#
# For suggestions and questions:
# <kivydevelopment@gmail.com>
# 
# LICENSE: MIT

# Точка входа в приложение. Запускает основной программный код program.py.
# В случае ошибки, выводит на экран окно с её текстом.

import os
import sys
import traceback
import logging
import sentry_sdk

logging.basicConfig(filename="barsicreport2.log", level=logging.INFO)

# Никнейм и имя репозитория на github,
# куда будет отправлен отчёт баг репорта.
NICK_NAME_AND_NAME_REPOSITORY = 'sendhello/Barsic.Report'

directory = os.path.split(os.path.abspath(sys.argv[0]))[0]
sys.path.insert(0, os.path.join(directory, 'libs/applibs'))

try:
    import webbrowser
    try:
        import six.moves.urllib
    except ImportError:
        pass

    import kivy
    kivy.require('1.9.2')

    from kivy.config import Config
    Config.set('kivy', 'keyboard_mode', 'system')
    Config.set('kivy', 'log_enable', 0)

    from kivy import platform
    if platform == 'android':
        from plyer import orientation
        orientation.set_sensor(mode='any')

    from kivymd.theming import ThemeManager
    # Activity баг репорта.
    from bugreporter import BugReporter
except Exception:
    traceback.print_exc(file=open(os.path.join(directory, 'error.log'), 'w'))
    sys.exit(1)


__version__ = 'v2.5.3'


def main():
    def create_error_monitor():
        class _App(App):
            theme_cls = ThemeManager()
            theme_cls.primary_palette = 'BlueGrey'

            def build(self):
                box = BoxLayout()
                box.add_widget(report)
                return box
        app = _App()
        app.run()

    app = None

    try:
        from loadplugin import load_plugin # функция загрузки плагинов
        from barsicreport2 import BarsicReport2  # основной класс программы

        # Запуск приложения.
        app = BarsicReport2()
        load_plugin(app, __version__)
        app.run()
    except Exception:
        from kivy.app import App
        from kivy.uix.boxlayout import BoxLayout


        text_error = traceback.format_exc()
        traceback.print_exc(file=open(os.path.join(directory, 'error.log'), 'w'))

        if app:
            try:
                app.stop()
            except AttributeError:
                app = None

        def callback_report(*args):
            '''Функция отправки баг-репорта.'''

            try:
                txt = six.moves.urllib.parse.quote(
                    report.txt_traceback.text.encode('utf-8')
                )
                url = f'https://github.com/{NICK_NAME_AND_NAME_REPOSITORY}/issues/new?body=' + txt
                webbrowser.open(url)
            except Exception:
                sys.exit(1)

        report = BugReporter(
            callback_report=callback_report, txt_report=text_error,
            icon_background=os.path.join('data', 'images', 'icon.png')
        )

        if app:
            try:
                app.screen.clear_widgets()
                app.screen.add_widget(report)
            except AttributeError:
            	create_error_monitor()
        else:
            create_error_monitor()


if __name__ in ('__main__', '__android__'):
    sentry_sdk.init(
        'https://ada8bc9717e344f7a0f6c3e186c0fb8c@o412552.ingest.sentry.io/5445543',
        traces_sample_rate=1.0
    )
    try:
        main()
    except Exception as e:
        with sentry_sdk.push_scope() as scope:
            # scope.set_context("state", {})
            # scope.set_tag("state_machine_name", '')
            sentry_sdk.capture_exception(e)
        raise e
