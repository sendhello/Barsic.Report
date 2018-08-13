"""
loadplugin.py
+++++++++++++++++++++++++++++++++++

Загружает плагины программы из папки plugins корневой директории программы.

:Автор: Virtuos86

Каждый плагин представляет из себя Python подпакет, находящийся
в пакете **plugins**.
Плагины подключаются путем последовательного импортирования их,
порядок на данный момент не установлен.

Структура простейшего плагина
+++++++++++++++++++++++++++++++++++
::

    plugins/
    |
    +-- __init__.py
    |
    +-- <plugin>
    |   |
    |   |-- __init__.py
    |   |
    |   +-- manifest.txt


Структура манифеста - manifest.txt
+++++++++++++++++++++++++++++++++++

    **App-Version-Min:** 0.0

    **App-Version-Max:** 1

    **Plugin-Name**: Plugin-Name

    **Plugin-Package** plugin

    **Plugin-Version** 0.1

    **Plugin-Author** Virtuos86

    **Plugin-mail** http://dimonvideo.ru/0/name/Virtuos86

+++++++++++++++++++++++++++++++++++

#. App-Version-Min - минимальная версия приложения,
   которая поддерживается плагином;
#. App-Version-Max - максимальная версия приложения,
   которая поддерживается плагином;
#. Plugin-Version - версия плагина;
#. Plugin-Name - имя плагина для отображения в настройках приложения;
#. Plugin-Package - пакет, соответствующий плагину;
#. Plugin-Author - имя автора плагина;
#. Plugin-mail - контакты автора плагина;

"""

import os
import sys
import traceback
import ast

from . manifest import Manifest


def load_plugin(app, string_version):
    """Загружает, верифицирует и запускает плагины."""

    def check_plugins_path():
        if not os.path.exists(plugins_path):
            os.mkdir(plugins_path)
            open(os.path.join(plugins_path, '__init__.py'), 'w').write('')
            open(os.path.join(plugins_path, 'README.rst'), 'w').write(
                'This directory for plugins of Program\n'
                '--------------------------------------\n')
        if not os.path.exists(os.path.join(plugins_path, 'plugins_list.list')):
            open(
                os.path.join(plugins_path, 'plugins_list.list'), 'w').write(str(list()))

    plugins_path = \
        '{}/libs/plugins'.format(os.path.split(os.path.abspath(sys.argv[0]))[0])
    app.started_plugins = {}
    check_plugins_path()
    plugin_list = \
        ast.literal_eval(open('{}/plugins_list.list'.format(plugins_path)).read())

    for name in os.listdir(plugins_path):
        if name.startswith('__init__.'):
            continue

        path = os.path.join(plugins_path, name)

        if not os.path.isdir(path):
            continue
        try:
            manifest = Manifest(os.path.join(path, 'manifest.txt'))
            application_version = string_version.split()[0]

            app.started_plugins[name] = \
                {
                    'plugin-name': manifest['plugin-name'],
                    'plugin-desc': manifest['plugin-desc'],
                    'plugin-package': manifest['plugin-package'],
                    'plugin-version': manifest['plugin-version'],
                    'plugin-author': manifest['plugin-author'],
                    'plugin-mail': manifest['plugin-mail'],
                    'app-version-min': manifest['app-version-min'],
                    'app-version-max': manifest['app-version-max'],
                    'app-version': application_version
                }

            # Запускаем плагин, если он присутствует в списке активированых.
            if name in plugin_list:
                exec(open(os.path.join(path, '__init__.py'), encoding='utf-8').read(),
                     {'app': app, 'path': path})
        except Exception:
            raise Exception(traceback.format_exc())
