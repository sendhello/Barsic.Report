# Barsic.Report

## Описание:
Программа предназначена для формирования нестандартных отчетов с использованием базы данных ПО ДатаКрат "Барс 2". 
Данная версия предназначена непосредственно для ООО "АКВА". В других конфигурациях Барс 2 
данное ПО будет работать некорректно.

## Зависимости:
- Windows 10
- Python 3.7.x only
- git

## Установка:
1. `git clone https://github.com/sendhello/Barsic.Report.git`
2. Перейти в папку с программой и запустить `install.bat`


## Исправление конфликтов совместимости:

#### При ошибке отсутствия модуля sdl2 выполнить:
```commandline
venv\Scripts\pip.exe install --upgrade pip wheel setuptools
venv\Scripts\pip.exe install docutils pygments pypiwin32 kivy.deps.sdl2 kivy.deps.glew --extra-index-url https://kivy.org/downloads/packages/simple/
venv\Scripts\pip.exe uninstall -y kivy
venv\Scripts\pip.exe uninstall -y kivy.deps.sdl2
venv\Scripts\pip.exe uninstall -y kivy.deps.glew
venv\Scripts\pip.exe uninstall -y kivy.deps.gstreamer
venv\Scripts\pip.exe uninstall -y image
venv\Scripts\pip.exe install kivy
venv\Scripts\pip.exe freeze > requirements.txt
```
