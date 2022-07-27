# Снятие паролей с Excel файлов

## Описание:
Скрипт снимает защиту с листов и VBA скриптов.

## Развертывание:
- Склонируйте проект на Ваш компьютер 
```sh 
git clone https://github.com/DenisShahbazyan/unlocking_excel_sheets.git
``` 
- Перейдите в папку с проектом 
```sh 
cd unlocking_excel_sheets
``` 
- Создайте и активируйте виртуальное окружение 
```sh 
python -m venv venv 
source venv/Scripts/activate 
``` 
- Обновите менеджер пакетов (pip) 
```sh 
pip install --upgrade pip 
``` 
- Установите необходимые зависимости 
```sh 
pip install -r requirements.txt
``` 
- Создание исполняемого файла
```sh
pyinstaller --onefile main.py
```
- После выполнения команды, в директории с проектом появиться папка `dist` в ней будет лежать исполняемый файл `main.exe`. 
- Положите свой `excel` файл рядом с исполняемым файлом и перетащите `excel` файл на `main.exe`. После чего защита будет снята.

P.S.: В папке `exe` лежит исполняемый файл.

## Системные требования:
- [Python](https://www.python.org/) 3.10.4

## Планы по доработке:
>Если появятся баги будем исправлять 😀 Отрефакторить код, так как много повторений.

## Используемые технологии:
- [pyinstaller](https://pypi.org/project/pyinstaller/) 5.2

## Авторы:
- [Denis Shahbazyan](https://github.com/DenisShahbazyan)

## Лицензия:
- MIT
