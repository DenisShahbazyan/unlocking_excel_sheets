import os
import re
import shutil
import sys
import tempfile
import zipfile
from pathlib import Path

SHEETS_DIR = 'xl/worksheets/'
STARTSWITH_SHEET = 'sheet'
ENDSWITH_SHEET = '.xml'
TMP_FOLDER = 'tmp/'
BIN_FILENAME = 'xl/vbaProject.bin'
WORKBOOK = 'xl/workbook.xml'


def sheet_file_modification(filename, EXTRACT_DIR):
    """Достает все листы таблицы, удаляет с листов защиту и сохраняет на диск.
    """
    with zipfile.ZipFile(filename) as z:
        for name in z.namelist():
            info = z.getinfo(name)
            if (
                info.filename.startswith(SHEETS_DIR + STARTSWITH_SHEET)
                and info.filename.endswith(ENDSWITH_SHEET)
            ):
                z.extract(info.filename, EXTRACT_DIR)

    for sheet in os.listdir(EXTRACT_DIR / SHEETS_DIR):
        with open(
            EXTRACT_DIR / SHEETS_DIR / sheet, 'r', encoding='utf-8'
        ) as file:
            data_sheet = file.read()
            data_sheet = re.sub(r'(<sheetProtection[^<]+)', r'', data_sheet)

        with open(
            EXTRACT_DIR / SHEETS_DIR / sheet, 'w', encoding='utf-8'
        ) as file:
            file.write(data_sheet)


def workbook_file_modification(filename, EXTRACT_DIR):
    """Достает книгу, удаляет с книги защиту и сохраняет на диск.
    """
    with zipfile.ZipFile(filename) as z:
        for name in z.namelist():
            info = z.getinfo(name)
            if info.filename == WORKBOOK:
                z.extract(info.filename, EXTRACT_DIR)

    if os.path.exists(EXTRACT_DIR / WORKBOOK):
        with open(EXTRACT_DIR / WORKBOOK, 'r') as file:
            data_book = file.read()
            data_book = re.sub(r'(<workbookProtection[^<]+)', r'', data_book)

        with open(EXTRACT_DIR / WORKBOOK, 'w', encoding='utf-8') as file:
            file.write(data_book)


def vbaProjectbin_file_modification(filename, EXTRACT_DIR):
    """Достает бинарник макросов, удаляет с него пароль и сохраняет на диск.
    """
    with zipfile.ZipFile(filename) as z:
        for name in z.namelist():
            info = z.getinfo(name)
            if info.filename == BIN_FILENAME:
                z.extract(info.filename, EXTRACT_DIR)

    if os.path.exists(EXTRACT_DIR / BIN_FILENAME):
        with open(EXTRACT_DIR / BIN_FILENAME, 'rb+') as file:
            data_bin = file.read()
            index = data_bin.find(b'DPB="')
            if index != -1:
                file.seek(index)
                file.write(b'DPc="')


def get_all_path():
    """Возвращает валидные пути для изменных файлов, которые лежат на диске.
    """
    paths = sorted(Path(TMP_FOLDER).glob('**/*.*'))
    paths = list(map(str, paths))

    correct_path = []
    for path in paths:
        correct_path.append(path.replace('tmp\\', '').replace('\\', '/'))

    return correct_path


def update_excel_file(zipname):
    """Добавляет в Excel файл, файлы без защиты.
    """
    filenames = get_all_path()

    tmpfd, tmpname = tempfile.mkstemp(dir=os.path.dirname(zipname))
    os.close(tmpfd)

    with zipfile.ZipFile(zipname, 'r') as zin:
        with zipfile.ZipFile(tmpname, 'w') as zout:
            zout.comment = zin.comment
            for item in zin.infolist():
                if item.filename not in filenames:
                    zout.writestr(item, zin.read(item.filename))

    os.remove(zipname)
    os.rename(tmpname, zipname)

    with zipfile.ZipFile(
        zipname, mode='a', compression=zipfile.ZIP_DEFLATED
    ) as z:
        for filename in filenames:
            z.write(TMP_FOLDER + filename, filename)

    shutil.rmtree(TMP_FOLDER)


def main():
    """Точка входа в программу.
    """
    filename = sys.argv[-1]

    if os.path.exists(filename):
        EXTRACT_DIR = Path(filename).parent / TMP_FOLDER

        sheet_file_modification(filename, EXTRACT_DIR)
        # workbook_file_modification(filename, EXTRACT_DIR)
        vbaProjectbin_file_modification(filename, EXTRACT_DIR)

        update_excel_file(filename)


if __name__ == '__main__':
    main()
