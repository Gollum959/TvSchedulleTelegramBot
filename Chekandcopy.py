import os
import pendulum
import re
import shutil
import win32com.client as win32
from win32com.client import constants

today = pendulum.today()

list_dir = os.listdir(path="\\Project\\Test Directory")
work_dir = "\\Project\\Test Right Directory\\"
this_week_local_file_name = "thisweek"
next_week_local_file_name = "nextweek"

current_week_start = today.start_of('week')
current_week_end = today.end_of('week')
next_week_start = current_week_start.add(weeks=1)
next_week_end = current_week_end.add(weeks=1)


def regular_string(week_start, week_end):
    """Returns regular expression for searching a dir."""
    return (
        r'(0' + str(week_start.day) + '|' +
        str(week_start.day) + ')[.]*(0' + str(week_start.month) +
        '|' + str(week_start.month) + ')*[.]*(' +
        str(week_start.year) + ')*[\s]*-[\s]*(0' +
        str(week_end.day) + '|' + str(week_end.day) +
        ')[.]*(0' + str(week_end.month) + '|' +
        str(week_end.month) + ')*[.]*(' + str(week_end.year) + ')*'
    )


def path(week_start, week_end):
    """Returns a dir with shcedule."""
    regular_strig_week = regular_string(week_start, week_end)
    for directory in list_dir:
        if re.findall(regular_strig_week, directory):
            return f'\\Project\\Test Directory\\{directory}'
    return ''


def file_path(week_start, week_end):
    """returns a file matching the conditions of the regular expression."""
    path_file = path(week_start, week_end)
    if path_file:
        list_file = os.listdir(path=path_file)
        for file in list_file:
            if re.findall(
                r"(p|р)(т|t)(s)[\w \W]*(.doc|.docx|rtf)$", file, re.IGNORECASE
            ):
                return f'{path_file}\\{file}'
    return ''


def chek_and_copy(path_to_file, local_file, text):
    """Checks if the file in the source directory has changed,
    copy and convert if needed to docx."""
    extension = '.docx' if path_to_file.endswith('.docx') else (
        '.doc' if path_to_file.endswith('.doc') else '')
    source_file_info = os.stat(path_to_file)
    source_file_size = source_file_info.st_size
    local_file_path = work_dir + local_file + extension

    if os.path.isfile(local_file_path):
        local_file_info = os.stat(local_file_path)
        local_file_size = local_file_info.st_size
    else:
        local_file_size = 0

    if source_file_size == local_file_size:
        print(f'File {text} is already exist')
    else:
        shutil.copyfile(path_to_file, local_file_path, follow_symlinks=True)
        print(f'File {text} is copied')
        if extension == '.doc':
            save_as_docx(local_file_path, local_file)
            print(f'File {text} is converted')


def save_as_docx(path, file_name):
    """Convert doc to docx."""
    word = win32.gencache.EnsureDispatch('Word.Application')
    doc = word.Documents.Open(path)
    doc.Activate()
    new_file_abs = work_dir+file_name+'.docx'
    word.ActiveDocument.SaveAs(
        new_file_abs, FileFormat=constants.wdFormatXMLDocument
    )
    doc.Close(False)


def chek_file_exist(source_file_path, local_file_name, text):
    """Сhecks if file exists."""
    if os.path.isfile(source_file_path):
        chek_and_copy(source_file_path, local_file_name, text)
    else:
        print(f'Source file {text} isn`t found')


this_weak_source_file_path = file_path(current_week_start, current_week_end)
next_weak_source_file_path = file_path(next_week_start, next_week_end)

if this_weak_source_file_path:
    chek_file_exist(
        this_weak_source_file_path,
        this_week_local_file_name,
        'for this week'
    )
else:
    print('Source file for this week isn`t found')

if next_weak_source_file_path:
    chek_file_exist(
        next_weak_source_file_path,
        next_week_local_file_name,
        'for next week')
else:
    print('Source file for next week isn`t found')
