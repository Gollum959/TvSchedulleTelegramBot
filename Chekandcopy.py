import os
import pendulum
import re


list_dir = os.listdir(path="E:\Work\Python\Project\Test Directory")
today = pendulum.today()
current_week_start = today.start_of('week')
current_week_end = today.end_of('week')
next_week_start = current_week_start.add(weeks=1)
next_week_end = current_week_end.add(weeks=1)


def regular_string(week_start, week_end):
    regular_string = r'(0' + str(week_start.day) + '|' + str(week_start.day) + ')\
[.]*\
(0' + str(week_start.month) + '|' + str(week_start.month) + ')*\
[.]*\
(' + str(week_start.year) + ')*\
[\s]*-[\s]*\
(0' + str(week_end.day) + '|' + str(week_end.day) + ')\
[.]*\
(0' + str(week_end.month) + '|' + str(week_end.month) + ')*\
[.]*\
(' + str(week_end.year) + ')*'
    return regular_string


def path(week_start, week_end):
    regular_strig_week = regular_string(week_start, week_end)
    for directory in list_dir:
        match_dir = re.findall(regular_strig_week, directory)
        if match_dir:
            return f'E:\Work\Python\Project\Test Directory\{directory}'
    return ''


def file_path(week_start, week_end):
    path_file = path(week_start, week_end)
    if path_file:
        list_file = os.listdir(path=path_file)
        for file in list_file:
            match_file = re.findall(r"(p|р)(т|t)(s)\w*(.doc|.docx)$", file, re.IGNORECASE)
            if match_file:
                return f'{path_file}\{file}'
    return ''


def chek_and_copy(file_path, local_file, text):
    sourse_file_info = os.stat(file_path)
    sourse_file_size = sourse_file_info.st_size

    if os.path.isfile(local_file):
        local_file_info = os.stat(local_file)
        local_file_size = local_file_info.st_size
    else:
        local_file_size = 0

    if sourse_file_size == local_file_size:
        print(f'File {text} is already exist')
    else:
        shutil.copyfile(file_path, local_file, follow_symlinks=True)
        print(f'File {text} is copied')


def chek_file_exist(soure_file_path, local_file_path, text):
    if os.path.isfile(soure_file_path):
        chek_and_copy(soure_file_path, local_file_path, text)
    else:
        print(f'Sourse file {text} isn`t found')


this_weak_sourse_file_path = file_path(current_week_start, current_week_end)
next_weak_sourse_file_path = file_path(next_week_start, next_week_end)

if this_weak_sourse_file_path:
    chek_file_exist(this_weak_sourse_file_path, this_week_local_file_name, 'for this week')
else:
    print(f'Sourse file for this week isn`t found')

if next_weak_sourse_file_path:
    chek_file_exist(next_weak_sourse_file_path, next_week_local_file_name, 'for next week')
else:
    print(f'Sourse file for next week isn`t found')