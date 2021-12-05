import pendulum
import re
import docx
class ScheduleDocument:

    today = pendulum.today()
    day_of_week = 7 if today.day_of_week == 0 else today.day_of_week

    def __init__(self, path_to_file):
        self.path_to_file = path_to_file
        self.doc = docx.Document(self.path_to_file)

    def parse_document(self):
        self.map_pts_and_hardware = {}
        for index_and_table in enumerate(self.doc.tables):
            number_of_colums = len(index_and_table[1].rows[0].cells)
            for number in range(1, number_of_colums):
                name_equipment = index_and_table[1].rows[0].cells[number].text.strip()
                self.map_pts_and_hardware[name_equipment] = [index_and_table[0], number]

    def show_day_timetable_pts(self, name_equipment):
        tables = self.doc.tables
        table_num = self.map_pts_and_hardware[name_equipment][0]
        colum_num = self.map_pts_and_hardware[name_equipment][1]
        result = self.print_paragraps(tables[table_num].rows[ScheduleDocument.day_of_week].cells[0].paragraphs)
        result += '_______________\n'
        work_schedule = self.print_paragraps(tables[table_num].rows[ScheduleDocument.day_of_week].cells[colum_num].paragraphs)
        result += work_schedule if work_schedule.strip() else 'Работ нет'
        return result

    def remove_strike_text(self, paragraph):
        lst = paragraph.runs
        small = filter(lambda x: not x.font.strike, lst)
        result = ''
        for item in list(small):
            result += item.text
        result = result + '\n' if result else result
        return result

    def print_paragraps(self, text_in_cells):
        result = ''
        for date_cell_text in text_in_cells:
            if re.findall(r"\w", date_cell_text.text):
                result += self.remove_strike_text(date_cell_text) + ' '
        return result
