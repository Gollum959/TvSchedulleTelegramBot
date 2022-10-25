import re
import pendulum


class ScheduleDocumentTables:
    """Class works with a special table with a work schedule."""

    today = pendulum.today()
    day_of_week = 7 if today.day_of_week == 0 else today.day_of_week

    def __init__(self, tables):
        self.tables = tables
        self.map_pts_and_hardware = {}

    def init(self):
        for index_and_table in enumerate(self.tables):
            numbers_of_colums = len(index_and_table[1].rows[0].cells)
            for number in range(1, numbers_of_colums):
                name_equipment = index_and_table[1].rows[0].cells[number].text.strip()
                self.map_pts_and_hardware[name_equipment] = [index_and_table[0], number]

    def get_table_num(self, name_equipment):
        return self.map_pts_and_hardware[name_equipment][0]

    def get_column_num(self, name_equipment):
        return self.map_pts_and_hardware[name_equipment][1]

    def show_day_timetable_pts(self, name_equipment):
        result = self.__print_paragraphs(self
                                         .tables[self.get_table_num(name_equipment)]
                                         .rows[ScheduleDocumentTables.day_of_week]
                                         .cells[0]
                                         .paragraphs)
        result += '_______________\n'
        work_schedule = self.__print_paragraphs(self
                                                .tables[self.get_table_num(name_equipment)]
                                                .rows[ScheduleDocumentTables.day_of_week]
                                                .cells[self.get_column_num(name_equipment)]
                                                .paragraphs)
        result += work_schedule if work_schedule.strip() else 'Работ нет'
        return result

    def show_week_timetable_pts(self, name_equipment):
        result = ''
        for i in range(1, 8):
            result += self.__print_paragraphs(self
                                              .tables[self.get_table_num(name_equipment)]
                                              .rows[i]
                                              .cells[0]
                                              .paragraphs)
            result += ('__________\n')
            work_schedule = self.__print_paragraphs(self
                                                    .tables[self.get_table_num(name_equipment)]
                                                    .rows[i]
                                                    .cells[self.get_column_num(name_equipment)]
                                                    .paragraphs)
            if work_schedule.strip().startswith('Изм.'):
                result += 'Работ нет\n'
            elif work_schedule.strip():
                result += work_schedule
            else:
                result += 'Работ нет\n'
            result += ('__________\n')
        return result

    @staticmethod
    def __remove_strike_text(paragraph):
        lst = paragraph.runs
        not_strike = filter(lambda x: not x.font.strike, lst)
        result = ''
        for item in list(not_strike):
            result += item.text
        result = result + '\n' if result else result
        return result

    @staticmethod
    def __print_paragraphs(text_in_cells):
        result = ''
        for date_cell_text in text_in_cells:
            if re.findall(r"\w", date_cell_text.text):
                result += ScheduleDocumentTables.__remove_strike_text(date_cell_text) + ' '
        return result
