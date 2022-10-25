import docx

from ScheduleDocumentTables import ScheduleDocumentTables


class ScheduleDocument:         
    """Creates answers from docx table."""

    def __init__(self, path_to_file):
        self.path_to_file = path_to_file
        self.__doc = None
        self.__tables = None

    def init(self):
        self.__doc = docx.Document(self.path_to_file)
        self.__tables = ScheduleDocumentTables(self.__doc.tables)
        self.__tables.init()

    def get_map(self):
        return self.__tables.map_pts_and_hardware

    def get_day_timetable_pts(self, name_equipment):
        return self.__tables.show_day_timetable_pts(name_equipment)

    def get_week_timetable_pts(self, name_equipment):
        return self.__tables.show_week_timetable_pts(name_equipment)
