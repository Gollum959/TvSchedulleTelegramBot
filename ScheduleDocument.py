import docx
import ScheduleDocumentTables

class ScheduleDocument:

    def __init__(self, path_to_file):
        self.path_to_file = path_to_file
        self.doc = None
        self.tables = None

    def init(self):
        self.doc = docx.Document(self.path_to_file)
        self.tables = ScheduleDocumentTables.ScheduleDocumentTables(self.doc.tables)
        self.tables.init()

    def get_map(self):
        return self.tables.map_pts_and_hardware

    def show_day_timetable_pts(self, name_equipment):
        return self.tables.show_day_timetable_pts(name_equipment)

    def show_week_timetable_pts(self, name_equipment):
        return self.tables.show_week_timetable_pts(name_equipment)

