from os.path import basename, dirname, normpath


class Report:
    def __init__(self, report_folder):
        self.report_title = self.get_report_name(report_folder)

    @staticmethod
    def get_report_name(report_folder):
        return basename(dirname(normpath(report_folder)))