from os import walk
from os.path import isfile, join
import subprocess

from Convertors import convert_to_images
from Utils import split_by_extension, create_batch, bcomp_clean
from Template import TemplateVariables, Template
from Report import Report


DEFAULT_COMPARE_SCRIPT_FOLDER = "C:\\CompareScript\\"
TEXT_COMPARE_SETTINGS = "compare.script"
PICTURE_COMPARE_SETTINGS = "picture-compare.script"


class FileFormat:
    PDF = ('.pdf',)
    EXCEL = ('.xls', '.xlsx')
    WORD = ('.doc', '.docx')
    POWERPOINT = ('.ppt', '.pptx')
    PNG = ('.png', '.PNG')


class Comparator:
    def __init__(self, report_folder_path, compare_script_folder=DEFAULT_COMPARE_SCRIPT_FOLDER):
        self.__text_compare_settings = self.__build_path__(compare_script_folder, TEXT_COMPARE_SETTINGS)
        self.__picture_compare_settings = self.__build_path__(compare_script_folder, PICTURE_COMPARE_SETTINGS)
        self.__report_folder_path = report_folder_path
        self.compare_report = None

    def compare(self):
        for dirname, dirnames, filenames in walk(self.__report_folder_path):
            report_under_test = Report(dirname)
            pdf_results = TemplateVariables()
            ppt_results = TemplateVariables()
            xls_results = TemplateVariables()
            doc_results = TemplateVariables()
            for filename in filenames:
                convert_to_images(dirname, filename)

            if '\\2010' in dirname:
                for reportdir, reportdirs, reportfiles in walk(dirname):
                    for report in reportfiles:
                        report_path = join(reportdir, report)
                        etalon_report = report_path.replace('\\2010\\', '\\2005\\')
                        if report.endswith(FileFormat.PDF + FileFormat.POWERPOINT + FileFormat.WORD):
                            self.__run_bcomp__(report_path, etalon_report, self.__text_compare_settings)
                            bat_name = create_batch(report_path, etalon_report)
                            if Report.get_report_name(reportdir) == report_under_test.report_title and \
                                    report.endswith(FileFormat.PDF):
                                pdf_results.old_report = etalon_report
                                pdf_results.new_report = report_path
                                pdf_results.bat_list.append(bat_name)
                                pdf_results.html_list.append(self.compare_report)

                            elif report.endswith(".ppt"):
                                ppt_results.old_report = etalon_report
                                ppt_results.new_report = report_path
                                ppt_results.bat_list.append(bat_name)
                                ppt_results.html_list.append(self.compare_report)

                            elif report.endswith(".doc"):
                                doc_results.old_report = etalon_report
                                doc_results.new_report = report_path
                                doc_results.bat_list.append(bat_name)
                                doc_results.html_list.append(self.compare_report)


                        elif report.endswith(FileFormat.EXCEL):
                            xls_results.old_report = etalon_report
                            xls_results.new_report = report_path
                        elif report.endswith(FileFormat.PNG):
                            self.__run_bcomp__(report_path, etalon_report, self.__picture_compare_settings)
                            bcomp_clean(reportdir, self.compare_report)
                            bat_name = create_batch(report_path, etalon_report)
                            if reportdir.endswith("-xls"):
                                xls_results.bat_list.append(bat_name)
                                xls_results.html_list.append(self.compare_report)

                        if reportdir.endswith("-pdf"):
                            pdf_results.bat_list.append(bat_name)
                            pdf_results.html_list.append(self.compare_report)
                        elif reportdir.endswith("-ppt"):
                            ppt_results.bat_list.append(bat_name)
                            ppt_results.html_list.append(self.compare_report)
                        elif reportdir.endswith("-doc"):
                            doc_results.bat_list.append(bat_name)
                            doc_results.html_list.append(self.compare_report)

            template_vars = {
                "report_title": report_under_test.report_title,
                "pdf_old": pdf_results.old_report,
                "pdf_new": pdf_results.new_report,
                "pdf_bats": pdf_results.bat_list,
                "pdf_htmls": pdf_results.html_list,

                "excel_old": xls_results.old_report,
                "excel_new": xls_results.new_report,
                "xls_bats": xls_results.bat_list,
                "xls_htmls": xls_results.html_list,

                "ppt_old": ppt_results.old_report,
                "ppt_new": ppt_results.new_report,
                "ppt_bats": ppt_results.bat_list,
                "ppt_htmls": ppt_results.html_list,

                "doc_old": doc_results.old_report,
                "doc_new": doc_results.new_report,
                "doc_bats": doc_results.bat_list,
                "doc_htmls": doc_results.html_list,
            }
            Template(template_vars, dirname, report_under_test.report_title).create_template()

    def __run_bcomp__(self, report_path, etalon_report, compare_settings):
        full_name, extension = split_by_extension(report_path)
        self.compare_report = '{path}-{file_format}--CompareReport.html'.format(path=full_name, file_format=extension)
        #print 'Create compare report.'
        try:
            process = subprocess.Popen(['BComp',
                                        "@" + compare_settings,
                                        etalon_report,
                                        report_path,
                                        self.compare_report,
                                        "/silent",
                                        "/closescript"])
        except WindowsError:
            raise WindowsError('Please make sure Beyond Compare is installed.')
        process.wait()

    @staticmethod
    def __build_path__(compare_script_folder, compare_settings):
        path = join(compare_script_folder, compare_settings)
        if isfile(path):
            return path
        else:
            raise IOError("File {filename} does not exist.".format(filename=path))