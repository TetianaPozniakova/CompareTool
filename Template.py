from os.path import join
try:
    from jinja import FileSystemLoader
    from jinja.environment import Environment
except ImportError:
    raise ImportError("Please make sure Jinja is installed.")


DEFAULT_TEMPLATE = 'Full_Compare_Report_template.html'
DEFAULT_TEMPLATE_FOLDER = '.'


class TemplateVariables:
    def __init__(self):
        self.bat_list = []
        self.html_list = []
        self.old_report = ''
        self.new_report = ''


class Template:
    def __init__(self, template_vars, dirname, report_title):
        self.__template_vars = template_vars
        self.__template_name = "Full_Compare_Report_template_" + report_title + ".html"
        self.__template_path = self.__get_current_template_path__(dirname)

    def __get_current_template_path__(self, dirname):
        return join(dirname, self.__template_name)

    def create_template(self):
        env = Environment()
        env.loader = FileSystemLoader(DEFAULT_TEMPLATE_FOLDER)
        template = env.get_template(DEFAULT_TEMPLATE)
        output_text = template.render(self.__template_vars)
        output_report = open(self.__template_path, 'w')
        output_report.write(output_text)
        output_report.close()
