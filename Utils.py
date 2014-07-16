from os import rename
from os.path import join, exists


BC_IMAGES = "BcImages"


def split_by_extension(path):
    full_name, extension = path.rsplit(".", 1)
    return full_name, extension


def bcomp_clean(reportdir, compare_report):
    full_name, extension = split_by_extension(compare_report)
    if exists(join(reportdir, BC_IMAGES)):
        rename(join(reportdir, BC_IMAGES), full_name)
        f1 = open(compare_report, 'r').read()
        f2 = open(compare_report, 'w')
        m = f1.replace(BC_IMAGES, full_name.rsplit('\\', 1)[-1])
        f2.write(m)


def create_batch(report_path, etalon_report):
    #print 'Create batch file.'
    full_name, extension = split_by_extension(report_path)
    bat_name = full_name + '-' + extension + '.bat'
    bat_file = open(bat_name, 'w+')
    bat_file.write('BComp "{etalon_report}" "{report_path}"'.format(etalon_report=etalon_report,
                                                                    report_path=report_path))
    bat_file.close()
    return bat_name