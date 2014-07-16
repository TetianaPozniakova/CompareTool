from os.path import join, exists, basename, dirname
from os import mkdir

from wand.image import Image
try:
    import comtypes.client
except ImportError:
    raise ImportError("Please make sure comtypes is installed.")


DEFAULT_RESOLUTION = 300
DEFAULT_IMAGE_HEIGHT = 768
DEFAULT_IMAGE_WEIGHT = 1366


def get_path_to_convert_folder(path):
    return path[::-1].replace(".", "-", 1)[::-1]


def create_folder(path_to_file):
    folder_path = get_path_to_convert_folder(path_to_file)
    if not (exists(folder_path)):
        mkdir(folder_path)


def get_filename_with_new_extension(path_to_file, file_type='png'):
    extension = path_to_file.split(".")[-1]
    file_path = path_to_file.replace(extension, file_type)
    return basename(file_path)


def convert_to_image(path_to_file, root_folder=False):
    with Image(filename=dirname(path_to_file.replace('-pdf', '.pdf')) if root_folder else path_to_file,
               resolution=DEFAULT_RESOLUTION) as img:
        img.resize(DEFAULT_IMAGE_WEIGHT, DEFAULT_IMAGE_HEIGHT)
        img.save(filename=path_to_file.replace('.pdf', '.png'))


def pdf_to_img(path_to_file):
    create_folder(path_to_file)
    print 'Converting PDF to png'
    convert_to_image(join(get_path_to_convert_folder(path_to_file),
                          get_filename_with_new_extension(path_to_file, 'pdf')), True)


def excel_to_img(path_to_file):
    create_folder(path_to_file)

    out_file = join(get_path_to_convert_folder(path_to_file), get_filename_with_new_extension(path_to_file, "pdf"))

    excel_app = comtypes.client.CreateObject('Excel.Application')
    excel_document = excel_app.Workbooks.Open(path_to_file)
    for sheet in excel_document.Sheets:
        sheet.PageSetup.Zoom = False
        sheet.PageSetup.FitToPagesWide = 1
        sheet.PageSetup.FitToPagesTall = False
    print 'Exporting Excel to PDF', out_file
    excel_document.ExportAsFixedFormat(0, out_file)
    excel_app.Quit()

    print 'Converting PDF to png'
    convert_to_image(out_file)


def ppt_to_img(path_to_file):
    create_folder(path_to_file)

    powerpoint_app = comtypes.client.CreateObject("Powerpoint.Application")
    powerpoint_app.Presentations.Open(path_to_file)
    print 'Exporting Power Point to png', path_to_file
    powerpoint_app.ActivePresentation.Export(join(get_path_to_convert_folder(path_to_file)), "png",
                                             ScaleWidth=DEFAULT_IMAGE_WEIGHT,
                                             ScaleHeight=DEFAULT_IMAGE_HEIGHT)
    powerpoint_app.Presentations[1].Close()
    powerpoint_app.Quit()


def doc_to_img(path_to_file):
    create_folder(path_to_file)

    out_file = join(get_path_to_convert_folder(path_to_file), get_filename_with_new_extension(path_to_file, "pdf"))

    word_app = comtypes.client.CreateObject('Word.Application')
    #word.Visible = True
    word_document = word_app.Documents.Open(path_to_file)
    print 'Exporting Word to PDF', out_file
    word_document.SaveAs(out_file, FileFormat=17)
    word_document.Close()
    word_app.Quit()

    print 'Converting PDF to png'
    convert_to_image(out_file)


def convert_to_images(dirname, filename):
    file_type = filename.split('.')[-1]
    path_to_file = join(dirname, filename)

    print "Convert {path_to_file} to image...".format(path_to_file=path_to_file)

    if file_type == "pdf":
        pdf_to_img(path_to_file)

    elif file_type == "xls":
        excel_to_img(path_to_file)

    elif file_type == "doc":
        doc_to_img(path_to_file)

    elif file_type == "ppt":
        ppt_to_img(path_to_file)
