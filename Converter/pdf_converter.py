import os
from Converter.convert_to_pdf import ConvertToPDF

FILE_FORMATS = {
    'doc': ['docx', 'doc', 'docm', 'rtf'],
    'ppt': ['ppt', 'pptx'],
    'xls': ['xlsx', 'xls', 'xlsm'],
    'txt': ['txt', 'xml', 'ini', 'log'],
    'image': ['jpg', 'jpeg', 'jpe', 'jp2', 'png', 'gif', 'tif', 'tiff', 'bmp', 'heic'],
    'html': ['html', 'htm']
}


def filter_files(folder_input):
    file_list = [os.path.join(folder_input, file) for file in os.listdir(folder_input)]

    format_list = [k for k in FILE_FORMATS]

    filtered_files = {k: [] for k in FILE_FORMATS}

    for file in file_list:
        for fr in format_list:
            frmt = os.path.splitext(file)[1][1:]
            if frmt in FILE_FORMATS[fr]:
                filtered_files[fr].append(file)

    return filtered_files


def convert_files(file_dict):

    for frmt in file_dict:
        for file in file_dict[frmt]:
            print('Converting:', file)
            if frmt == 'doc':
                ConvertToPDF.docx_to_pdf(file)
            elif frmt == 'ppt':
                ConvertToPDF.pptx_to_pdf(file)
            elif frmt == 'xls':
                ConvertToPDF.xls_to_pdf(file)
            elif frmt == 'image':
                ConvertToPDF.image_to_pdf(file)
            elif frmt == 'txt':
                ConvertToPDF.text_to_pdf(file)
            elif frmt == 'html':
                ConvertToPDF.html_to_pdf(file)
            else:
                continue
            print('Removing File:', file)
            # os.remove(file)
