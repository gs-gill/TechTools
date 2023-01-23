from Converter.pdf_converter import *


if __name__ == '__main__':
    input_path = r'C:\Users\shahbaz.ansari\Downloads\sample_input\New folder\test'

    filtered_file_dict = filter_files(input_path)
    convert_files(filtered_file_dict)
