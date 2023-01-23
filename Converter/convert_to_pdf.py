import os


class ConvertToPDF:
    # @staticmethod
    # def docs_to_pdf(source):
    #     # we need to install first
    #     # pip install docx2pdf
    #     from docx2pdf import convert
    #
    #     try:
    #         target = os.path.splitext(source)[0] + '.pdf'
    #         convert(source, target)
    #         return True
    #
    #     except Exception as ex:
    #         raise ex

    @staticmethod
    def docx_to_pdf(source):
        import comtypes.client

        wdFormatPDF = 17
        out_file = os.path.splitext(source)[0]
        word = comtypes.client.CreateObject('Word.Application')
        doc = word.Documents.Open(source)

        try:
            doc.SaveAs(out_file, FileFormat=wdFormatPDF)
            return True

        except Exception as ex:
            raise ex
        finally:
            doc.Close()
            word.Quit()

    @staticmethod
    def image_to_pdf(source):
        # pip install img2pdf
        # pip install pillow_heif

        import img2pdf
        from PIL import Image
        try:
            # opening image
            if source.endswith('tiff') or source.endswith('tif'):
                return ConvertToPDF.__tiff_to_pdf(source)
            elif source.endswith('heic'):
                from pillow_heif import register_heif_opener
                register_heif_opener()

            image = Image.open(source)

            target = os.path.splitext(source)[0] + '.pdf'

            # converting into chunks using img2pdf
            pdf_bytes = img2pdf.convert(image.filename)

            # opening or creating pdf file
            file = open(target, "wb")

            # writing pdf files with chunks
            file.write(pdf_bytes)

            # closing image file
            image.close()
            file.close()
            return True

        except Exception as ex:
            raise ex

    @staticmethod
    def __tiff_to_pdf(tiff_path: str) -> str:
        from PIL import Image, ImageSequence

        if tiff_path.endswith('tiff'):
            pdf_path = tiff_path.replace('.tiff', '.pdf')
        else:
            pdf_path = tiff_path.replace('.tif', '.pdf')

        if not os.path.exists(tiff_path):
            raise Exception(f'{tiff_path} does not find.')
        image = Image.open(tiff_path)

        images = []
        for i, page in enumerate(ImageSequence.Iterator(image)):
            page = page.convert("RGB")
            images.append(page)
        if len(images) == 1:
            images[0].save(pdf_path)
        else:
            images[0].save(pdf_path, save_all=True, append_images=images[1:])
        return True

    @staticmethod
    def pptx_to_pdf(source):
        from win32com import client
        import os
        out_file = os.path.splitext(source)[0]
        powerpoint = client.Dispatch("Powerpoint.Application")
        pdf = powerpoint.Presentations.Open(source, WithWindow=False)

        try:
            pdf.Saveas(out_file, 32)
            return True

        except Exception as ex:
            raise ex
        finally:
            pdf.Close()
            powerpoint.Quit()

    @staticmethod
    def text_to_pdf(source):
        from fpdf import FPDF
        import textwrap
        try:
            target = os.path.splitext(source)[0] + '.pdf'
            file = open(source, encoding='UTF-8')
            text = file.read()
            file.close()

            a4_width_mm = 210
            pt_to_mm = 0.35
            fontsize_pt = 10
            fontsize_mm = fontsize_pt * pt_to_mm
            margin_bottom_mm = 10
            character_width_mm = 7 * pt_to_mm
            width_text = a4_width_mm / character_width_mm

            pdf = FPDF(orientation='P', unit='mm', format='A4')
            pdf.set_auto_page_break(True, margin=margin_bottom_mm)
            pdf.add_page()
            pdf.set_font(family='Courier', size=fontsize_pt)
            splitted = text.split('\n')

            for line in splitted:
                lines = textwrap.wrap(line, width_text)

                if len(lines) == 0:
                    pdf.ln()

                for wrap in lines:
                    pdf.cell(0, fontsize_mm, wrap, ln=1)

            pdf.output(target, 'F')
            return True

        except Exception as ex:
            raise ex

    @staticmethod
    def html_to_pdf(source):
        # pip install pdfkit
        #
        # install wkhtmltopdf
        # https://github.com/wkhtmltopdf/wkhtmltopdf/releases/download/0.12.4/wkhtmltox-0.12.4_msvc2015-win64.exe

        import pdfkit

        try:
            target = os.path.splitext(source)[0] + '.pdf'
            pdfkit.from_file(source, target)
            return True

        except Exception as ex:
            raise ex

    @staticmethod
    def xls_to_pdf(source):
        from win32com import client

        target = os.path.splitext(source)[0] + '.pdf'
        # Open Microsoft Excel
        excel = client.Dispatch("Excel.Application")

        # Read Excel File
        sheets = excel.Workbooks.Open(source)

        try:
            ws_index_list = [i + 1 for i in range(len(sheets.WorkSheets))]
            sheets.WorkSheets(ws_index_list).Select()

            # Convert into PDF File
            sheets.ActiveSheet.ExportAsFixedFormat(0, target)
            return True

        except Exception as ex:
            raise ex
        finally:
            sheets.Close(False)
            excel.Quit()
