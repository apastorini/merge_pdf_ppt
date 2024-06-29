import os
from PyPDF2 import PdfReader, PdfWriter
from pptx import Presentation
import comtypes.client


def convert_ppt_to_pdf(ppt_path, pdf_path):
    powerpoint = comtypes.client.CreateObject("Powerpoint.Application")
    powerpoint.Visible = 1
    presentation = powerpoint.Presentations.Open(ppt_path)
    presentation.SaveAs(pdf_path, 32)  # 32 es el formato para PDF
    presentation.Close()
    powerpoint.Quit()


def merge_pdfs(input_folder, output_file):
    pdf_writer = PdfWriter()

    for root, _, files in os.walk(input_folder):
        for file in files:
            file_path = os.path.join(root, file)
            if file.endswith('.pdf'):
                pdf_reader = PdfReader(file_path)
                for page in pdf_reader.pages:
                    pdf_writer.add_page(page)
            elif file.endswith('.ppt') or file.endswith('.pptx'):
                pdf_temp_path = file_path.rsplit('.', 1)[0] + '.pdf'
                convert_ppt_to_pdf(file_path, pdf_temp_path)
                pdf_reader = PdfReader(pdf_temp_path)
                for page in pdf_reader.pages:
                    pdf_writer.add_page(page)
                os.remove(pdf_temp_path)

    with open(output_file, 'wb') as output_pdf:
        pdf_writer.write(output_pdf)


if __name__ == "__main__":
    input_folder = r'C:\docpdf1'  # Usa r'' para rutas de Windows
    output_file = 'resultado.pdf'
    merge_pdfs(input_folder, output_file)
    print(f'Archivo PDF combinado creado: {output_file}')
