from docx import Document
from docx2pdf import convert
from datetime import datetime
from utilities import read_excel


def generate_pdf(input_docx_path, temp_docx_path, output_pdf_path, user_data, town_data):
    # Load the DOCX file
    doc = Document(input_docx_path)

    for paragraph in doc.paragraphs:
        for run in paragraph.runs:
            # Perform the replacement
            new_text = run.text
            new_text = new_text.replace("data", datetime.today().strftime('%d.%m.%Y'))
            new_text = new_text.replace("imie", user_data[0])
            new_text = new_text.replace("nazwisko", user_data[1])
            new_text = new_text.replace("miejscowosc", user_data[2])
            new_text = new_text.replace("czlonekczlonkini", user_data[3])
            new_text = new_text.replace("mail", user_data[4])
            new_text = new_text.replace("nazwagminy", town_data[0])

            if town_data[1].upper() == "P":
                new_text = new_text.replace("lposiada", "posiada")
                new_text = new_text.replace("lplanuje", "planuje")
            elif town_data[1].upper() == "M":
                new_text = new_text.replace("lposiada", "posiadają")
                new_text = new_text.replace("lplanuje", "planują")

            if town_data[2].upper() == "G":
                new_text = new_text.replace("miastogmina", "Gminy")
            elif town_data[2].upper() == "M":
                new_text = new_text.replace("miastogmina", "Miasta")

            if town_data[3].upper() == "W":
                new_text = new_text.replace("tytul", "Wójt")
            elif town_data[3].upper() == "B":
                new_text = new_text.replace("tytul", "Burmistrz")
            elif town_data[3].upper() == "P":
                new_text = new_text.replace("tytul", "Prezydent")

            if town_data[4].upper() == "M":
                new_text = new_text.replace("zwrot", "Szanowny Pan")
            elif town_data[4].upper() == "K":
                new_text = new_text.replace("zwrot", "Szanowna Pani")

            new_text = new_text.replace("w_in", town_data[5])

            # Update the text of the run
            run.text = new_text

            # Preserve formatting
            if run.bold:
                run.font.bold = True
            if run.italic:
                run.font.italic = True
            if run.underline:
                run.font.underline = True
            if run.font.size:
                run.font.size = run.font.size
            if run.font.color.rgb:
                run.font.color.rgb = run.font.color.rgb

    # Save the modified DOCX file to a temporary file
    doc.save(temp_docx_path)

    # Convert the temporary DOCX file to PDF
    convert(temp_docx_path, output_pdf_path)


def generate_docs():
    # Load the Excel workbook
    user_data, towns_data = read_excel()

    for town_data in towns_data:
        generate_pdf("debug/input/application.docx", f"debug/bin/{town_data[0]}.docx", f"debug/output/{town_data[0]}.pdf", user_data, town_data)


if __name__ == "__main__":
    generate_docs()
