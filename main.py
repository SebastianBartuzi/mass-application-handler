from docx import Document
from docx2pdf import convert
from datetime import datetime
from utilities import read_excel, replace_townname

DOCUMENT_TEMPLATE = "debug/input/application.docx"
BIN_PATH = "debug/bin/{name}.docx"
OUTPUT_PATH = "debug/output/{name}.pdf"


def generate_pdf(user_data, i, towns_data):
    # Load the DOCX file
    doc = Document(DOCUMENT_TEMPLATE)
    town_name = towns_data[i][0]
    towns_data[i][0] = replace_townname(towns_data, i)
    town_data = towns_data[i]

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
            new_text = new_text.replace("nazwagminy", town_name)

            if town_data[1].upper() == "P":
                new_text = new_text.replace("lposiada", "posiada")
                new_text = new_text.replace("lplanuje", "planuje")
            elif town_data[1].upper() == "M":
                new_text = new_text.replace("lposiada", "posiadają")
                new_text = new_text.replace("lplanuje", "planują")

            if town_data[2].upper() == "W":
                new_text = new_text.replace("tytul", "Wójt")
                new_text = new_text.replace("miastogmina", "Gminy")
            elif town_data[2].upper() == "B":
                new_text = new_text.replace("tytul", "Burmistrz")
                new_text = new_text.replace("miastogmina", "Miasta i Gminy")
            elif town_data[2].upper() == "P":
                new_text = new_text.replace("tytul", "Prezydent")
                new_text = new_text.replace("miastogmina", "Miasta")

            if town_data[3].upper() == "M":
                new_text = new_text.replace("zwrot", "Szanowny Pan")
            elif town_data[3].upper() == "K":
                new_text = new_text.replace("zwrot", "Szanowna Pani")

            new_text = new_text.replace("w_in", town_data[4])

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
    print(town_data[0])
    bin_path = BIN_PATH.replace("{name}", town_data[0])
    doc.save(bin_path)

    # Convert the temporary DOCX file to PDF
    output_path = OUTPUT_PATH.replace("{name}", town_data[0])
    convert(bin_path, output_path)


def generate_docs():
    # Load the Excel workbook
    user_data, towns_data = read_excel()

    for i in range(len(towns_data)):
        generate_pdf(user_data, i, towns_data)


if __name__ == "__main__":
    generate_docs()
