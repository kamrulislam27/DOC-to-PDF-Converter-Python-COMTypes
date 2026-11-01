import os
import comtypes.client

def doc_to_pdf(source_folder, target_folder):
    if not os.path.exists(target_folder):
        os.makedirs(target_folder)

    word = comtypes.client.CreateObject("Word.Application")
    word.Visible = 1

    for file in os.listdir(source_folder):
        if file.endswith((".doc", ".docx")):
            doc_file = os.path.join(source_folder, file)
            pdf_file = os.path.join(target_folder, os.path.splitext(file)[0] + ".pdf")

            document = word.Documents.Open(doc_file)
            document.SaveAs(pdf_file, 17)  # 17 = PDF format
            document.Close()

    word.Quit()


# Example usage:
source = r"F:\Backup Disk (D)\AHUT\Study Metarials\5th Semester\Computer Networks\DOCS"
target = r"F:\Backup Disk (D)\AHUT\Study Metarials\5th Semester\Computer Networks\DOCS"
doc_to_pdf(source, target)
