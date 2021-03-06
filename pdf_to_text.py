import sys
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.pdfpage import PDFPage
from pdfminer.converter import XMLConverter, HTMLConverter, TextConverter
from pdfminer.layout import LAParams
import io


def pdfparser(data):
    fp = open(data, 'rb')
    rsrcmgr = PDFResourceManager()
    retstr = io.StringIO()
    codec = 'utf-8'
    laparams = LAParams()
    device = TextConverter(rsrcmgr, retstr, laparams=laparams)
    # Create a PDF interpreter object.
    interpreter = PDFPageInterpreter(rsrcmgr, device)
    # Process each page contained in the document.

    for page in PDFPage.get_pages(fp):
        interpreter.process_page(page)
        data =  retstr.getvalue()

    print(data)


def pdf_to_word():
    # from https://stackoverflow.com/questions/52327434/is-there-any-way-to-convert-pdf-file-to-docx-using-python
    doc_folder = r"C:\Users\mavec\Desktop\word_to_text\\"
    import glob
    import win32com.client
    import os

    word = win32com.client.Dispatch("Word.Application")
    word.Visible = 0
    reqs_path = r"C:\Users\mavec\Desktop\word_to_text\generated_doc\\"

    for i, doc in enumerate(glob.iglob(doc_folder + "*.pdf")):
        print(doc)
        if "~$" in doc: pass
        filename = doc.split('\\')[-1]
        in_file = os.path.abspath(doc)
        print(in_file)
        wb = word.Documents.Open(in_file)
        out_file = os.path.abspath(reqs_path + filename[0:-4] + ".docx".format(i))
        print("outfile\n", out_file)
        wb.SaveAs2(out_file, FileFormat=16)  # file format for docx
        print("success...")
        wb.Close()

    word.Quit()



if __name__ == '__main__':
    pdfparser("20200424-sitrep-95-covid-19.pdf")


# Code credit to https://stackoverflow.com/a/21564675
# Test pdf from https://www.who.int/docs/default-source/coronaviruse/situation-reports/20200424-sitrep-95-covid-19.pdf?sfvrsn=e8065831_4