import tabula
import PyPDF2 as pdf
import urllib.request

def pdf_page_count(self, filename):
    # Returns page count of filename/filepath
    pdfFileObj = open(filename, 'rb')
    pdfReader = pdf.PdfFileReader(pdfFileObj)
    length = pdfReader.numPages
    return length

length = pdf_page_count(self.pdf_path)
new_length = length + 1

for x in tqdm(range(3, new_length)):
    df = tabula.read_pdf(self.pdf_path, pandas_options={'header': 2}, pages = x, multiple_tables=True, guess = False, stream=True,lattice=False)

    assert len(df) == 1, 'Tabula read more than one table on Page {}'.format(x)
    df = df[0] #Tabula returns list 
