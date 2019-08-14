import PyPDF2

file = open('FILE.pdf', 'rb')
filereader = PyPDF2.PdfFileReader(file)
print(filereader.numPages)
