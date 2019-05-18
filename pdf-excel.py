import PyPDF2

file = open('china.pdf', 'rb')

filereader = PyPDF2.PdfFileReader(file)

print(filereader.numPages)
