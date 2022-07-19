
from segno import helpers
import PyPDF2
from PyPDF2 import PdfReader

reader = PdfReader("Business Card Request.pdf")
fields = reader.get_form_text_fields()
print (fields)


# creating a pdf file object 
pdfFileObj = open('Business Card Request.pdf', 'rb') 
# creating a pdf reader object 
pdfReader = PyPDF2.PdfFileReader(pdfFileObj) 
# creating a page object 
pageObj = pdfReader.getPage(0) 
# closing the pdf file object 
pdfFileObj.close()


qrcode = helpers.make_vcard(name= fields["Employee Name"].split()[1]+ ";" + fields["Employee Name"].split()[0], displayname=fields["Employee Name"], org="NYCDDC", url="nyc.gov/ddc",
                             email=fields["Email Address"], street = "30-30 Thompson Ave.", city="Long Island City", region="NY", zipcode = "11101",
                             workphone=fields["Office Phone"], phone = fields["Telephone Number"].strip(" /-\\"), cellphone = fields["Cellphone Number optional"].strip(" /-\\"), title=fields["Office Title"])


qrcode.save('TESTQRCODE.png', scale=7)




# f = open("C:\\Users\\tsuiwi\\Desktop\\raft.txt", "r")
# print(f.readlines())

