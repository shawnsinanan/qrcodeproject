# O(n) runtime
from segno import helpers 
from PyPDF2 import PdfReader #For reading pdf-may be useful later
from openpyxl import load_workbook
import xlwt #pip install xlsxwriter, xlrd, xlwt, xlutils
from xlwt import Workbook
import openpyxl 
from openpyxl import load_workbook
from openpyxl.drawing.image import Image

#xlrd, xlutils and xlwt modules need to be installed.  
#Can be done via pip install <module>
wrkb = openpyxl.Workbook()



#File should be read in through the program. Drag and drop in? Maybe user can select which excel file to input?
data_file = 'Business Card Order #42.xlsx'  

# Load the entire workbook.
wb = load_workbook(data_file, data_only=True)

# List all the sheets in the file.
# print("Found the following worksheets:")
# for sheetname in wb.sheetnames:
#     print(sheetname)

for sheetname in wb.sheetnames:
    ws = wb['Sheet1']
    all_rows = list(ws.rows)
# print(len(all_rows))
# for cell in all_rows[0]:
#     print(cell.value)
i=0
tempforfile=""
for row in all_rows[1:]: #every row
    rowinfo=[]
    i+=1
    for cell in all_rows[i]:  #every cell in row
        rowinfo.append(cell.value)
    print(rowinfo)
    qrcode = helpers.make_vcard(rowinfo[0][rowinfo[0].find(" "):] + ";"+rowinfo[0].split()[0], displayname=rowinfo[0], org="NYCDDC: "+rowinfo[3], url="nyc.gov/ddc",
                              email=rowinfo[7], street = "30-30 Thompson Ave.", city="Long Island City", region="NY", zipcode = "11101",
                               workphone=rowinfo[5].replace("-",""), cellphone = rowinfo[6].replace("-",""), title=rowinfo[2])          
    qrcode.save(rowinfo[0]+" QRCODE.png", scale=7)
    img = openpyxl.drawing.image.Image(rowinfo[0]+" QRCODE.png")
    img.anchor = "K"+str(i+1)
    img.height=50
    img.width=50
    ws.add_image(img)
    wb.save("BBusiness Card Order #42.xlsx")



# print (type(rowinfo[5]))


# reader = PdfReader("Business Card Request.pdf")
# fields = reader.get_form_text_fields()
# testfields = reader.getFields()
# print (fields)
# print(testfields)

# try:  # In the case EmployeeName field is empty or has one word
#     firstname = fields["Employee Name"].split()[0]
# except:
#     firstname=""

# try: 
#     lastname = fields["Employee Name"].split()[1]
# except: 
#     lastname=""


# qrcode = helpers.make_vcard(name= lastname+ ";"+firstname, displayname=fields["Employee Name"], org="NYCDDC-"+fields["DivisionUnit"], url="nyc.gov/ddc",
#                              email=fields["Email Address"], street = "30-30 Thompson Ave.", city="Long Island City", region="NY", zipcode = "11101",
#                               workphone=fields["Office Phone"].strip(" /-\\"), phone = fields["Telephone Number"].strip(" /-\\"), cellphone = fields["Cellphone Number optional"].strip(" /-\\"), title=fields["Office Title"])
# qrcode.save("TESTQRCODE.png", scale=7)



# f = open("C:\\Users\\tsuiwi\\Desktop\\raft.txt", "r")
# print(f.readlines())


