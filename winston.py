
from segno import helpers 
from PyPDF2 import PdfReader
from openpyxl import load_workbook

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
for row in all_rows[1:]:
    rowinfo=[]
    i+=1
    for cell in all_rows[i]:
        rowinfo.append(cell.value)
    print(rowinfo)
    qrcode = helpers.make_vcard(rowinfo[0][rowinfo[0].find(" "):] + ";"+rowinfo[0].split()[0], displayname=rowinfo[0], org="NYCDDC: "+rowinfo[3], url="nyc.gov/ddc",
                              email=rowinfo[7], street = "30-30 Thompson Ave.", city="Long Island City", region="NY", zipcode = "11101",
                               workphone=rowinfo[5].replace("-",""), cellphone = rowinfo[6].replace("-",""), title=rowinfo[2])
    qrcode.save(rowinfo[0]+" QRCODE.png", scale=7)
print (type(rowinfo[5]))



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

