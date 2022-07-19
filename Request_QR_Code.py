
#Generating Business Card with QR code as per request

#importing the modules
from segno import helpers
import pandas as pd 

full_name = input("Enter Employee name: ")
#id_number = input("Enter Employee_ID_Number: ")
#office_title = input("Enter Office Title: ")
office_phone = input("Enter  office phone: ")
#division_unit = input("Enter division unit: ")
#supervisor_name = input("Enter supervisor name ")
email_address = input("Enter email address: ")

data = {full_name, office_phone, email_address}

for i in range(4):
    qrcode = helpers.make_vcard(name= full_name, displayname= full_name, email= email_address, phone= office_phone)

    qrcode.save(f"{full_name}.png", scale=10)
