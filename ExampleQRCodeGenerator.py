from segno import helpers
import pandas as pd 

data = {'Names': ["Bob Ross", "Joe Moe", "Anna Jhon", "Peter Parker"],
'Email' : ["bobross@gmail.com" , "joemoe@gmail.com" , "annajhon@gmail.com" , "peterparker@gmail.com"],
'Phone' : ["3333333333" , "4444444444" , "5555555555" , "6666666666"]
}

data_pandas = pd.DataFrame(data)

for i in range(4):
    qrcode = helpers.make_vcard(name=data_pandas.at[i, 'Names'], displayname=data_pandas.at[i, 'Names'], email=[data_pandas.at[i, 'Email']], phone=[data_pandas.at[i, 'Phone']])
    qrcode.save(f"{data_pandas.at[i, 'Names']}.png", scale=10)








