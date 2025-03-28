import os.path
from zebra import Zebra
import pandas as pd
import os.path
import os

# This script is meant to read an excel inventory file and clean the data to show errors based on the conditions given.
# Creating dataframes for each condition and then compile the dataframes to 1.
# Finally, the Zebra library is meant to print off Thermal Labels that will have a "Reason for Return" per Product.

Username = os.path.join(os.path.expandvars("%userprofile%"))

col = ['ProductID', 'CategoryName', 'FUNCTIONALITY', 'ProductDescription', 'ProgramName']
File = 'Upload.xlsx'
for root, dirs, files in os.walk(Username):
    if File in files:
        path = ''.join(root + '\\' + File)

df = pd.read_excel(path, usecols=col)

# This is where we create a data frame from the columns we specified at the beginning.
# Creating 3 different frames that have specific columns.
df['ProductID'] = 'Product-' + df['ProductID'].astype(str)
df.fillna('N/A', inplace=True)
df1 = df[['ProductID', 'CategoryName', 'ProgramName']]
df5 = df[['ProductID', 'CategoryName', 'Classification_TECHNICAL FUNCTIONALITY']]
df7 = df[['ProductID', 'CategoryName', 'ProductDescription']]

# Our first frame we clean, is the functionality of Product1. Finding any product that is not Fully Functional.
lrslt_tech = df5.loc[(df5['CategoryName'] == 'Category -> Subcategory -> Product1') &
                  (df5['FUNCTIONALITY'] != 'FULLY FUNCTIONAL')]
lrslt_tech = lrslt_tech.rename(columns={'FUNCTIONALITY': 'Condition Pull'})

# Our second frame we clean, is the functionality of Product2. Finding any product that is not Fully Functional.
drslt_tech = df5.loc[(df5['CategoryName'] == 'Category -> Subcategory -> Product2') &
                  (df5['FUNCTIONALITY'] != 'FULLY FUNCTIONAL')]
drslt_tech = drslt_tech.rename(columns={'FUNCTIONALITY': 'Condition Pull'})
rslt_cto = df7.loc[df7['ProductDescription'].str.contains('CTO', case=False)]
rslt_cto = rslt_cto.rename(columns={'ProductDescription': 'Condition Pull'})

# Now we are compiling the frames to create an excel export of all the errors within inventory.
pdList = [lrslt_tech, drslt_tech, rslt_cto]  # List of your dataframes
new_df = pd.concat(pdList)
new_df.loc[new_df['Condition Pull'].str.contains('CTO'), 'Condition Pull'] = 'CTO'
new_df = new_df.dropna()
new_df = new_df.rename(columns={'CategoryName': 'ProgramName'})

# 2 files are created, picklist_output is what we use to find the product in the inventory.
# ZPrint-Errors is what the Zebra library uses to print off the thermal labels with the Reason for Return.
new_df.to_excel(Username + f'\Downloads\picklist_output.xlsx', index=False)
new_df.to_excel(Username + f'\Downloads\ZPrint-Errors.xlsx', index=False)


# This function is where we print the product within the ZPrint-Errors.
# Manually built the Thermal Label layout.
def Product():

    columns = ['ProductID', 'Condition Pull', 'ProgramName']

    username = os.path.join(os.path.expandvars("%userprofile%"))
    File = 'ZPrint-Errors.xlsx'
    for root, dirs, files in os.walk(username):
        if File in files:
            path = ''.join(root + '\\' + File)
    df = pd.read_excel(path, usecols=columns, na_values=' ')

    for index, row in df.iterrows():
        data = row["ProductID"], row["Condition Pull"], row["ProgramName"]
        pid, cond, prog = data
        label1 = ("CT~~CD,~CC^~CT~"
                  "^XA~TA000~JSN^LT0^MNW^MTD^PON^PMN^LH0,0^JMA^PR5,5~SD15^JUS^LRN^CI0^XZ"
                  "^XA"
                  "^MMT"
                  "^PW609"
                  "^LL0406"
                  "^LS0"
                  "^FT581,343^A0I,59,60^FH\^FDInventory Control Error^FS"
                  "^FO26,37^GB555,0,3^FS"
                  "^FO29,206^GB555,0,3^FS"
                  "^FO27,328^GB556,0,3^FS"
                  "^FT583,43^A0I,28,28^FH\^FDProductID:^FS"
                  "^FT143,42^A0I,28,28^FH\^FD10/17/2022^FS"
                  "^FT417,43^A0I,28,28^FH\^FDProduct-12345678^FS"
                  "^FT418,148^A0I,28,28^FH\^FDProgramName^FS"
                  "^FT417,258^A0I,34,33^FH\^FDFunctionality Error^FS"
                  "^FT582,144^A0I,37,38^FH\^FDProgram:^FS"
                  "^FT460,262^A0I,37,38^FH\^FD:^FS"
                  "^FT581,279^A0I,37,38^FH\^FDReturn ^FS"
                  "^FT581,233^A0I,37,38^FH\^FDReason^FS"
                  "^FO429,37^GB0,293,3^FS"
                  "^PQ1,0,1,Y^XZ")
        new_trg = label1.replace('^FDProduct-12345678^FS', '^FD'+pid+'^FS')
        final = new_trg.replace('^FDFunctionality Error^FS', '^FD'+cond+'^FS')
        z = Zebra('ZDesigner GK420d')
        z.output(final)


Product()
