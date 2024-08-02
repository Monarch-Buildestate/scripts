import openpyxl
from flask import render_template_string
from flask import Flask
import pdfkit
import os 

app = Flask(__name__)

wb = openpyxl.open("statement.xlsx")
skipped_first= False
for row in wb.active.rows:
    if not skipped_first:
        skipped_first = True
        continue
    print(row[0].value)
    customer_code = row[0].value
    calc_area = row[1].value
    unitrate = row[2].value
    unitname = row[3].value
    custname = row[4].value
    mobile = row[5].value
    agent = row[6].value
    netamount = row[7].value
    netadjusted = row[8].value
    projectname = row[9].value
    bookingid = row[10].value
    bookingdate = row[11].value.split("T")[0]
    output_file_name = f"{custname}_{unitname}_{mobile}".replace("/", "_")
    if output_file_name+".pdf" in os.listdir("./output"):
        continue
    with app.app_context():
        template = open("templates/print_receipt.html").read()
        string = render_template_string(
            template, 
            pname=projectname,
            uname=unitname, 
            bdate=bookingdate, 
            bno=bookingid, 
            cname=custname, 
            mob=mobile, 
            netpay=netamount, 
            paidamount=netadjusted, 
            balance = netamount - netadjusted
            )
        
    pdfkit.from_string(string, f"./output/{output_file_name}.pdf", options={"zoom": 1.5})
    with open(f"./html/{output_file_name}.html", "w+") as f:
        f.write(string)
