import openpyxl
from flask import render_template_string
from flask import Flask
import pdfkit
import os 
import fitz
import json
import shutil
import pywhatkit as pwk

app = Flask(__name__)

wb = openpyxl.open("statement.xlsx")
skipped_first= False

if "done.json" in os.listdir("."):
    with open("done.json", "r") as f:
        done = json.load(f)
else:
    done = []

sent_to_numbers = []

def get_receipts(bookingid:str):
    if "statement2.xlsx" not in os.listdir("."):
        return []
    rs= []
    wb2 = openpyxl.open("statement2.xlsx")
    for row in wb2.active.rows:
        if bookingid in row[7].value:
            r = {
                "vr_no": row[3].value,
                "amount": row[14].value,
                "mode": row[2].value,
                "date": row[4].value.split("T")[0]
            }
            rs.append(r)    
    return rs



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
    netamount = row[6].value
    netadjusted = row[7].value
    projectname = row[8].value
    bookingid = row[9].value
    bookingdate = row[10].value.split("T")[0]
    output_file_name = f"file"
    print(bookingid)
    reciepts = get_receipts(bookingid)
    rtitles = "Receipts"
    if len(reciepts) > 15:
        rtitles = ""
        reciepts = []
    
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
            balance = netamount - netadjusted,
            receipts = reciepts,
            rtitle = rtitles
            )
        
    pdfkit.from_string(string, f"{output_file_name}.pdf", options={"zoom": 1.5})
    with open(f"{output_file_name}.html", "w+") as f:
        f.write(string)
    doc = fitz.open(f"{output_file_name}.pdf")
    page = doc.load_page(0)
    pix = page.get_pixmap()
    pix.save(f"{output_file_name}.png")
    text = """नमस्कार , 
ये मोनार्क बिल्डस्टेट प्राइवेट लिमिटेड की वट्सअप सेवा है।
अब आप अपने प्लाट की सारी जानकारी इस नंबर पर मैसेज अथवा कॉल  से प्राप्त कर सकते है ।
आपके बुक किए गए प्लॉट की जानकारी आपको भेजी गई है, कृप्या चेक कर लें।
अधिक जानकरी या किसी भी संशोधन के लिए निचे दिए गए आधिकारिक नंबरों पर सम्पर्क कर सकते है। 
धन्यवाद 
मोनार्क बिल्डस्टेट प्रा. लि. बीकानेर  
xxx"""
    if f"{custname} - {unitname} - {netadjusted}" in done:
        print("skipping")
        # add to sent_to_numbers
        sent_to_numbers.append(mobile)
        continue
    
    if mobile in sent_to_numbers:
        text = ""
    if "xxx" in mobile:
        print("skipping")
        continue

    sent_to_numbers.append(mobile)
    doc = fitz.open(f"{output_file_name}.pdf")
    page = doc.load_page(0)
    pix = page.get_pixmap()
    pix.save(f"{output_file_name}.png")
    doc.close()
    #pwk.sendwhats_image(f"+91{mobile}", f"{output_file_name}.", text, 10, True)
    done.append(f"{custname} - {unitname} - {netadjusted}")
    with open("done.json", "w") as f:
        json.dump(done, f, indent=4)
    # remove this file after sending
    os.remove(f"{output_file_name}.pdf")
    os.remove(f"{output_file_name}.html")
    os.remove(f"{output_file_name}.png")
    