import openpyxl
wb = openpyxl.open("statement.xlsx")
import pywhatkit as pwk
import shutil
import os
from time import sleep
import json

if "done.json" in os.listdir("."):
    with open("done.json", "r") as f:
        done = json.load(f)
else:
    done = []
# print(done)



import fitz
sent_to_numbers = []
count = 0
skipped_first = False
for row in wb.active.rows:
    if not skipped_first:
        skipped_first = True
        continue
    print(row[0].value)
    count+=1
    print(count)
    
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
    # copy the pdf to a folder so that a person can drag drop it easily
    for file in os.listdir("target"):
        os.remove(f"target/{file}")
    shutil.copy(f"output/{output_file_name}.pdf", f"target/{output_file_name}.pdf")
    doc = fitz.open(f"target/{output_file_name}.pdf")
    page = doc.load_page(0)
    pix = page.get_pixmap()
    output = "outfile.png"
    pix.save("target/"+output)
    doc.close()
    pwk.sendwhats_image(f"+91{mobile}", f"target/{output}", text, 10, True)
    done.append(f"{custname} - {unitname} - {netadjusted}")
    with open("done.json", "w") as f:
        json.dump(done, f, indent=4)
    print("after")
    

    
print("done")
    
    