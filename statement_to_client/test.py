import openpyxl
wb = openpyxl.open("statement.xlsx")
import pywhatkit as pwk
import fitz


mobile = "xxx"
text = "Hello world"
doc = fitz.open(f"target/1.pdf")
page = doc.load_page(0)
pix = page.get_pixmap()
output = "outfile.png"
pix.save("target/"+output)
pwk.sendwhats_image(f"+91{mobile}", f"target/{output}", text, 10, True)

    
    