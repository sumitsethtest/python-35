from wand.image import Image
from PIL import Image as PI
import pyocr as pyocr1
import pyocr.builders
import io

outfile = open("test.txt","wb")

path = "C:\\upw\\chin\\test.pdf"
tool = pyocr1.get_available_tools()[0]
lang = tool.get_available_languages()[0]

req_image = []
final_text = []

image_pdf = Image(filename=path, resolution=300)
image_jpeg = image_pdf.convert('jpeg')

for img in image_jpeg.sequence:
    img_page = Image(image=img)
    req_image.append(img_page.make_blob('jpeg'))

for img in req_image:
    txt = tool.image_to_string(
        PI.open(io.BytesIO(img)),
        lang=lang,
        builder=pyocr.builders.TextBuilder()
    )
    final_text.append(txt.encode("utf-8"))
    #outfile.write(txt.encode("utf-8"))
outfile.writelines(final_text)
print(final_text)
outfile.close()