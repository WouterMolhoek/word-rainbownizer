from docx import Document
from docx.shared import RGBColor

doc = Document('test-file.docx')

# All the paragraphs stored in an array
allTxt = doc.paragraphs

# Function to delete a paragraph
def delete_paragraph(paragraph):
    p = paragraph._element
    p.getparent().remove(p)
    p._p = p._element = None

# Loop through the existing paragraphs, delete the old ones and replace them with the colorful ones.
for p in allTxt:
    # Remove existing paragraphs
    delete_paragraph(p)
    # Create new paragraph
    run = doc.add_paragraph().add_run(p.text)
    font = run.font
    # Set font color to red
    font.color.rgb = RGBColor(255, 0, 0)
    # Set font color to orange
    #font.color.rgb = RGBColor(255, 165, 0)
    # Set font color to yellow
    #font.color.rgb = RGBcolor(255, 215, 0)


# Create a new file and save it
doc.save('demo1.docx')