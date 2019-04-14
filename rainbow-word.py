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


totalWords = []

# Loop through the existing paragraphs, delete the old ones and replace them with the colorful ones.
for p in allTxt:
    # Remove existing paragraphs
    delete_paragraph(p)
    # Split every word between a space
    words = p.text.split(' ')
    # Push the words in the totalWords array
    totalWords.append(words)

for word in range(0, len(totalWords)):
    run = doc.add_paragraph().add_run(totalWords[word])
    font = run.font

    # Set font color to red
    font.color.rgb = RGBColor(255, 0, 0)

print(totalWords)

# Create a new file and save it
doc.save('demo1.docx')
