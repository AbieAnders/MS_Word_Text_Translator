import os
#docx library
import docx
from docx.document import Document
from docx.text.paragraph import Paragraph 
from docx.table import _Cell, _Row, Table
from docx.oxml.text.paragraph import CT_P
from docx.oxml.table import CT_Tbl
#translatepy library
from translatepy import Translator

os.chdir('Path of the desired new working directory')
doc = docx.Document()
doc.save('Fully Translated Document.docx')
doc1 = docx.Document('Document of choice.docx')
print("The documents were accessed successfully.")

translator_object = Translator()
def Translate_Text(file_text):
    try:
        translation = translator_object.translate(file_text,'ta')
        if translation is None:
            return None
        return str(translation)
    except IndexError:
        return None
    '''except Exception as e:
        print("An error has occurred during translation:", e)
        return None'''

def iter_block_items(parent):
    """
    Generate a reference to each paragraph and table child within *parent*,
    in document order. Each returned value is an instance of either Table or
    Paragraph. *parent* would most commonly be a reference to a main
    Document object, but also works for a _Cell object, which itself can
    contain paragraphs and tables.
    """
    if isinstance(parent, Document):
        parent_elm = parent.element.body
    elif isinstance(parent, _Cell):
        parent_elm = parent._tc
    elif isinstance(parent, _Row):
        parent_elm = parent._tr
    else:
        raise ValueError("something's not right")
    for child in parent_elm.iterchildren():
        if isinstance(child, CT_P):
            yield Paragraph(child, parent_elm)
        elif isinstance(child, CT_Tbl):
            yield Table(child, parent_elm)

#execution loop
para_list = []
text_segment = ""
for count, block in enumerate(iter_block_items(doc1)):
    print("block count:", count)
    if isinstance(block, Paragraph):
        try:
            if block.text.strip() == '':
                raise Exception('Whitespace')
            #print("#Exception is not being called.")
            #print(block.text, "##Has been appended")
            para_list.append(block.text) #doesnt append after block 264 but the rest of the program still executes and fails at block 310(whole thing crumbles)
        except Exception as e:
            #print("#Exception is being called.")
            #print(para_list)
            text_segment = '\n'.join(para_list)
            data = Translate_Text(text_segment)
            #data = text_segment
            print(data)
            doc.add_paragraph(data)
            para_list = []
            text_segment = ""
            doc.save('Fully Translated Document.docx')
            continue
    elif isinstance(block, Table):
        n = len(block.rows)
        m = int(len(block._cells)/n)
        print(n, m)
        table = doc.add_table(n, m)
        for a,row in enumerate(block.rows):
            for b,cell in enumerate(row.cells):
                translation = translator_object.translate(cell.text,'ta')
                table.cell(a, b).text = str(translation)  #str type casting is necessary since the text function automatically assumes the TranslationResult type
    doc.save('Fully Translated Document.docx') #saves the document after every individual table or paragraph is added to the document.
#table.style = 'Colorful List'
#doc.save('Fully Translated Document.docx') saves the document after every paragraph and table has been successfully added to the document.
#Use the above statement in case the code works perfectly and remove all the other saves, since it eliminates the need for the other unnecessary saves.
