import os
#docx library
import docx
from docx.document import Document
from docx.text.paragraph import Paragraph
from docx.enum.style import WD_STYLE_TYPE
from docx.table import _Cell, _Row, Table
from docx.oxml.text.paragraph import CT_P
from docx.oxml.table import CT_Tbl
from docx.shared import Pt
#translatepy library
from translatepy import Translator

#Path of the desired new working directory
os.chdir('Path of the desired new working directory')
#New document
doc = docx.Document()
doc.save('Fully Translated Document 2.docx')
#Document of choice
doc1 = docx.Document('Document of choice.docx')
#Throwaway document
doc2 = docx.Document()
doc2.save('Throwaway Document.docx')
print("The documents were accessed successfully.")

translator_object = Translator()
def Translate_Text(file_text):
    try:
        translation = translator_object.translate(file_text,'ta')
        if translation is None:
            return None
        return str(translation)
    except IndexError:
        print("Index error")
        return None
    except Exception as e:
        print("An error has occurred during translation:", e)
        return None

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

#Execution loop

#Style and Font info are gathered here
font_list = []
style_list = []
para_counter = 1
for para in doc1.paragraphs:
    print("block count:", para_counter)
    if para.text.strip() == '':
        print(para_counter)
        continue
    font_list.append(para.style.font.name)
    style_list.append(para.style.name)
    para_counter += 1
print(font_list)
print(style_list)

#The 'Arial' font style is created here since direct access of the font style is not supported by the library
doc_styles = doc.styles
doc_charstyle = doc_styles.add_style('Ordinary Arial', WD_STYLE_TYPE.CHARACTER)
obj_font = doc_charstyle.font
obj_font.name = 'Arial'

#Exists to simply check the style implementation in the English language(causes performance hit so remove it and the statements connected to it if it isnt necessary)
doc_styles = doc2.styles
doc_charstyle = doc_styles.add_style('Ordinary Arial', WD_STYLE_TYPE.CHARACTER)
obj_font = doc_charstyle.font
obj_font.name = 'Arial'

#Main execution
font_counter = 0
style_counter = 0
for count, block in enumerate(iter_block_items(doc1)):
    print("block count:", count)
    if isinstance(block, Paragraph):
        if block.text.strip() == '':
            print(count)
            continue
        if(font_list[font_counter] == 'Arial'):  #(or) if(style_list[style_counter] != List Paragraph or Body Text)
            translated_text = Translate_Text(block.text)
            para = doc.add_paragraph()
            para_runner = para.add_run(translated_text, style = 'Ordinary Arial').bold = True
            '''
            This section is unnecessary if output in English is not required
            #Note:This section only activates if the paragraph is found to be a heading(always has 'Arial' font type in the given document)
            para2 = doc2.add_paragraph()
            para_runner2 = para2.add_run(block.text, style = 'Ordinary Arial').bold = True
            doc.save('Fully Translated Document 2.docx')
            doc2.save('Proxy Document.docx')
            '''
            font_counter += 1    #Used since the continue statement prevents the font_counter statement that is found below from executing
            continue
        block = Translate_Text(block.text)
        doc.add_paragraph(block)
        doc.save('Fully Translated Document 2.docx') #Can also be omitted since all of the changes are being saved in line 129
    elif isinstance(block, Table):
        n = len(block.rows)
        m = int(len(block._cells)/n)
        print(n, m)
        table = doc.add_table(n, m)
        for a,row in enumerate(block.rows):
            for b,cell in enumerate(row.cells):
                translation = translator_object.translate(cell.text,'ta')
                table.cell(a, b).text = str(translation)  #str type casting is necessary since the text function automatically assumes the TranslationResult type
        doc.save('Fully Translated Document 2.docx')
    font_counter += 1
    style_counter += 1

#table.style = 'Colorful List'
doc.save('Fully Translated Document 2.docx')
