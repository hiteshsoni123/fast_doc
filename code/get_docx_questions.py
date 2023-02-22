import lxml.etree as ET
from docx_utils.flatten import opc_to_flat_opc
import json 

MATH_NS = {'m': 'http://schemas.openxmlformats.org/officeDocument/2006/math'}
WORD_NAMESPACE = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
MATHML_NAMESPACE = '{http://schemas.openxmlformats.org/officeDocument/2006/math}'
PARA = WORD_NAMESPACE + 'p'
TEXT = WORD_NAMESPACE + 't'
TABLE = WORD_NAMESPACE + 'tbl'
ROW = WORD_NAMESPACE + 'tr'
CELL = WORD_NAMESPACE + 'tc'
MATH_TEXT = MATHML_NAMESPACE + 't'
MATH_PARA = './/m:oMathPara'

opc_to_flat_opc("sample_questions.docx", "sample_questions.xml")
tree  = ET.parse('sample_questions.xml')
mathml_xslt = ET.XSLT(ET.parse("xsltml_2.0/mmltex.xsl"))
xsltfile = ET.XSLT(ET.parse('OMML2MML.XSL'))

mathml_start = '<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage"><pkg:part pkg:name="word/document.xml" pkg:contentType="application/xml"><pkg:xmlData><w:document xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
mathml_end = '</w:document></pkg:xmlData></pkg:part></pkg:package>'
questions = []

for i, table in enumerate(tree.iter(TABLE)):
    # print('i==>', i)
    questions.append({"sno": i + 1, "question": "", "answer_1": "", "answer_2": "", "answer_3": "", "answer_4": "", "correct_ans": "", "explanation": ""})
    for j, row in enumerate(table.iter(ROW)):
        # print('j==>', j)
        for k, cell in enumerate(row.iter(CELL)):
            # print('k==>', k)
            cell_text = ''.join(node.text for node in cell.iter(TEXT))
            if cell_text == '':
              math_ml = ET.tostring(cell.find(MATH_PARA, MATH_NS))
              # print(math_ml)
              cell_text = xsltfile(ET.XML(mathml_start + str(math_ml)  + mathml_end))
              cell_text = str(mathml_xslt(cell_text))
            if(j == 0 and k == 1):
              questions[i]["question"] = cell_text
            if(j == 1 and k == 1):
              questions[i]["answer_1"] = cell_text
            if(j == 2 and k == 1):
              questions[i]["answer_2"] = cell_text
            if(j == 3 and k == 1):
              questions[i]["answer_3"] = cell_text
            if(j == 4 and k == 1):
              questions[i]["answer_4"] = cell_text
            if(j == 5 and k == 1):
              questions[i]["correct_ans"] = cell_text
            if(j == 6 and k == 1):
              questions[i]["explanation"] = cell_text
            # table_list.append(cell_text)

json_object = json.dumps(questions, indent = 2) 
print(json_object)

# student 
# teaches
# health info
# corse -> student(helping)
