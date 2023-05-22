# стандартный шаблон Word
from docxtpl import DocxTemplate
# import pydocxtpl
tpl = DocxTemplate('707-new.docx')
# tpl.render(context_dict)
set_of_variables = tpl.get_undeclared_template_variables()
print(set_of_variables,len(set_of_variables))