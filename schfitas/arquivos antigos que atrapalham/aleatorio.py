import docx 
from docx.enum.table import WD_TABLE_ALIGNMENT

doc = docx.Document() 
infoLinha = 'excel_data'
infoForma = 'excel_data'
infoCód = 'excel_data'

data = (
    ( 'Linha: ' + infoLinha, 'Geek 1'), 
    ('Forma: ' + infoForma, 'Geek 2'), 
    ('Cód: ' + infoCód, 'Geek 3') 
)
table = doc.add_table(rows=1, cols=2) 
table.alignment = WD_TABLE_ALIGNMENT.CENTER
row = table.rows[0].cells 
row[0].text = 'Linha' + infoLinha
row[1].text = 'Name'

for id, name in data: 
  
    row = table.add_row().cells 
    
    row[0].text = str(id)
    row[1].text = name 
doc.save('gfg.docx')