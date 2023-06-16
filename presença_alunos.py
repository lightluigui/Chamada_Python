import openpyxl as opn
import win32com.client as win32

tabela = opn.load_workbook("alunos.xlsx")

ativa = tabela.active
for celula in ativa['A']:
    if celula.value == None:
        linha = celula.row
        ativa[f'A{linha}'] = str(input('nome do aluno:'))

for celula in ativa['C']:
    if celula.value == None:
        linha = celula.row
        ativa[f'C{linha}'] = str(input('o aluno veio?:'))


tabela.save('alunosopn.xlsx')