import os, openpyxl
from openpyxl.styles import Alignment 

import pandas as pd

import matplotlib.pyplot as plt

from fpdf import FPDF

print("Análisando planilha...")
# Criando a planilha
planilha= openpyxl.Workbook()

page= planilha['Sheet']
page.title= "Dados da covid"

page= planilha.active

# Adicionando registros
page['A1'].value= "Data" 
page['B1'].value= "Novos casos" 
page['C1'].value= "Óbitos" 

page['A2'].value= "01/2" 
page['B2'].value= "17130" 
page['C2'].value= "206" 

page['A3'].value= "02/2" 
page['B3'].value= "1096" 
page['C3'].value= "363" 

page['A4'].value= "03/2" 
page['B4'].value= "37611" 
page['C4'].value= "349" 

page['A5'].value= "04/2" 
page['B5'].value= "18535" 
page['C5'].value= "370" 

page['A6'].value= "05/2"
page['B6'].value= "17066"
page['C6'].value= "294" 

page['A7'].value= "06/2" 
page['B7'].value= "5351" 
page['C7'].value= "53" 

page['A8'].value= "07/2" 
page['B8'].value= "3585" 
page['C8'].value= "26" 

page['A9'].value= "08/2" 
page['B9'].value= "15696" 
page['C9'].value= "445" 

page['A10'].value= "09/2" 
page['B10'].value= "17464" 
page['C10'].value= "482" 

page['A11'].value= "10/2" 
page['B11'].value= "19046" 
page['C11'].value= "297" 

# ALinhamento das células
page['A1'].alignment= Alignment(horizontal= 'center')
page['B1'].alignment= Alignment(horizontal= 'center')
page['C1'].alignment= Alignment(horizontal= 'center') 
page['A2'].alignment= Alignment(horizontal= 'center') 
page['B2'].alignment= Alignment(horizontal= 'center')
page['C2'].alignment= Alignment(horizontal= 'center') 
page['A3'].alignment= Alignment(horizontal= 'center') 
page['B3'].alignment= Alignment(horizontal= 'center')
page['C3'].alignment= Alignment(horizontal= 'center') 
page['A4'].alignment= Alignment(horizontal= 'center')
page['B4'].alignment= Alignment(horizontal= 'center')
page['C4'].alignment= Alignment(horizontal= 'center') 
page['A5'].alignment= Alignment(horizontal= 'center')
page['B5'].alignment= Alignment(horizontal= 'center') 
page['C5'].alignment= Alignment(horizontal= 'center') 
page['A6'].alignment= Alignment(horizontal= 'center')
page['B6'].alignment= Alignment(horizontal= 'center')
page['C6'].alignment= Alignment(horizontal= 'center') 
page['A7'].alignment= Alignment(horizontal= 'center')
page['B7'].alignment= Alignment(horizontal= 'center')
page['C7'].alignment= Alignment(horizontal= 'center') 
page['A8'].alignment= Alignment(horizontal= 'center') 
page['B8'].alignment= Alignment(horizontal= 'center')
page['C8'].alignment= Alignment(horizontal= 'center') 
page['A9'].alignment= Alignment(horizontal= 'center')
page['B9'].alignment= Alignment(horizontal= 'center')
page['C9'].alignment= Alignment(horizontal= 'center') 
page['A10'].alignment= Alignment(horizontal= 'center')
page['B10'].alignment= Alignment(horizontal= 'center')
page['C10'].alignment= Alignment(horizontal= 'center') 


planilha.save("dados.xlsx")
print("")
print("")
print("Aguarde um instante...")
print("")
print("")
print("Gerando os gráficos....")

# Fazendo os gráficos
planilha= pd.read_excel("dados.xlsx")

dia= planilha['Data']
casos= planilha['Novos casos']
obitos= planilha['Óbitos']

plt.title("Casos nos 10 primeiros dias 1 mês após último natal e ano novo")
plt.bar(dia, casos, color='blue', width=0.5)
plt.grid()
plt.savefig("casos.png")
plt.show()


plt.title("Óbitos nos 10 primeiros dias 1 mês após último natal e ano novo")
plt.bar(dia, obitos, color='red', width=0.5)
plt.grid()
plt.savefig("obitos.png")
plt.show()
print("Gerando os gráficos....")
print("")
print("")


# Gerando o pdf

pdf= FPDF('P', 'mm', 'A4')

pdf.add_page()
pdf.set_font('Times', '', 14)
pdf.multi_cell(w=0, h=8, txt="Pedro de Carvalho Moraes\n\n Analisando dados da Covid-19 1 mês após festas, Natal e Ano novo\n\nComo podemos observar nos gráficos elaborados de acordo com a página https://www.seade.gov.br/coronavirus/ (os gráficos estão situados na segunda e terceira página do pdf), percebemos que os casos de Covid-19 explodiram após as aglomerações de Natal e Ano novo, mas as mortes diárias por Covid-19 diminuíram (se formos comparar com as mortes que teve no último pico da Covid-19 em 2021), graças a vacinação em massa da população.\nO agravamento de casos interfere principalmente nos comércios e nas escolas, pois essas instituições correm o risco do estado decretar lockdown, ou seja, terão que fechar completamente por tempo indeterminado até os casos diminuírem novamente, igual aconteceu nos últimos períodos que a covid estava em fase crítica. Isso é um problema, pois enfraquece os comerciantes financeiramente e as as aulas terão que retornar EAD.\n\n Explicando melhor...Qual o problema disso tudo?\n\nEm relação aos comerciantes:\nO enfraquecimento dos comerciantes, leva-os a falência tendo que demitir vários funcionários prejudicando muitas famílias que tinham o trabalho no comércio como principal fonte de renda delas, algumas até vão para as ruas por não terem dinheiro para se sustentar.\n\nEm relação aos estudantes:\nA aula EAD causou muitos problemas psciológicos nos alunos, pois muitos não aguentaram ficar isolados em casa por 2 anos e acabaram entrando em depressão, uma doença muito séria e delicadíssima, e também, muitas pessoas não possuem os requisitos tecnológicos mínimos para ter uma aula EAD descente e acabam não aprendendo nada, resultando em um ano perdido na escola.\n\nConclusão\nSe as pessoas tivessem respeitado o isolamento social entre o final do ano de 2021 e no comecinho de 2022, não teríamos esse agravamento de casos\nFrase: 'Quais os principais problemas causados pelo agravamento da Covid-19 para os próximos meses?'", align='J')

pdf.image(name="casos.png", x=0, y=100, w=200)

pdf.add_page()
pdf.image(name="obitos.png", x=0, y=100, w=200)

pdf.output("relatorio.pdf")
print("PDF criado")
print("")
print("")
print("Pronto! Programa finalizado!")