import pandas as pd
from pathlib import Path
import warnings
import win32com.client as win32

outlook = win32.Dispatch('outlook.application')
warnings.simplefilter(action='ignore', category=FutureWarning)



lojas_df = pd.read_csv(r'C:\Users\Pichau\Desktop\Projetos Python\1\Projeto AutomacaoIndicadores\Bases de Dados\Lojas.csv',encoding=' iso-8859-1',sep= ';')
email_df = pd.read_excel(r'C:\Users\Pichau\Desktop\Projetos Python\1\Projeto AutomacaoIndicadores\Bases de Dados\Emails.xlsx')
vendas_df = pd.read_excel(r'C:\Users\Pichau\Desktop\Projetos Python\1\Projeto AutomacaoIndicadores\Bases de Dados\Vendas.xlsx')
vendas = vendas_df.merge(lojas_df,on='ID Loja')
d_lojas = {}
for loja in lojas_df['Loja']:
    d_lojas[loja] = vendas.loc[vendas['Loja']==loja ,:]

##### criando os arquivos de cada loja
for loja in d_lojas:
    d_lojas[loja].to_excel(f'{loja}.xlsx',index=False)

##### realocando os arquivos de backup para a pasta feita
for x in lojas_df['Loja']:
    Path(f'{x}.xlsx').rename(fr'C:\Users\Pichau\Desktop\Projetos Python\1\Projeto AutomacaoIndicadores\Backup Arquivos Lojas\{x}.xlsx')

fat = 0
fat_meta_d = 1000
fat_meta_a = 1650000
div = 0
div_meta_d = 4
div_meta_a = 120
tic = 0
tic_meta_d = 500
tic_meta_a = 500
dia = '2019-12-25'
    # (vendas['Data'].iloc)[-1]


for lojas in d_lojas:
    a = d_lojas[lojas]
    b =a.loc[a['Data']== dia,:]
    #vendas

    vendas_dia_1 = (b[['Valor Final']].sum(axis=0))
    vendas_dia = round(vendas_dia_1,2).to_string(index=False)
    cen_vendas_dia = 'GREEN'
    if float(vendas_dia_1) < fat_meta_d:
        cen_vendas_dia = 'RED'

    vendas_ano_1 = (a[['Valor Final']].sum(axis=0))
    vendas_ano = round(vendas_ano_1,2).to_string(index=False)
    cen_vendas_ano = 'GREEN'
    if float(vendas_ano_1) < fat_meta_a:
        cen_vendas_ano = 'RED'

#ticket
    ticket_dia_1 = (b[['Valor Final']].sum(axis=0) / int(b.groupby(by=['Código Venda']).size().count()))
    ticket_dia = round(ticket_dia_1,2).to_string(index=False)
    cen_tic_dia = 'GREEN'
    if float(ticket_dia_1) < tic_meta_d:
        cen_tic_dia = 'RED'

    ticket_ano_1 = (a[['Valor Final']].sum(axis=0) / int(a.groupby(by=['Código Venda']).size().count()))
    ticket_ano = round(ticket_ano_1,2).to_string(index=False)
    cen_tic_ano = 'GREEN'
    if float(ticket_ano_1) < tic_meta_a:
        cen_tic_ano = 'RED'

    #diversidade
    diver_dia_1 =b.groupby(by=['Produto']).size().count()
    diver_dia = diver_dia_1.astype(str)
    cen_diver_dia = 'GREEN'
    if float(diver_dia_1) < div_meta_d:
        cen_diver_dia = 'RED'

    diver_ano_1 = a.groupby(by=['Produto']).size().count()
    diver_ano = diver_ano_1.astype(str)
    cen_diver_ano = 'GREEN'
    if float(diver_ano_1) < div_meta_a:
        cen_diver_ano = 'RED'

    gerente = email_df.loc[email_df['Loja']==lojas,:]['Gerente'].to_string(index=False)
    email_gerente = email_df.loc[email_df['Loja']==lojas,:]['E-mail'].to_string(index=False)

    mail = outlook.CreateItem(0)
    mail.to = email_gerente
    mail.Subject = f'Report diário da loja {lojas}'
    texto = f'''
    Bom dia {gerente},
    O resultado de ontem {dia} da Loja {lojas} foi :
    
    --------------Valor dia-----------Meta dia----------Cenário dia
    Faturamento---R${vendas_dia}{'-'*(21-len(vendas_dia))}R${fat_meta_d}{'-'*(21-len(str(fat_meta_d)))}{cen_vendas_dia}
    Ticket Médio--R${ticket_dia}{'-'*(21-len(ticket_dia))}R${tic_meta_d}{'-'*(21-len(str(tic_meta_d)))}{cen_tic_dia}
    Diversidade---{diver_dia}{'-'*(21-len(diver_dia))}{div_meta_d}{'-'*(20-len(str(div_meta_d)))}{cen_diver_dia}
    
    --------------Valor Ano-----------Meta Ano----------Cenário Ano
    Faturamento---R${vendas_ano}{'-'*(21-len(vendas_ano))}R${fat_meta_a}{'-'*(21-len(str(fat_meta_a)))}{cen_vendas_ano}
    Ticket Médio--R${ticket_ano}{'-'*(21-len(ticket_ano))}R${tic_meta_a}{'-'*(21-len(str(tic_meta_a)))}{cen_tic_ano}
    Diversidade---{diver_ano}{'-'*(21-len(diver_ano))}{div_meta_a}{'-'*(20-len(str(div_meta_a)))}{cen_diver_ano}
    
    
    Segue em anexo a planilha com todos os dados para mais detalhes.
    Qualquer dúvida estou à disposição
           
'''
    mail.Body = texto
    attachment = fr'C:/Users/Pichau/Desktop/Projetos Python/1/Projeto AutomacaoIndicadores/Backup Arquivos Lojas/{lojas}.xlsx'
    mail.Attachments.Add(attachment)

    mail.Send()

#encontrando os dados da diretoria:

email_diretoria = email_df.loc[email_df['Loja']=='Diretoria',:]['E-mail'].to_string(index=False)
diretor = email_df.loc[email_df['Loja']=='Diretoria',:]['Gerente'].to_string(index=False)


faturamento_dia = vendas.loc[vendas['Data']==dia,:].groupby(['ID Loja']).sum(numeric_only=True)
faturamento_total = vendas.groupby(['ID Loja']).sum(numeric_only=True)

faturamento_dia_ordenado = faturamento_dia.sort_values(by=['Valor Final']).merge(lojas_df,on='ID Loja')
faturamento_total_ordenado = faturamento_total.sort_values(by=['Valor Final']).merge(lojas_df,on='ID Loja')

melhores_lojas_dia = faturamento_dia_ordenado.tail(3)[['Loja','Valor Final']].to_string(index=False)
melhores_lojas_ano = faturamento_total_ordenado.tail(3)[['Loja','Valor Final']].to_string(index=False)
piores_lojas_dia = faturamento_dia_ordenado.head(3)[['Loja','Valor Final']].to_string(index=False)
piores_lojas_ano = faturamento_total_ordenado.head(3)[['Loja','Valor Final']].to_string(index=False)

mail = outlook.CreateItem(0)
mail.to = email_diretoria
mail.Subject = f'Report diário das lojas'
texto = f'''
Bom dia {diretor},
Olá {diretor},
Segue dados das 3 melhores e piores lojas até o dia anterior:

3 MELHORES LOJAS DO DIA:
{melhores_lojas_dia}
--------------------------------
3 MELHORES LOJAS DO ANO:
{melhores_lojas_ano}
********************************
3 PIORES LOJAS DO DIA:
{piores_lojas_dia}
--------------------------------
3 PIORES LOJAS DO ANO:
{piores_lojas_ano}
********************************
Em anexo, segue uma planilha completa com todos os dados por dia, caso queira tirar mais alguma dúvida.
Tambem estou disponível para qualquer esclarecimento.

'''

mail.Body = texto

mail.Send()







