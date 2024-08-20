import pandas as pd
import win32com.client as win32
import os
import io
from dotenv import load_dotenv
from buttons import get_month, get_option


# -------------------------------------------------------------------------------------------


# DADOS DO CONTRATO
# Lê a relação de e-mails das empresas - Excel
load_dotenv()
SHEET_PATH = os.getenv('SHEET_PATH')
dfContratos = pd.read_excel(SHEET_PATH, sheet_name='Listagens')
numContrato = dfContratos['Nº Contrato'].tolist()
nomeContrato = dfContratos['Contrato'].tolist()

# Retorna informações dos contratos (apenas para visualização)
print('Lista dos contratos trabalhados:')
print('------------------------------------')
col1 = 0
col2 = 0
qtdContratos = []
while(col1 < len(dfContratos)):
    while(col2 < len(dfContratos)):
        print(numContrato[col1] + " (" + nomeContrato[col2] + ")")
        qtdContratos.append(numContrato[col1])
        col2 = col2 + 1
        break
    col1 = col1 + 1
print('------------------------------------')
print(f'{len(qtdContratos)} contratos encontrados.')


# -------------------------------------------------------------------------------------------


# LISTAGEM DE E-MAILS E DESTINATÁRIOS
# COORDENADORIA - CAC
mailingSetor = ["felipe.matos@tjce.jus.br; lucas.rocha@tjce.jus.br; valeria.monteiro@tjce.jus.br; fransilvia.paiva@tjce.jus.br; ana.sousa6@tjce.jus.br"]

# EMPRESAS
mailingEmpresas = []
mailingEmpresasTest = []
mailingEmpresasEmpty = []
for i in range(len(dfContratos)):
    contrato = dfContratos.loc[i, 'Contrato']
    destinatario = dfContratos.loc[i, 'Endereços de Emails']
    destin_teste = dfContratos.loc[i, 'Emails Teste']    
    if pd.isna(destinatario):
        mailingEmpresasEmpty.append(contrato)
        print(f'Contrato(s) com grupo de e-mails vazio: {mailingEmpresasEmpty}\n')
    else:
        # Adicionar o email à lista do mesmo grupo
        if i > 0 and contrato == dfContratos.loc[i-1, 'Contrato']:
            mailingEmpresas[-1].append(destinatario)
            mailingEmpresasTest[-1].append(destin_teste)
            
        # Criar um novo grupo e adicionar o primeiro email
        else:
            mailingEmpresas.append([destinatario])
            mailingEmpresasTest.append([destin_teste])


# Listar emails
print('Lista dos e-mails trabalhados:')
print('------------------------------------')
for emails in mailingEmpresas:
    print(emails)
print('------------------------------------')
print(f'{len(mailingEmpresas)} grupo de e-mails encontrados.')


# -------------------------------------------------------------------------------------------


# ELABORAÇÃO DO EMAIL
def create_email(assunto, anexo, corpoEmail, emailTo, emailCC=None):
    outlook = win32.Dispatch('outlook.application')
    email = outlook.CreateItem(0)
    email.Subject = assunto
    email.To = ";".join(emailTo)
    email.CC = ";".join(emailCC) if emailCC is not None else ""
    email.Attachments.Add(anexo)
    email.HTMLBody = corpoEmail
    email.Send()


# -------------------------------------------------------------------------------------------


# EXECUÇÃO DO EMAIL
i = 0
ref_month = get_month()
option = get_option()
print(f'\nMês de referência: {ref_month}\n')
qtd_emails = []

print('Lista dos e-mails enviados:')
print('------------------------------------')
# ENVIO DO EMAIL
while i < len(mailingEmpresas):

    assunto = f"""Listagem atualizada dos colaboradores ativos - {ref_month}/2024 - CT Nº {numContrato[i]} ({nomeContrato[i]})"""
    load_dotenv()
    anexo = os.getenv('ATTACHMENT')
    corpoEmail = f"""
    <p>Prezados, boa tarde.</p>

    <p>Por gentileza, providenciar a <strong>LISTAGEM ATUALIZADA DE MOVIMENTAÇÕES DE COLABORADORES</strong> 
    durante o mês de <strong>{ref_month}</strong>, referente ao <strong>CONTRATO Nº {numContrato[i]} ({nomeContrato[i]})</strong>, 
    em formato de planilha, contendo as movimentações deste mês ocorridas até o presente momento.</p>

    <p>Enfatizo que deve ser enviada a planilha modelo padrão, a qual segue anexa, devidamente preenchida.</p>

    <p><em>Obs.: Reforço ainda que na planilha, há as orientações do quê e como deve ser preenchido. Pedimos também, por gentileza, 
    que não alterem a quantidade de linhas ou de colunas da planilha nem nome de campos ou de abas. Por fim, mantemo-nos à disposição para sanar quaisquer dúvidas.</em>

    <p>Gratidão!</p>

    <p></p>
    <p>Atenciosamente,</p>

    <strong><em><font color="#00642D"; font size=2>Luigi Teles Alcântara de Castro</font></em></strong><br>
    <em><font size=2>Coordenadoria de Acombanhamento de Contratos (CAC)</font></em><br>
    <em><font size=2>Tribunal de Justiça do Estado do Ceará</font></em><br>
    """
    
    if option.upper() == "TESTE":
        create_email(assunto, anexo, corpoEmail, mailingEmpresasTest[i])
        print(f"E-mail teste do contrato {nomeContrato[i]} enviado para: {mailingEmpresasTest[i]}")
    else:
        create_email(assunto, anexo, corpoEmail, mailingEmpresas[i], mailingSetor)
        print(f"E-mail do contrato {nomeContrato[i]} enviado para: {mailingEmpresas[i]}")
    
    qtd_emails.append([create_email])
    i += 1
    
print('------------------------------------')
print(f'Qtd de e-mails enviados: {len(qtd_emails)}')
print(f"\nE-mail não enviado para o(s) contrato(s): {mailingEmpresasEmpty}")
    

