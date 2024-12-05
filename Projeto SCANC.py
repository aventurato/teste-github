##### ----- BIBLIOTECAS ----- #####
import os
import pandas as pd
import time
import re
from datetime import datetime
import win32com.client as client


##### ----- VERIFICAR O TEMPO DE EXECUÇÃO ----- #####
inicio = time.time()

data_atual = datetime.now()
print(data_atual)
mes_ano = data_atual.strftime('%m-%Y')
print(f'Mês e ano atual: {mes_ano}')

##### ----- ACESSANDO PASTA (SCANC) ----- #####
pasta_atual = os.getcwd()
pasta_scanc = os.path.join(pasta_atual, 'SCANC')
os.chdir(pasta_scanc)
arquivos_scanc = os.listdir()


##### ----- CARREGANDO DADOS (SCANC) ----- #####
base_scanc = os.path.join(pasta_scanc, 'YMMR048.xlsx')
dados_scanc = pd.read_excel(base_scanc)


##### ----- CARREGANDO E-MAILS (SCANC) ----- #####
#emails_scanc = pd.read_excel(base_scanc, sheet_name='Email_Teste')
#print(emails_scanc.head())


##### ----- LIMPEZA DADOS (COLUNAS) ----- #####
dados_scanc.drop(['Exercício', 'Período atual', 'Emitente', 'Modelo nota fiscal', 'Cód.sit.doc.', 'Séries',
                  'CPF', 'Local', 'Chave acesso 44 posições', 'Data documento', 'Data de lançamento',
                  'Valor Documento', 'Valor Mercadoria', 'Mod.Frete', 'Valor BC ICMS', 'Valor ICMS', 'Valor BC ICMS ST',
                  'Valor ICMS ST','Nº item do documento', 'Material', 'Código de controle', 'Unid.medida básica',
                  'Val Mercadoria', 'Descrição', 'Texto ICMS', 'BC ICMS', 'Alíq. ICMS', 'Valor ICMS.1',
                  'BC Outras ICMS', 'Val. Outras ICMS', 'Val. Isentas ICMS', 'BC ICMS ST', 'Alíq. ICMS ST dest.',
                  'Valor ICMS ST.1', 'BC ICMS ST cobrado ant.', 'ICMS ST cobrado ant.', 'BC ICMS ST UF destino',
                  'ICMS ST UF destino'], axis=1, inplace=True)


##### ----- TRANFORMANDO CNPJ CONGÊNERE (SCANC) ----- #####
def formatar_cnpj(cnpj):
    if pd.isna(cnpj):
        return 'CNPJ Inválido'

    cnpj = str(int(cnpj))
    cnpj = re.sub(r'\D', '', cnpj)
    cnpj = cnpj.zfill(14)
    return f'{cnpj[:2]}.{cnpj[2:5]}.{cnpj[5:8]}.{cnpj[8:12]}-{cnpj[12:]}'

dados_scanc['CNPJ'] = dados_scanc['CNPJ'].apply(formatar_cnpj)


##### ----- FUNÇÃO CNPJ ROYAL FIC ----- #####
def cnpj_fic(x, c):
    if x == 'FI02':
        return '01.349.764/0004-00'
    elif x == 'FI03':
        return '01.349.764/0008-26'
    elif x == 'FI05':
        return '01.349.764/0010-40'
    elif x == 'FI06':
        return '01.349.764/0011-21'
    elif x == 'FI08':
        return '01.349.764/0013-93'
    elif x == 'FI09':
        return '01.349.764/0014-74'
    elif x == 'FI10':
        return '01.349.764/0015-55'
    elif x == 'FI11':
        return '01.349.764/0016-36'
    elif x == 'FI12':
        return '01.349.764/0017-17'
    elif x == 'FI14':
        return '01.349.764/0019-89'
    elif x == 'FI16':
        return '01.349.764/0021-01'
    elif x == 'FI17':
        return '01.349.764/0023-65'
    elif x == 'FI18':
        return '01.349.764/0025-27'
    elif x == 'FI22':
        return '01.349.764/0031-75'
    elif x == 'FI24':
        return '01.349.764/0029-50'
    elif x == 'FI26':
        return '01.349.764/0034-18'
    elif x == 'FI29':
        return '01.349.764/0041-47'
    elif x == 'FI30':
        return '01.349.764/0035-07'
    elif x == 'FI32':
        return '01.349.764/0040-66'
    elif x == 'FI34':
        return '01.349.764/0046-51'
    elif x == 'FI36':
        return '01.349.764/0043-09'
    elif x == 'FI37':
        return '01.349.764/0047-32'
    elif x == 'FI38':
        return '01.349.764/0038-41'
    elif x == 'FI39':
        return '01.349.764/0048-13'
    elif x == 'FI40':
        return '01.349.764/0049-02'
    else:
        return 'Erro - CNPJ Não Preenchido!'

dados_scanc['CNPJ FIC'] = dados_scanc['Centro'].apply(lambda x: cnpj_fic(x, 'CNPJ FIC'))


##### ----- LISTA ARMAZENA INFORMAÇÕES DOS E-MAILS ----- #####
emails_para_enviar = []


##### ----- REGRAS E-MAILS ----- #####
base_regras = os.path.join(pasta_scanc, 'Base de dados - Ordem de entrega SCANC.xlsx')
dados_regras = pd.read_excel(base_regras)
dados_regras['Nome'] = dados_regras['Nome'].str.replace('/', '')
print(dados_regras['CNPJ'].unique())
print(len(dados_regras['CNPJ'].unique()))
#print(dados_regras.columns)

#### ----- DEFINIÇÃO REGRAS DE SEGMENTAÇÃO ----- #####
regras = {}
for index, row in dados_regras.iterrows():
    regras[row['CNPJ']] = {
        'SCANC': row['SCANC'].split(','),
        'Nome': row['Nome'],
        'CFOP': row['CFOP'].split(','),
        'Email': row['Email'],
        'Cc': row['Cc'].split(',')
    }


##### ----- FILTRO POR SEGMENTAÇÃO ----- #####
dados_scanc_filtro = dados_scanc[dados_scanc['CNPJ'].isin(regras.keys())]
print(dados_scanc.head())


##### ----- LAÇO FOR PARA SEPARAR OS CONGÊNERES SELECIONADOS ----- #####
congeneres = list(regras.keys())
print(congeneres)


congeneres_sem_dados = []


for congenere in congeneres:
    scanc = regras[congenere]['SCANC']
    cfop = regras[congenere]['CFOP']
    nome_congenere = regras[congenere]['Nome']

    filtro_congerene = dados_scanc_filtro[
        (dados_scanc_filtro['CNPJ'] == congenere) &
        (dados_scanc_filtro['SCANC.'].isin(scanc)) &
        (dados_scanc_filtro['CFOP'].isin(cfop))]

    print(f'\nCNPJ Congenere: {congenere}')
    #print(filtro_congerene.head())


######------ SALVANDO ARQUIVO EM EXCEL ------ ######
    if filtro_congerene.empty:
        print(f'\nNão foram encontrados dados para {nome_congenere} - {congenere}')
        congeneres_sem_dados.append({'CNPJ': congenere, 'Nome': nome_congenere})
    else:
        ultimos_num_cnpj = congenere[-7:]
        tabela_saida_com_dados = os.path.join(pasta_scanc, f'SCANC {mes_ano} - {nome_congenere} {ultimos_num_cnpj}.xlsx')
        arquivo_com_dados = pd.ExcelWriter(tabela_saida_com_dados, engine='xlsxwriter')

        with pd.ExcelWriter(tabela_saida_com_dados, engine='xlsxwriter') as arquivo_com_dados:
            filtro_congerene.to_excel(arquivo_com_dados, index=False)
            print(f'Arquivo gerado para o congênere {nome_congenere}: {tabela_saida_com_dados}')


######------ DEFININDO MENSAGEM CORPO DO E-MAIL COM BASE NAS OPERAÇÕES ------ ######
        if 'Entrada' in filtro_congerene['Operação'].values and 'Saída' in filtro_congerene['Operação'].values:
            mensagem = (f'Prezados,\n\nA ROYAL FIC informa que adquiriu e comercializou combustível no período {mes_ano} com a {nome_congenere}. '
                         'Para evitarmos a quebra de Congêneres prevista no Convênio 110/2007, peço a gentileza de nos '
                         'avisar sua primeira entrega para verificarmos se ocorrerá looping.\n\n'
                         'Atenciosamente.')

        elif 'Entrada' in filtro_congerene['Operação'].values:
            mensagem = (f'Prezados,\n\nA ROYAL FIC informa que adquiriu combustível no período {mes_ano} da {nome_congenere}. '
                         'Para evitarmos a quebra de Congêneres prevista no Convênio 110/2007, peço a gentileza de nos '
                         'aguardar.\n\n'
                         'Atenciosamente.')

        elif 'Saída' in filtro_congerene['Operação'].values:
            mensagem = (f'Prezados,\n\nA ROYAL FIC informa que comercializou combustível no período {mes_ano} com a {nome_congenere}. '
                        'Para evitarmos a quebra de Congêneres prevista no Convênio 110/2007, peço a gentileza de nos '
                        'avisar assim que fizer a entrega dos seus anexos do SCANC.\n\n'
                        'Atenciosamente.')
        else:
            mensagem = 'Este e-mail contém informações sobre operações diversas.'

        emails_para_enviar.append({'subject': f'SCANC {mes_ano} - {nome_congenere}',
                                   'body': mensagem,
                                   'to_email': regras[congenere]['Email'],
                                   'cc_email': regras[congenere]['Cc'],
                                   'attachments': [tabela_saida_com_dados]})

        print(emails_para_enviar)


##### ----- FUNÇÃO PARA GERAR RELATÓRIO DE CONGÊNERES SEM DADOS ----- #####
if congeneres_sem_dados:
    tabela_saida_sem_dados = os.path.join(pasta_scanc, f'SCANC {mes_ano} - RELATÓRIO CONGENERÊS SEM DADOS.xlsx')
    dados_congeneres_sem_dados = pd.DataFrame(congeneres_sem_dados)
    print(dados_congeneres_sem_dados.head(60))
    arquivo_sem_dados = pd.ExcelWriter(tabela_saida_sem_dados, engine='xlsxwriter')

    with pd.ExcelWriter(tabela_saida_sem_dados, engine='xlsxwriter') as arquivo_sem_dados:
        dados_congeneres_sem_dados.to_excel(arquivo_sem_dados, index=False)
        print(f'Relatório gerado: {tabela_saida_sem_dados}')
else:
    print('Todos os congenêres possuem dados no período informado.')


##### ----- CONFIGURAÇÕES GERAIS ----- #####
def send_email(subject, body, to_email, cc_email=None, attachments=None):
##### ----- INICIALIZAR OUTLOOKL ----- #####
    outlook = client.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0)

##### ----- CONFIGURAR E-MAIL ----- #####
    mail.Subject = subject
    mail.Body = body
    mail.To = to_email
    mail.CC = '; '.join(cc_email)

##### ----- ADICIONAR ANEXOS (SE HOUVER) ----- #####
    if attachments:
        for attachment in attachments:
            if os.path.exists(attachment):
                anexo = os.path.join(pasta_scanc, attachment)
                print(f'Adicionando anexo: {anexo}')
                mail.Attachments.Add(anexo)
            else:
                print(f'Erro: O arquivo {attachment} não pode ser encontrado.')

##### ----- ENVIAR E-MAIL ----- #####
    try:
        mail.Send()
        print('E-mail enviado com sucesso!')
    except Exception as e:
        print(f'Falha ao enviar e-mail: {e}')


##### ----- ENVIAR E-MAIL ----- #####
enviar_email = input('\nDeseja enviar o e-mail? (S/N): ')

if enviar_email.upper() == 'S':
    for email_info in emails_para_enviar:
        send_email(email_info['subject'], email_info['body'], email_info['to_email'], email_info['cc_email'], email_info['attachments'])
else:
    print('Dados gerados, e-mail não enviado.')


##### ----- FUNÇÃO PARA GERAR RELATÓRIO DE CONGÊNERES SEM CADASTRO ----- #####


final = time.time()
tempo_seg = final - inicio
tempo_min = tempo_seg / 60
print(f'\nTempo em segundos: {tempo_seg:.3f}')
print(f'Tempo em minutos: {tempo_min:.3f}')
