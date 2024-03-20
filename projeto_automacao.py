import pandas as pd
import pathlib
import os
import win32com.client as win32

# importar e tratar base de dados

df_email = pd.read_excel(r'Bases de Dados\Emails.xlsx')
df_vendas = pd.read_excel(r'Bases de Dados\Vendas.xlsx')
df_lojas = pd.read_csv(r'Bases de Dados\Lojas.csv', sep=";", encoding="latin1")
# Criar um arquivo/tabela para cada loja para cada loja
df_vendas=df_vendas.merge(df_lojas, on="ID Loja")
dicionario_lojas = {}
for loja in  df_lojas['Loja']:
    dicionario_lojas[loja] = df_vendas.loc[df_vendas['Loja']==loja,:]


#Calcular Indicador do ultimo dia
dia_indicador = df_vendas['Data'].max()
print(f'{dia_indicador.day}/{dia_indicador.month}/{dia_indicador.year}')

# Salva o Backup nas pastas
diretorio_backup = pathlib.Path(r'Backup Arquivos Lojas')
arquivos_diretorio = diretorio_backup.iterdir()
lista_nomes_arquivos = [arquivo.name for arquivo in arquivos_diretorio]

# Verificar se existe pasta da loja
for loja in dicionario_lojas:
    try:
        if loja not in lista_nomes_arquivos:
            nova_pasta = diretorio_backup/loja
            nova_pasta.mkdir()
    except FileExistsError:
        pass
    
# Nomear arquivo e salvar dentro das pastas
    nome_arquivo = f"{dia_indicador.day}_{dia_indicador.month}_{loja}.xlsx"
    local_salvar = diretorio_backup/loja/nome_arquivo
#salvar dentro da pasta
    try:
        dicionario_lojas[loja].to_excel(local_salvar)
    except:
        pass
print('Até aqui esta ok')
# Calcular os indicadores
for loja in dicionario_lojas:
    vendas_loja = dicionario_lojas[loja]
    vendas_loja_dia = vendas_loja.loc[vendas_loja['Data']==dia_indicador,:]
    # Faturamento
    faturamento_ano = vendas_loja['Valor Final'].sum()
    faturamento_dia = vendas_loja_dia['Valor Final'].sum()
    # Diversidade de Produtos
    qtde_produtos_ano =len(vendas_loja['Produto'].unique())
    qtde_produtos_dia = len(vendas_loja_dia['Produto'].unique())
    # TIcket Médio
    valor_venda = vendas_loja.groupby('Código Venda').sum()
    ticket_medio_ano = valor_venda['Valor Final'].mean()

    valor_venda_dia = vendas_loja_dia.groupby('Código Venda').sum()
    ticket_medio_dia = valor_venda_dia['Valor Final'].mean()

    # definição metas
    meta_faturamento_dia = 1000
    meta_faturamento_ano = 1650000
    meta_produtos_dia = 4
    meta_produtos_ano = 120
    meta_ticket_dia = 500
    meta_ticket_ano = 500
    # Enviar o OnePage # Enviar um E mail para a diretoria
    outlook = win32.Dispatch("outlook.application")

    nome = df_email.loc[df_email['Loja']==loja, 'Gerente'].values[0]   
    mail = outlook.CreateItem(0)
    mail.To = df_email.loc[df_email['Loja']==loja, 'E-mail2'].values[0]   
    # mail.CC = 'email@gmail.com'
    # mail.BCC = 'email@gmail.com'
    mail.Subject = f'OnePage Dia {dia_indicador.day} / {dia_indicador.month} - Loja {loja}'
    # mail.Body = 'texto do E-mail'#Ou 
    if faturamento_dia >= meta_faturamento_dia:
        cor_fat_dia = "green"
    else:
        cor_fat_dia = "red"

    if faturamento_ano >= meta_faturamento_ano:   
        cor_fat_ano ='green'
    else: "red"

    if qtde_produtos_dia >= meta_produtos_dia:
        cor_qtde_dia = "green"
    else:
        cor_qtde_dia = "red"

    if qtde_produtos_ano >= meta_produtos_ano:
        cor_qtde_ano = "green"
    else:
        cor_qtde_ano = "red"

    if ticket_medio_dia >= meta_ticket_dia:
        cor_ticket_dia = "green"
    else:
        cor_ticket_dia = "red"
    if ticket_medio_ano >= meta_ticket_ano:
        cor_ticket_ano = "green"
    else:
        cor_ticket_ano = "red"

    mail.HTMLBody = f'''<p>Olá, {nome}</p>
    <p>O resultado de ontem <strong>({dia_indicador.day}/{dia_indicador.month})<strong/> da <strong>Loja {loja}<strong/> foi: </p>

    <table>
        <tr>
            <th>Indicador</th>
            <th>Valor dia</th>
            <th>Meta dia</th>
            <th>Cenário dia</th>
        </tr>
        <tr>
            <td>Faturamento dia</td>
            <td style="text-align:center">R${faturamento_dia:.2f}</td>
            <td style="text-align:center">R${meta_faturamento_dia:.2f}</td>
            <td style="text-align:center"><font color = "{cor_fat_dia}">◙</font></td>
        </tr>
        <tr>
            <td>Quantidade Produtos dia</td>
            <td style="text-align:center">{qtde_produtos_dia}</td>
            <td style="text-align:center">{meta_produtos_dia}</td>
            <td style="text-align:center"><font color = "{cor_qtde_dia}">◙</font></td>
        </tr>
            <tr>
            <td>Ticket Medio dia</td>
            <td style="text-align:center">R${ticket_medio_dia:.2f}</td>
            <td style="text-align:center">R${meta_ticket_dia:.2f}</td>
            <td style="text-align:center"><font color = "{cor_ticket_dia}">◙</font></td>
        </tr>
    <table/>
    <br>
    <table>
        <tr>
            <th>Indicador</th>
            <th>Valor ano</th>
            <th>Meta ano</th>
            <th>Cenário ano</th>
        </tr>
        <tr>
            <td>Faturamento dia</td>
            <td style="text-align:center">R${faturamento_ano:.2f}</td>
            <td style="text-align:center">R${meta_faturamento_ano:.2f}</td>
            <td style="text-align:center"><font color = "{cor_fat_ano}">◙</font></td>
        </tr>
        <tr>
            <td>Quantidade Produtos dia</td>
            <td style="text-align:center">{qtde_produtos_ano}</td>
            <td style="text-align:center">{meta_produtos_ano}</td>
            <td style="text-align:center"><font color = "{cor_qtde_ano}">◙</font></td>
        </tr>
            <tr>
            <td>Ticket Medio dia</td>
            <td style="text-align:center">R${ticket_medio_ano:.2f}</td>
            <td style="text-align:center">R${meta_ticket_ano:.2f}</td>
            <td style="text-align:center"><font color = "{cor_ticket_ano}">◙</font></td>
        </tr>
    <table/>

    <p>Segue em anexo a planilha com os dados detalhados</p>

    <p>Qualquer duvida estou á disposição</p>

    <p>Att; Lucas</p>
    '''

    #Anexos (pode colocar aonde quiser)
    try:
        attachment = pathlib.Path.cwd()/diretorio_backup/loja/f"{dia_indicador.day}_{dia_indicador.month}_{loja}.xlsx"
        # attachment = r'C:\Users\lucas\OneDrive - 1 TIME AGENTE AUTONOMO DE INVESTIMENTOS LTDA\Documentos\Lucas Mazetto\Curso Python Impressionador\Projeto AutomaçãoIndicadores\Projeto AutomacaoIndicadores\{}\{}\{}'.format(diretorio_backup)
        mail.Attachments.Add(str(attachment))

        mail.Send()
    except:
        pass
print('fim do processo total')

# Ranking da melhor e pior loja do dia e do ano
# Ano
faturamento_loja_ano = df_vendas.groupby('Loja')[['Loja', 'Valor Final']].sum()
ranking_ano = faturamento_loja_ano.sort_values(by='Valor Final', ascending=False)


# dia
vendas_dia = df_vendas.loc[df_vendas['Data']== dia_indicador, :]
vendas_dia = vendas_dia.groupby('Loja')[['Loja', 'Valor Final']].sum()
ranking_dia = vendas_dia.sort_values(by='Valor Final', ascending=False)

# Salvando Arquivo para mandar para a diretoria
nome_arquivo_dia = r'Ranking {}_{}'.format(dia_indicador.day, dia_indicador.month) 
ranking_dia.to_excel(r'Backup Arquivos Lojas\{}.xlsx'.format(nome_arquivo_dia))

nome_arquivo_ano = r'Ranking {}'.format(dia_indicador.year)
ranking_ano.to_excel(r'Backup Arquivos Lojas\{}.xlsx'.format(nome_arquivo_ano))
# mandando e-mail para diretoria
outlook = win32.Dispatch("outlook.application")
mail = outlook.CreateItem(0)
mail.To = df_email.loc[df_email['Loja'] == "Diretoria", 'E-mail2'].values[0]
mail.Subject = f'Ranking dia {dia_indicador.day}/{dia_indicador.month}/{dia_indicador.year}'
mail.HTMLBody = f'''Olá, Prezados. Como estão?
Segue em anexo o ranking das lojas no ano e no dia atual

Melhor loja do dia em Faturamento: Loja {ranking_dia.index[0]} com Faturamento R${ranking_dia.iloc[0,0]:.2f}
Pior loja do dia em Faturamento: Loja {ranking_dia.index[-1]} com Faturamento R${ranking_dia.iloc[-1,0]:.2f}

Melhor loja do ano em Faturamento: Loja {ranking_ano.index[0]} com Faturamento R${ranking_ano.iloc[0,0]:.2f}
Pior loja do ano em Faturamento: Loja {ranking_ano.index[-1]} com Faturamento R${ranking_ano.iloc[-1,0]:.2f}  

Qualquer duvida estou a disposição

Att; Lucas Mazetto


'''

attachment = pathlib.Path.cwd()/diretorio_backup/f'{nome_arquivo_dia}.xlsx'        
mail.Attachments.Add(str(attachment))
attachment = pathlib.Path.cwd()/diretorio_backup/f'{nome_arquivo_ano}.xlsx'        
mail.Attachments.Add(str(attachment))

mail.Send()
print("e-mail da Diretoria enviado")

print('Fim do Projeto de Automação de Processos. Bora pra próxima')
