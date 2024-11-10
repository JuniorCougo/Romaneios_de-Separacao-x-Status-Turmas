#Romaneio Supply_Aprovado#Este deu certo
#Este deu certo, segui como se fosse o excel ordenando do maior para o menor e depois agrupando o status com data inicio e final
import pandas as pd
import os
import datetime

import warnings
warnings.filterwarnings("ignore", category=UserWarning) 

# Leitura dos arquivos Excel
df1 = pd.read_excel(r"C:\Users\sar8577\Documents\AA_DATA FRAME_PYTHON\df_romaneio_supply_aprovado.xlsx")
#"C:\Users\sar8577\Documents\AA_DATA FRAME_PYTHON\df_romaneio_supply_aprovado.xlsx"
df2 = pd.read_excel(r"C:\Users\sar8577\Documents\AA_DATA FRAME_PYTHON\df_listagemTurmas.xlsx")#atentar a extensão do arquivi

#Imprimindo informações sobre df1 e df2
print(f"Romaneio possui {df1.shape[0]} linhas e {df1.shape[1]} colunas.")
print(f"Listagem Turmas possui {df2.shape[0]} linhas e {df2.shape[1]} colunas.")

#df2 excluir as linhas do 1 até ao 21
df2 = df2.iloc[21:]
#ordenar do maior para o menor número turmas
df2 = df2.sort_values(by=['Unnamed: 0'], ascending=False)

#Criar nova Coluna e extrair texto
df1['Num_Projeto_Romaneio'] = df1['PROJETO'].str[:13]
df2['Num_Projeto_Listagem_Turmas'] = df2['Unnamed: 0'].str[:13]

# ... (Fim)


# Agrupar df2 por 'Num_Projeto_Listagem_Turmas' e obter as informações desejadas
df2_grouped = df2.groupby('Num_Projeto_Listagem_Turmas').agg({'Unnamed: 21': 'first',
                                                         'Unnamed: 28': 'first',
                                                         'Unnamed: 31': 'first'})

# Mapear as informações de df2 para df1
df1['Status da Turma'] = df1['Num_Projeto_Romaneio'].map(df2_grouped['Unnamed: 21'])
df1['Inicio da Turma'] = df1['Num_Projeto_Romaneio'].map(df2_grouped['Unnamed: 28'])
df1['Termino da Turma'] = df1['Num_Projeto_Romaneio'].map(df2_grouped['Unnamed: 31'])

# Criar a coluna 'VALOR TOTAL' como o produto de 'VALOR UNIT.' e 'QTDE SOLICITADA NA FILIAL'
df1['VALOR TOTAL'] = df1['VALOR UNIT.'] * df1['QTDE SOLICITADA CD']


######################################
# Criar condicional para concatenar 'Estágio' ao 'Status da Turma' quando 'TIPO DE SOLICITAÇÃO' for 'EST'
df1['Status da Turma'] = df1.apply(
    lambda row: f"Estágio - {row['Status da Turma']}" if row['TIPO DE SOLICITAÇÃO'] == 'EST' else row['Status da Turma'],
    axis=1
)

# Ajustar a ordem das colunas para incluir 'VALOR TOTAL'
nova_ordem = ['tipoRomaneio', 'Num_Projeto_Romaneio',	'Status da Turma',	'Inicio da Turma',	'Termino da Turma',	'CODFILIAL',	
              'FILIAL',	'Nº DA REQ.',	'SEQ.',	'TIPO DE SOLICITAÇÃO',	'C CUSTO',	'PROJETO',	'STATUS DO PROJETO',	
              'DATA DE EMISSÃO',	'DATA DE ENTREGA',	'MATRÍCULA DO REQ.',	'STATUS DA REQ.',	'OBS.',	'JUSTIFICATIVA',	
              'GRUPO DE COTAÇÃO',	'CÓD DO ITEM',	'DESCRIÇÃO',	'UNID.',	'VALOR UNIT.',	'VALOR TOTAL',	'QTDE SOLICITADA CD',	'SALDO FÍSICO DO ESTOQUE CD',	'QTDE DISPONÍVEL NO CD',	
              'ESTOQUE SEPARAR',	'OPERAÇÃO',	'QTDE PENDENTE DE ENTREGA NO CD',	'PROJEÇÃO DE ATENDIMENTO_TRÂNSITO P/ O CD',
]

#'QTDE SOLICITADA NA FILIAL',	'SALDO DE ESTOQUE DA FILIAL',#

# Reordenar o DataFrame
df1 = df1[nova_ordem]


# Converter todas as colunas de texto para letras maiúsculas
df1 = df1.apply(lambda col: col.str.upper() if col.dtype == 'object' else col)

# Converter o cabeçalho (nomes das colunas) para letras maiúsculas
df1.columns = df1.columns.str.upper()

# Ordenar do maior para o menor número turmas
df1 = df1.sort_values(by=['STATUS DA TURMA'], ascending=False)

###########################################
# Criar um formato de data e hora personalizado
formato_data_hora = "%Y%m%d_%H%M%S"
data_hora_atual = datetime.datetime.now().strftime(formato_data_hora)

# Construir o nome do arquivo com a data e hora
nome_arquivo = f"Romaneio_Supply_Aprovado_x_Status_da_Turma_{data_hora_atual}.xlsx"

# Caminho completo para o arquivo
caminho_saida = "C:/Users/sar8577/Downloads/"

# Exportar df1 para o arquivo Excel
df1.to_excel(caminho_saida + nome_arquivo, index=False)


print(df1.head(10))
print('Romaneio Supply Aprovado Exportado Com Sucesso!!!')