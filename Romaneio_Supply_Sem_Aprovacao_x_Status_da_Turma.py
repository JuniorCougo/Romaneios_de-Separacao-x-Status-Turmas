print('Executando...............Romaneio Supply Sem Aprovação x Status da Turma')
import pandas as pd
import os
import datetime

import warnings
warnings.filterwarnings("ignore", category=UserWarning) 

# Leitura dos arquivos Excel
df1 = pd.read_excel(r"C:.xlsx")
df2 = pd.read_excel(r"C:.xlsx")
print(f"Romaneio possui {df1.shape[0]} linhas e {df1.shape[1]} colunas.")
print(f"Listagem Turmas possui {df2.shape[0]} linhas e {df2.shape[1]} colunas.")

df2 = df2.iloc[21:]
#ordenar do maior para o menor número turmas
df2 = df2.sort_values(by=['Unnamed: 0'], ascending=False)

df1['Num_Projeto_Romaneio'] = df1['PROJETO'].str[:13]
df2['Num_Projeto_Listagem_Turmas'] = df2['Unnamed: 0'].str[:13]


df2_grouped = df2.groupby('Num_Projeto_Listagem_Turmas').agg({'Unnamed: 21': 'first',
                                                         'Unnamed: 28': 'first',
                                                         'Unnamed: 31': 'first'})

df1['Status da Turma'] = df1['Num_Projeto_Romaneio'].map(df2_grouped['Unnamed: 21'])
df1['Inicio da Turma'] = df1['Num_Projeto_Romaneio'].map(df2_grouped['Unnamed: 28'])
df1['Termino da Turma'] = df1['Num_Projeto_Romaneio'].map(df2_grouped['Unnamed: 31'])

df1['VALOR TOTAL'] = df1['VALOR UNIT.'] * df1['QTDE SOLICITADA CD']

df1['Status da Turma'] = df1.apply(
    lambda row: f"Estágio - {row['Status da Turma']}" if row['TIPO DE SOLICITAÇÃO'] == 'EST' else row['Status da Turma'],
    axis=1
)

nova_ordem = ['tipoRomaneio', 'Num_Projeto_Romaneio',	'Status da Turma',	'Inicio da Turma',	'Termino da Turma',	'CODFILIAL',	
              'FILIAL',	'Nº DA REQ.',	'SEQ.',	'TIPO DE SOLICITAÇÃO',	'C CUSTO',	'PROJETO',	'STATUS DO PROJETO',	
              'DATA DE EMISSÃO',	'DATA DE ENTREGA',	'MATRÍCULA DO REQ.',	'STATUS DA REQ.',	'OBS.',	'JUSTIFICATIVA',	
              'GRUPO DE COTAÇÃO', 'CÓD DO ITEM',	'DESCRIÇÃO',	'UNID.',	'VALOR UNIT.',	'VALOR TOTAL',	'QTDE SOLICITADA CD',	
              'SALDO FÍSICO DO ESTOQUE CD',	'QTDE DISPONÍVEL NO CD','ESTOQUE SEPARAR',	'OPERAÇÃO',	'QTDE PENDENTE DE ENTREGA NO CD','PROJEÇÃO DE ATENDIMENTO_TRÂNSITO P/ O CD',

             
]
#'QTDE SOLICITADA NA FILIAL','SALDO DE ESTOQUE DA FILIAL',

df1 = df1[nova_ordem]

df1 = df1.apply(lambda col: col.str.upper() if col.dtype == 'object' else col)
df1.columns = df1.columns.str.upper()

df1 = df1.sort_values(by=['STATUS DA TURMA'], ascending=False)

###########################################
formato_data_hora = "%Y%m%d_%H%M%S"
data_hora_atual = datetime.datetime.now().strftime(formato_data_hora)
nome_arquivo = f"Romaneio_Supply_Sem_Aprovacao_x_Status_da_Turma_{data_hora_atual}.xlsx"

caminho_saida = "C:Downloads/"

# Exportar df1 para o arquivo Excel
df1.to_excel(caminho_saida + nome_arquivo, index=False)


print(df1.head(10))
print('Romaneio Supply Em Aprovação Exportado Com Sucesso!!!')
