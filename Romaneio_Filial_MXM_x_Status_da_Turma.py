print('Executando.............Romaneio Filial x Status das Turmas')
# Importação de bibliotecas
import pandas as pd
import os
import datetime
import numpy as np

#import warnings
#warnings.filterwarnings("ignore", category=UserWarning)

df1 = pd.read_excel(r"C:.xlsx")
df2 = pd.read_excel(r"C:.xlsx")

print(f"Romaneio possui {df1.shape[0]} linhas e {df1.shape[1]} colunas.")
print(f"Listagem Turmas possui {df2.shape[0]} linhas e {df2.shape[1]} colunas.")


df2 = df2.iloc[21:]  # Excluir as primeiras 21 linhas
df2 = df2.sort_values(by=['Unnamed: 0'], ascending=False)  # Ordenar do maior para o menor


df1['Num_Projeto_Romaneio'] = df1['PROJETO'].str[:13]
df2['Num_Projeto_Listagem_Turmas'] = df2['Unnamed: 0'].str[:13]


df2_grouped = df2.groupby('Num_Projeto_Listagem_Turmas').agg({'Unnamed: 21': 'first',
                                                              'Unnamed: 28': 'first',
                                                              'Unnamed: 31': 'first'})


df1['Status da Turma'] = df1['Num_Projeto_Romaneio'].map(df2_grouped['Unnamed: 21'])
df1['Inicio da Turma'] = df1['Num_Projeto_Romaneio'].map(df2_grouped['Unnamed: 28'])
df1['Termino da Turma'] = df1['Num_Projeto_Romaneio'].map(df2_grouped['Unnamed: 31'])


df1['Status da Turma'] = df1.apply(
    lambda row: f"Estágio - {row['Status da Turma']}" if row['TIPO DE SOLICITAÇÃO'] == 'EST' else row['Status da Turma'],
    axis=1
)


df1['DATA DE ENTREGA'] = pd.to_datetime(df1['DATA DE ENTREGA'])


hoje = pd.to_datetime(datetime.date.today())


indice_data_entrega = df1.columns.get_loc('DATA DE ENTREGA')


df1['STATUS DA DATA DE ENTREGA'] = np.where(df1['DATA DE ENTREGA'] < hoje, 'Entrega em Atraso', 'Dentro do Prazo')
df1['SUPERIOR A 30 DIAS'] = np.where(df1['DATA DE ENTREGA'] < hoje - pd.Timedelta(days=30),
                                     'Data de Entrega em Atraso Superior a 30 dias', 'Dentro do Prazo ou Menor que 30 dias')


df1['VALOR TOTAL'] = df1['VALOR UNIT.'] * df1['QTDE SOLICITADA NA FILIAL']


df1['CODFILIAL'] = df1['CODFILIAL'].astype(str)

df1 = df1.query('CODFILIAL != "1"')


nova_ordem = ['Num_Projeto_Romaneio', 'Status da Turma', 'Inicio da Turma', 'Termino da Turma', 'CODFILIAL', 'FILIAL',
              'Nº DA REQ.', 'SEQ.', 'TIPO DE SOLICITAÇÃO', 'C CUSTO', 'PROJETO', 'STATUS DO PROJETO', 'DATA DE EMISSÃO',
              'DATA DE ENTREGA', 'STATUS DA DATA DE ENTREGA', 'SUPERIOR A 30 DIAS', 'MATRÍCULA DO REQ.', 'STATUS DA REQ.',
              'OBS.', 'JUSTIFICATIVA', 'GRUPO DE COTAÇÃO', 'CÓD DO ITEM', 'DESCRIÇÃO', 'UNID.', 'VALOR UNIT.',
              'QTDE SOLICITADA NA FILIAL', 'VALOR TOTAL', 'SALDO DE ESTOQUE DA FILIAL', 'QTDE DISPONÍVEL NA FILIAL',
              'ESTOQUE SEPARAR', 'OPERAÇÃO', 'ESTOQUE EM TRÂNSITO P/ A FILIAL']


df1 = df1[nova_ordem]


df1 = df1.apply(lambda col: col.str.upper() if col.dtype == 'object' else col)



df1 = df1.sort_values(by=['Status da Turma'], ascending=False)

#################### Consolidação por Filial x Turmas Canceladas&Concluídas ################

df_filtrado = df1[df1['Status da Turma'].isin(['TURMA CANCELADA', 'TURMA CONCLUIDA'])]

df_agrupado = df_filtrado.groupby('CODFILIAL')['VALOR TOTAL'].sum().reset_index()


df_final = pd.merge(df_agrupado, df1[['CODFILIAL', 'FILIAL']].drop_duplicates(), on='CODFILIAL')


df_reqs_filtradas = df1[df1['Status da Turma'].isin(['TURMA CANCELADA', 'TURMA CONCLUIDA'])]

df_contagem_reqs = df_reqs_filtradas.groupby('CODFILIAL')['Nº DA REQ.'].nunique().reset_index()

df_contagem_reqs = df_contagem_reqs.rename(columns={'Nº DA REQ.': 'TOTAL DE REQ. CANCELADAS/CONCLUIDAS'})


df_final = pd.merge(df_final, df_contagem_reqs, on='CODFILIAL', how='left')


df_final['TOTAL DE REQ. CANCELADAS/CONCLUIDAS'] = df_final['TOTAL DE REQ. CANCELADAS/CONCLUIDAS'].fillna(0).astype(int)

nova_ordem = ['CODFILIAL', 'FILIAL', 'TOTAL DE REQ. CANCELADAS/CONCLUIDAS', 'VALOR TOTAL']
df_final = df_final[nova_ordem]


df_final = df_final.sort_values(by=['VALOR TOTAL'], ascending=False)

df1.columns = df1.columns.str.upper()

########INICIAR EXPORTAÇÃO ###################

formato_data_hora = "%Y%m%d_%H%M%S"
data_hora_atual = datetime.datetime.now().strftime(formato_data_hora)


nome_arquivo = f"Romaneio_Filial_MXM_Status_da_Turma_{data_hora_atual}.xlsx"

caminho_saida = "C:/Users/sar8577/Downloads/"


with pd.ExcelWriter(caminho_saida + nome_arquivo, engine='openpyxl') as writer:
    df1.to_excel(writer, sheet_name='Romaneio Filial', index=False)
    df_final.to_excel(writer, sheet_name='Resumo Req. Canc & Conc.', index=False)

# Exibir as primeiras 10 linhas de df1
print(df1.head(10))
print('Exportado Com Sucesso!!!')


#####GRAFICO################
import plotly.express as px

# Gráfico de barras interativo mostrando o valor total por filial
fig = px.bar(df_final, x='FILIAL', y='VALOR TOTAL', text='TOTAL DE REQ. CANCELADAS/CONCLUIDAS',
             title="Valor Total por Filial com Requisições Canceladas ou Concluídas",
             labels={'VALOR TOTAL': 'Valor Total (R$)', 'FILIAL': 'Filial'})

fig.update_traces(texttemplate='%{text}', textposition='outside')
fig.update_layout(uniformtext_minsize=8, uniformtext_mode='hide')

# Exibir o gráfico
fig.show()

print('grafico plotly ok') 
