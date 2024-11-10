#Este também deu certo, agora inclui soma do total de requisições com turmas cancelasdas/concluidas
#proximo passo inclui a soma das turmas com data de entrega em atraso 30 dias

print('Executando.............Romaneio Filial x Status das Turmas')
# Importação de bibliotecas
import pandas as pd
import os
import datetime
import numpy as np

#import warnings
#warnings.filterwarnings("ignore", category=UserWarning)

# Leitura dos arquivos Excel
df1 = pd.read_excel(r"C:\Users\sar8577\Documents\AA_DATA FRAME_PYTHON\df_romaneio_filial_mxm.xlsx")
df2 = pd.read_excel(r"C:\Users\sar8577\Documents\AA_DATA FRAME_PYTHON\df_listagemTurmas.xlsx")

# Imprimindo informações sobre df1 e df2
print(f"Romaneio possui {df1.shape[0]} linhas e {df1.shape[1]} colunas.")
print(f"Listagem Turmas possui {df2.shape[0]} linhas e {df2.shape[1]} colunas.")

# Pré-processamento de df2
df2 = df2.iloc[21:]  # Excluir as primeiras 21 linhas
df2 = df2.sort_values(by=['Unnamed: 0'], ascending=False)  # Ordenar do maior para o menor

# Criar nova Coluna e extrair texto
df1['Num_Projeto_Romaneio'] = df1['PROJETO'].str[:13]
df2['Num_Projeto_Listagem_Turmas'] = df2['Unnamed: 0'].str[:13]

# Agrupar df2 por 'Num_Projeto_Listagem_Turmas' e obter as informações desejadas
df2_grouped = df2.groupby('Num_Projeto_Listagem_Turmas').agg({'Unnamed: 21': 'first',
                                                              'Unnamed: 28': 'first',
                                                              'Unnamed: 31': 'first'})

# Mapear as informações de df2 para df1
df1['Status da Turma'] = df1['Num_Projeto_Romaneio'].map(df2_grouped['Unnamed: 21'])
df1['Inicio da Turma'] = df1['Num_Projeto_Romaneio'].map(df2_grouped['Unnamed: 28'])
df1['Termino da Turma'] = df1['Num_Projeto_Romaneio'].map(df2_grouped['Unnamed: 31'])

# Criar condicional para concatenar 'Estágio' ao 'Status da Turma' quando 'TIPO DE SOLICITAÇÃO' for 'EST'
df1['Status da Turma'] = df1.apply(
    lambda row: f"Estágio - {row['Status da Turma']}" if row['TIPO DE SOLICITAÇÃO'] == 'EST' else row['Status da Turma'],
    axis=1
)

# Assumindo que 'DATA DE ENTREGA' está no formato datetime
df1['DATA DE ENTREGA'] = pd.to_datetime(df1['DATA DE ENTREGA'])

# Converter a data de hoje para datetime
hoje = pd.to_datetime(datetime.date.today())

# Identificar o índice da coluna "DATA DE ENTREGA"
indice_data_entrega = df1.columns.get_loc('DATA DE ENTREGA')

# Criar as novas colunas com as condições
df1['STATUS DA DATA DE ENTREGA'] = np.where(df1['DATA DE ENTREGA'] < hoje, 'Entrega em Atraso', 'Dentro do Prazo')
df1['SUPERIOR A 30 DIAS'] = np.where(df1['DATA DE ENTREGA'] < hoje - pd.Timedelta(days=30),
                                     'Data de Entrega em Atraso Superior a 30 dias', 'Dentro do Prazo ou Menor que 30 dias')

# Criar a coluna 'VALOR TOTAL' como o produto de 'VALOR UNIT.' e 'QTDE SOLICITADA NA FILIAL'
df1['VALOR TOTAL'] = df1['VALOR UNIT.'] * df1['QTDE SOLICITADA NA FILIAL']

# Converter a coluna 'CODFILIAL' para string para garantir a comparação correta
df1['CODFILIAL'] = df1['CODFILIAL'].astype(str)

# Excluir as linhas onde CODFILIAL é igual a "1"
df1 = df1.query('CODFILIAL != "1"')

# Ajustar a ordem das colunas para incluir 'VALOR TOTAL'
nova_ordem = ['Num_Projeto_Romaneio', 'Status da Turma', 'Inicio da Turma', 'Termino da Turma', 'CODFILIAL', 'FILIAL',
              'Nº DA REQ.', 'SEQ.', 'TIPO DE SOLICITAÇÃO', 'C CUSTO', 'PROJETO', 'STATUS DO PROJETO', 'DATA DE EMISSÃO',
              'DATA DE ENTREGA', 'STATUS DA DATA DE ENTREGA', 'SUPERIOR A 30 DIAS', 'MATRÍCULA DO REQ.', 'STATUS DA REQ.',
              'OBS.', 'JUSTIFICATIVA', 'GRUPO DE COTAÇÃO', 'CÓD DO ITEM', 'DESCRIÇÃO', 'UNID.', 'VALOR UNIT.',
              'QTDE SOLICITADA NA FILIAL', 'VALOR TOTAL', 'SALDO DE ESTOQUE DA FILIAL', 'QTDE DISPONÍVEL NA FILIAL',
              'ESTOQUE SEPARAR', 'OPERAÇÃO', 'ESTOQUE EM TRÂNSITO P/ A FILIAL']

# Reordenar o DataFrame
df1 = df1[nova_ordem]

# Converter todas as colunas de texto para letras maiúsculas
df1 = df1.apply(lambda col: col.str.upper() if col.dtype == 'object' else col)


# Ordenar do maior para o menor número turmas
df1 = df1.sort_values(by=['Status da Turma'], ascending=False)

#################### Consolidação por Filial x Turmas Canceladas&Concluídas ################
# Filtrar as turmas canceladas e concluídas
df_filtrado = df1[df1['Status da Turma'].isin(['TURMA CANCELADA', 'TURMA CONCLUIDA'])]

# Agrupar por filial e somar o valor total
df_agrupado = df_filtrado.groupby('CODFILIAL')['VALOR TOTAL'].sum().reset_index()

# Mergiar com o DataFrame original para obter o nome da filial
df_final = pd.merge(df_agrupado, df1[['CODFILIAL', 'FILIAL']].drop_duplicates(), on='CODFILIAL')

# Adicionando a nova coluna "TOTAL DE REQ. CANCELADA/CONCLUIDAS"
df_reqs_filtradas = df1[df1['Status da Turma'].isin(['TURMA CANCELADA', 'TURMA CONCLUIDA'])]

# Contar o número de requisições únicas por filial
df_contagem_reqs = df_reqs_filtradas.groupby('CODFILIAL')['Nº DA REQ.'].nunique().reset_index()

# Renomear a coluna para "TOTAL DE REQ. CANCELADA/CONCLUIDAS"
df_contagem_reqs = df_contagem_reqs.rename(columns={'Nº DA REQ.': 'TOTAL DE REQ. CANCELADAS/CONCLUIDAS'})

# Fazer o merge com df_final para adicionar a nova coluna
df_final = pd.merge(df_final, df_contagem_reqs, on='CODFILIAL', how='left')

# Preencher os valores ausentes com 0, caso existam filiais sem requisições canceladas ou concluídas
df_final['TOTAL DE REQ. CANCELADAS/CONCLUIDAS'] = df_final['TOTAL DE REQ. CANCELADAS/CONCLUIDAS'].fillna(0).astype(int)

# Reordenando as colunas
nova_ordem = ['CODFILIAL', 'FILIAL', 'TOTAL DE REQ. CANCELADAS/CONCLUIDAS', 'VALOR TOTAL']
df_final = df_final[nova_ordem]

# Ordenar do maior para o menor valor total
df_final = df_final.sort_values(by=['VALOR TOTAL'], ascending=False)

# Imprimir o DataFrame final
#print(df_final)

# Converter os nomes das colunas para letras maiúsculas############# Tive que colocar aqui porque o código já estava montado###
df1.columns = df1.columns.str.upper()

########INICIAR EXPORTAÇÃO ###################
# Criar um formato de data e hora personalizado
formato_data_hora = "%Y%m%d_%H%M%S"
data_hora_atual = datetime.datetime.now().strftime(formato_data_hora)

# Construir o nome do arquivo com a data e hora
nome_arquivo = f"Romaneio_Filial_MXM_Status_da_Turma_{data_hora_atual}.xlsx"

caminho_saida = "C:/Users/sar8577/Downloads/"

# Exportar df1 e df_final para o arquivo Excel
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

############################################################################



''''''''''''''''

#####OUTROS GRAFICOS #####
from dash import Dash, dcc, html
from dash.dependencies import Input, Output

# Inicializando o app Dash
app = Dash(__name__)

# Layout do app
app.layout = html.Div([
    dcc.Dropdown(
        id='dropdown-filial',
        options=[{'label': filial, 'value': filial} for filial in df_final['FILIAL'].unique()],
        value=df_final['FILIAL'].unique()[0],  # Valor padrão
        multi=True  # Permite selecionar múltiplas filiais
    ),
    dcc.Graph(id='grafico-barras')
])

# Callback para atualizar o gráfico dinamicamente com base no dropdown
@app.callback(
    Output('grafico-barras', 'figure'),
    Input('dropdown-filial', 'value')
)
def update_graph(filiais_selecionadas):
    df_filtrado = df_final[df_final['FILIAL'].isin(filiais_selecionadas)]
    fig = px.bar(df_filtrado, x='FILIAL', y='VALOR TOTAL', text='TOTAL DE REQ. CANCELADAS/CONCLUIDAS')
    return fig

if __name__ == '__main__':
    app.run_server(debug=True)
    
print('grafico 1 Dash ok')

########################
####GRAFICO 3####################
import streamlit as st
import plotly.express as px

def criar_grafico(df):
    if df.empty:
        st.write("Não há dados para a seleção feita.")
        return

    fig = px.bar(
        df,
        x='FILIAL',
        y='VALOR TOTAL',
        text='TOTAL DE REQ. CANCELADAS/CONCLUIDAS',
        title='Valor Total por Filial para Requisições Canceladas/Concluídas',
        labels={'VALOR TOTAL': 'Valor Total (R$)'},
        color_discrete_sequence=px.colors.qualitative.Pastel
    )

    fig.update_layout(
        yaxis=dict(tickprefix='R$ ')
    )

    st.plotly_chart(fig)

# Título da aplicação
st.title('Análise de Requisições Canceladas/Concluídas por Filial')

# Filtro de filiais
filiais_selecionadas = st.multiselect('Selecione as filiais:', df_final['FILIAL'].unique())

# Filtrar o DataFrame com base na seleção
if filiais_selecionadas:
    df_filtrado = df_final[df_final['FILIAL'].isin(filiais_selecionadas)]
else:
    df_filtrado = df_final

# Chamar a função para criar o gráfico
criar_grafico(df_filtrado)



####GRAFICO 3 ####################
'''''''''
 
