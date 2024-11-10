#Priorizar "Bloqueada Matricula_volux"
print('Inicializando Romaneio Volux ')
import pandas as pd
import os
import datetime
import numpy as np


# Leitura dos arquivos Excel (substitua pelos seus caminhos)
df1 = pd.read_excel(r"C:\Users\sar8577\Documents\AA_DATA FRAME_PYTHON\df_romaneio volux.xlsx")
df2 = pd.read_excel(r"C:\Users\sar8577\Documents\AA_DATA FRAME_PYTHON\df_listagemTurmas.xlsx")
df3=pd.read_excel(r"C:\Users\sar8577\Documents\AA_DATA FRAME_PYTHON\df_estoque_CustoMedio_Escola_Barbacena.xlsx") 

# Imprimindo informações sobre df1 e df2
print(f"Romaneio Volux possui {df1.shape[0]} linhas e {df1.shape[1]} colunas.")
print(f"Listagem Turmas possui {df2.shape[0]} linhas e {df2.shape[1]} colunas.")

#df2 excluir as linhas do 1 até ao 21
df2 = df2.iloc[21:]
#ordenar do maior para o menor número turmas
df2 = df2.sort_values(by=['Unnamed: 0'], ascending=False)

#Criar nova Coluna e extrair texto
df1['Num_Projeto_Romaneio'] = df1['Descricao'].str[:13]
df2['Num_Projeto_Listagem_Turmas'] = df2['Unnamed: 0'].str[:13]

# Função para priorizar "Bloqueada Matricula"
def prioriza_bloqueada(x):
    """
    Prioriza o status 'Bloqueada Matricula' em um grupo.

    Args:
        x: Série com os valores do grupo.

    Returns:
        str: 'Bloqueada Matricula' se estiver presente, caso contrário, o último valor.
    """
    if 'Bloqueada Matricula' in x.values:
        return 'Bloqueada Matricula'
    else:
        return x.iloc[-1]  # Mantém a lógica atual para outros casos

# Agrupando e aplicando a função de priorização, incluindo todas as colunas de interesse
df2_grouped = df2.groupby('Num_Projeto_Listagem_Turmas').agg({'Unnamed: 21': prioriza_bloqueada,
                                                         'Unnamed: 28': 'last',
                                                         'Unnamed: 31': 'last'})

# Mapeamento para df1
df1['Status da Turma'] = df1['Num_Projeto_Romaneio'].map(df2_grouped['Unnamed: 21'])
df1['Inicio da Turma'] = df1['Num_Projeto_Romaneio'].map(df2_grouped['Unnamed: 28'])
df1['Termino da Turma'] = df1['Num_Projeto_Romaneio'].map(df2_grouped['Unnamed: 31'])

# Assumindo que 'DATA DE ENTREGA' está no formato datetime
df1['Data_Desejada'] = pd.to_datetime(df1['Data_Desejada'])

# Converter a data de hoje para datetime
hoje = pd.to_datetime(datetime.date.today())


# Criar as novas colunas com as condições
df1['STATUS DA DATA DE ENTREGA'] = np.where(df1['Data_Desejada'] < hoje, 'Entrega em Atraso', 'Dentro do Prazo')
df1['SUPERIOR A 30 DIAS'] = np.where(df1['Data_Desejada'] < hoje - pd.Timedelta(days=30),
                                     'Data de Entrega em Atraso Superior a 30 dias', 'Dentro do Prazo ou Menor que 30 dias')



nova_ordem= ['Sequencial',	'data',	'Cod_Almoxarifado',	'almox','Cod_CentroResultado', 'Descricao','Status da Turma',	
'Inicio da Turma',	'Termino da Turma',	'Cod_Funcionario',	'Nome',	'Autorizado',	'Situacao',	'STATUS DA DATA DE ENTREGA','SUPERIOR A 30 DIAS','Data_Desejada',	
'Qtde_Pedida',	'Qtde_Recebida',	'SaldoRequisicao',	'Qtde_Unidade_Estoque',	'Num_ItemRequisicao',	'item',	
'cod_secundario',	'desitem',	'Unidade',	'Qtde_Sub_Em_Unidade',	'Subunidade', 'Operacao',	'estoqueSeparar',	'estoqueAjustado',]

df1['data'] = pd.to_datetime(df1['data']).dt.strftime('%d/%m/%Y') #formatar data

df1['Data_Desejada'] = pd.to_datetime(df1['Data_Desejada']).dt.strftime('%d/%m/%Y')


df1=df1[nova_ordem]

# Renomeando a coluna 'Coluna1' para 'NovaColuna'
df1 = df1.rename(columns={'data': 'DATA DE EMISSÃO','Data_Desejada': 'DATA DE ENTREGA','Sequencial': 'NÚMERO DA REQ.','Cod_CentroResultado': 'CÓD. REDUZIDO','Descricao':'PROJETO' })

# Converter todas as colunas de texto para letras maiúsculas
df1 = df1.apply(lambda col: col.str.upper() if col.dtype == 'object' else col)
# Converter os nomes das colunas para letras maiúsculas############# Tive que colocar aqui porque o código já estava montado###
df1.columns = df1.columns.str.upper()








# Criar um formato de data e hora personalizado
formato_data_hora = "%Y%m%d_%H%M%S"  # Exemplo: 20231120_153025 (AAAA-MM-DD_HH-MM-SS)
data_hora_atual = datetime.datetime.now().strftime(formato_data_hora)

# Construir o nome do arquivo com a data e hora
nome_arquivo = f"Romaneio_Volux_Requisições_x_Status_da_Turma_{data_hora_atual}.xlsx"

# Caminho completo para o arquivo
caminho_saida = "C:/Users/sar8577/Downloads/"

# Exportar df1 para o arquivo Excel
df1.to_excel(caminho_saida + nome_arquivo, index=False)

print(df1.head(10))

print('Exportado com Sucesso')