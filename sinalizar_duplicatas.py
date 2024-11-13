import pandas as pd

# Carregar o arquivo Excel e a aba relevante
file_path = 'Relatório.xlsx'

# Ler os dados da aba
df = pd.read_excel(file_path)

# Criar uma nova chave concatenada sem espaços nos traços
df['Chave'] = df['CNPJ'].astype(str) + '-' + df['Número da nota'].astype(str) + '-' + df['EAN do produto'].astype(str) + '-' + df['Quantidade Faturada'].astype(str)

# Contar a quantidade de EANs diferentes por SUB PEDIDO
ean_counts = df.groupby('ID SUB PEDIDO')['EAN do produto'].nunique().reset_index()

# Renomear a coluna para evitar conflitos
ean_counts = ean_counts.rename(columns={'EAN do produto': 'EAN_count'})

# Mesclar a contagem de EANs de volta com o DataFrame original
df = df.merge(ean_counts, on='ID SUB PEDIDO', how='left')

# Criar uma coluna 'Status' inicializando com 'OK'
df['Status'] = 'OK'

# Identificar os duplicados
duplicated_keys = df[df.duplicated(subset='Chave', keep=False)]

# Identificar o SUB PEDIDO com a maior quantidade de EANs diferentes para cada conjunto de chaves
max_ean_sub_pedidos = duplicated_keys.loc[duplicated_keys.groupby('Chave')['EAN_count'].idxmax()]

# Identificar os subpedidos duplicados que não são os com maior quantidade de EANs
menor_ean_sub_pedidos = duplicated_keys[~duplicated_keys.index.isin(max_ean_sub_pedidos.index)]

# Marcar os subpedidos corretos e incorretos com os números dos subpedidos
for chave in duplicated_keys['Chave'].unique():
    subpedidos = duplicated_keys[duplicated_keys['Chave'] == chave]
    if len(subpedidos) > 1:
        corretos = subpedidos.loc[subpedidos['EAN_count'].idxmax(), 'ID SUB PEDIDO']
        incorretos = subpedidos[subpedidos['ID SUB PEDIDO'] != corretos]['ID SUB PEDIDO'].tolist()
        
        df.loc[(df['Chave'] == chave) & (df['ID SUB PEDIDO'] == corretos), 'Status'] = f"Correto"
        df.loc[(df['Chave'] == chave) & (df['ID SUB PEDIDO'].isin(incorretos)), 'Status'] = df['ID SUB PEDIDO'].apply(lambda x: f"Incorreto" if x in incorretos else '')

# Reorganizar as colunas para colocar 'Chave' e 'Status' no início
cols = ['Chave', 'Status'] + [col for col in df.columns if col not in ['Chave', 'Status']]
df = df[cols]

df.drop(columns=['EAN_count'], inplace=True)

# Salvar o DataFrame ajustado no arquivo Excel
df.to_excel('Resultados.xlsx', index=False)