import os
import pandas as pd

# Obtendo o caminho absoluto do diretório atual
diretorio_atual = os.path.abspath('.')

# Nomes dos arquivos de entrada
nome_arquivo1 = 'antiga.xlsx'
nome_arquivo2 = 'nova.xlsx'

# Caminhos completos para os arquivos de entrada e saída
caminho_arquivo1 = os.path.join(diretorio_atual, nome_arquivo1)
caminho_arquivo2 = os.path.join(diretorio_atual, nome_arquivo2)
caminho_arquivo_final = os.path.join(diretorio_atual, 'arquivo_final.xlsx')

# Carregar os dados dos arquivos Excel
arquivo1 = pd.read_excel(caminho_arquivo1)
arquivo2 = pd.read_excel(caminho_arquivo2)

# Definir a coluna "descricao" como índice para facilitar a atualização
arquivo1.set_index('descricao', inplace=True)
arquivo2.set_index('descricao', inplace=True)

# Atualizar os dados do arquivo1 com base no arquivo2 usando o método update
arquivo1.update(arquivo2)

# Redefinir a coluna "descricao" como coluna novamente
arquivo1.reset_index(inplace=True)

# Ordenar as colunas na ordem correta (considerando 'produto' primeiro e 'descricao' em seguida)
colunas_ordenadas = ['produto', 'descricao'] + [coluna for coluna in arquivo1.columns if coluna not in ['produto', 'descricao']]
arquivo1 = arquivo1[colunas_ordenadas]

# Salvar o arquivo final em um novo arquivo Excel
arquivo1.to_excel(caminho_arquivo_final, index=False)
