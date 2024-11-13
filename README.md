# Subpedido Status Validation

Este repositório contém um script em Python desenvolvido para validar e classificar o status de subpedidos a partir de um relatório em Excel. 
O código processa dados de CNPJ, número da nota, EAN do produto e quantidade faturada para identificar duplicidades de subpedidos e determinar quais são corretos e incorretos.

## Funcionalidades

- **Validação de Duplicidades**: O código verifica duplicidades de subpedidos usando uma chave composta por CNPJ, número da nota, EAN do produto e quantidade faturada.
- **Classificação de Status**: Subpedidos com maior quantidade de EANs distintos são marcados como "Correto", enquanto os demais são marcados como "Incorreto".
- **Geração de Relatório**: O script gera um arquivo Excel com a coluna "Status", indicando os subpedidos corretos e incorretos, além de outros dados relevantes.

## Como Usar

1. Clone este repositório em sua máquina local:

   ```bash
   git clone https://github.com/yago-campos/validacao-status.git
Instale as dependências necessárias:
pip install pandas openpyxl

Coloque o arquivo Relatório.xlsx no mesmo diretório do script.

Execute o script:
python nome_do_script.py

O arquivo de saída, Resultados.xlsx, será gerado com a coluna "Status" e outras informações.

Estrutura do Arquivo Excel
O script espera um arquivo Excel (Relatório.xlsx) com as seguintes colunas:

CNPJ: CNPJ do fornecedor.
Número da nota: Número da nota fiscal.
EAN do produto: Código EAN do produto.
Quantidade Faturada: Quantidade faturada do item.
ID SUB PEDIDO: Identificador do subpedido.
Exemplo de Uso
Suponha que o arquivo Relatório.xlsx contenha dados de pedidos com múltiplos subpedidos. O script irá processar os dados, verificar duplicidades e gerar um relatório em que os subpedidos corretos (com mais EANs distintos) são marcados como "Correto", enquanto os outros serão marcados como "Incorreto".

Contribuição
Se você quiser contribuir para este projeto, siga as etapas:

Faça um fork deste repositório.
Crie uma branch para sua feature (git checkout -b minha-feature).
Faça as alterações necessárias e comite-as (git commit -am 'Adicionando nova feature').
Envie sua branch para o repositório remoto (git push origin minha-feature).
Abra um Pull Request.
