# Mercadinho-de-Bairro
Disponibilizo um arquivo .xlsx e um .py para a simulação de um sistema de controle; .xlsx é um arquivo Excel simulando um sistema para um comercio básico/intermediário, já o .py é um script para preencher tabelas do sistema com dados fictícios.

Mercadinho do Bairro - Projeto de Excel + Script em python

Projeto pessoal desenvolvido para aplicar na prática os conteúdos dos cursos de Excel Iniciante, Intermediário e Avançado da Fundação Bradesco.

A ideia foi simular a gestão completa de um pequeno mercado de bairro: estoque, vendas, funcionários e financeiro, tudo dentro de uma única planilha, com um dashboard para visualizar os resultados.

Arquivos
mercadinho.xlsx	-> Planilha final, pronta para uso (tabelas Vendas e Estoque já preenchidos).
gerar_vendas_estoque.py	-> Script Python usado para popular as abas Vendas e Estoque com dados aleatórios (usado por causa do grande volume de linhas).

Estrutura da planilha
Dashboard ; KPIs (receita, lucro líquido, margem, ticket médio, funcionários, produtos) e gráficos: receita/despesas/lucro por mês, vendas por categoria e top 5 produtos mais vendidos.
Produtos ; Catálogo com código, categoria, preço de custo/venda, margem, validade, estoque mínimo e estoque atual (calculado automaticamente).
Funcionarios ; Dados cadastrais: nome, idade, telefone, cargo, data de admissão, salário e status.
Financas ; Demonstrativo mês a mês: receita, custo dos produtos, lucro bruto, folha de pagamento, aluguel, água, energia, despesas totais, lucro líquido e margem líquida.
Vendas ; Histórico de vendas geradas pelo script Python.
Estoque ; Movimentações de entrada (compras/reposição) e saída (vendas), usadas para calcular o estoque atual de cada produto.

Todos os valores da planilha são calculados por fórmulas (não estão fixos), então mudar qualquer dado de origem atualiza o restante automaticamente.

Como usar o script Python
O script gerar_vendas_estoque.py abre a planilha, lê o catálogo de produtos e gera:

Vendas aleatórias ao longo de um ano (quantidade controlada pela constante QUANTIDADE_DE_VENDAS);
Entradas de estoque suficientes para cobrir o que foi vendido, evitando estoque negativo.

Para gerar um novo conjunto de dados:

bash
pip install openpyxl
python gerar_vendas_estoque.py

O script sobrescreve as abas Vendas e Estoque da planilha indicada na constante ARQUIVO_PLANILHA, no início do código.

