# Analise de Velocidade de Jobs
Nessa aplicação Python é feita uma comparação de jobs, no caso, para ver se ficou mais lento ou mais rápido e qual foi a diferença de minutos entre os meses de Setembro e Outubro.

# Documentos
1. No documento 'app.py' está a importação das bibliotecas de leitura de planilhas, as contas matemáticas e as inserção dos dados nas planilhas.
2. Na planilha 'setembro.xlsx' estão todos os jobs e o tempo que eles foram executados.
3. Na planilha 'outubro.xlsx' estão todos os jobs e o tempo que eles foram executados.
4. No arquivo 'nomes_unicos.txt' estão apenas os nomes dos jobs das duas planilhas que no caso são os mesmos jobs.
5. Na planilha 'planilhajobs.xlsx' estão todos os jobs, a média em minutos de uma semana de setembro e outubro e qual foi a diferença.

# App.py
Utilizando a biblioteca 'openpyxl' é carregado um arquivo excel já existente 'planilhajobs.xlsx' e utilizando a biblioteca 'pandas' é carregado o arquivo 'nomes_unicos.txt' onde está o nome de todos os jobs, tanto do mês de Setembro, quanto do mês de Outubro.

Após isso, é carrego os arquivos Setembro.xlsx e Outubro.xlsx para poder salvar em listas as datas de uma semana do mês de setembro (dia 10 ao 17) e datas de uma semana do mês de outubro (dia 08 ao dia 15).

Com isso, utilizando a coluna RUNSTARTTIMESTAMP é possível salvar o valor desses dias que estão na coluna ELAPSEDRUNSECS.

No caso, é somado os minutos da semana e divido por 8, assim, é possível saber a média de tempo daquele job. Após isso, é divido por 60, para sabermos em minutos o tempo daquele job. Assim sendo tanto para setembro, quanto para outubro.
Por fim é feito a subtração da media da semana em minutos de outubro com a de setembro para saber a diferença. Se a diferença for negativa é por que o job foi executado de forma mais lenta em relação ao mês anterior, se não, foi executado com mais velocidade.

Após tudo isso, os dados são organizados e inseridos no arquivo 'planilhasjobs.xlsx', com as colunas do nome do job, a média de setembro, outubro e a diferença.
