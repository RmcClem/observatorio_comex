Objetivo
A proposta era identificar empresas passíveis de comércio exterior no município de Cotia e publicar os resultados no site da Fatec Cotia. Utilizando apenas de ferramentas e dados gratuitos.

Como foi feito
Para isso foram feitas buscas das empresas que existem em Cotia, e foi encontrado arquivos de CNPJ em dados públicos da Receita Federal no site dados.gov.br. Todos em formato CSV ("comma separated values" - "valores separados por vírgulas").
Então decidiu-se utilizar três dos diferentes tipos de arquivos disponíveis: os arquivos de Empresas (9 arquivos), os arquivos de Estabelecimentos (9 arquivos), e o arquivo Cnae (1 arquivo). Após esta etapa foram filtrados quais colunas de dados seriam utilizados, com ajuda da consulta em metadados (site: "https://www.gov.br/receitafederal/dados/cnpj-metadados.pdf”).
Foi identificado qual coluna tratava-se do município e selecionado apenas o município de Cotia através de consulta ao arquivo ("Municipios", também disponível pela receita federal), utilizando o código de município para Cotia = "6361", então o código foi substituído para COTIA, também foram selecionadas apenas as empresas descritas com situação cadastral ATIVA = "02", conforme descrito no metadados, e o código foi substituído para ATIVA.
E para cada tipo de arquivo foram selecionadas as seguintes colunas:

Estabelecimentos:
"column_1", "column_2", "column_3", "column_4", "column_5", "column_6", "column_7","column_11", "column_12", "column_13", "column_14", "column_15", "column_16","column_17", "column_18", "column_19", "column_21", "column_22", "column_23", "column_28".
E renomeadas para:
"column_1": "cnpj_basico",
"column_2": "cnpj_ordem",
"column_3": "cnpj_dv",
"column_4": "id_matriz_filial",
"column_5": "nome_fantasia",
"column_6": "sit_cadastral",
"column_7": "dt_sit_cadastral",
"column_11": "dt_inicio_atv",
"column_12": "cnae_principal",
"column_13": "cnae_secundario",
"column_14": "tipo_logradouro",
"column_15": "logradouro",
"column_16": "numero",
"column_17": "complemento",
"column_18": "bairro",
"column_19": "cep",
"column_21": "municipio",
"column_22": "ddd",
"column_23": "telefone",
"column_28": "email"

Foram feitos tratamento nas colunas "email", transformando tudo para minúsculo, no formato de data nas colunas "dt_sit_cadastral", "dt_inicio_atv" de string para aaaa-mm-dd. também foi concatenado as colunas "ddd" e "telefone".

Empresas:
"column_1", "column_2", "column_4", "column_5", "column_6".
E renomeadas para:
"column_1": "cnpj_basico",
"column_2": "razao_social",
"column_4": "qualif_responsavel",
"column_5": "capital_social",
"column_6": "porte"

Para a coluna "porte", foi realizado o a substituição dos valores "00"="NÃO INFORMADO", "01"="MICRO EMPRESA", "03"="EMPRESA DE PEQUENO PORTE", "05"="DEMAIS", conforme metadados.

As duas tabelas (dataframes) foram mescladas usando join, tendo como identificador a coluna "cnpj_basico".

Cnaes:
Foram selecionadas todas as colunas e renomeadas para:
"column_1": "cod_cnae",
"column_2": "descricao"

Para a coluna "cod_cnae" foram completados com zero à esquerda os códigos que não continham 7 dígitos.

Também foi utilizado um arquivo auxiliar em formato excel, com a explicação detalhada dos códigos de cnae do site do ibge. Este arquivo foi colocado em uma pasta no google drive, e foram mantidas apenas as linhas onde a coluna "Classe" não possui valores ausentes, convertida para tipo string e entao removido "." e "-" e "/", então esses dados foram colocados em uma nova coluna chamada de "cnae_padronizado".
O objetivo desta operação é poder filtrar a classe da atividade realizada pela empresa, não apenas a atividade específica. Pois a classe contém 5 dígitos e é mais abrangente do que o código completo que contém 7 dígitos.

Referente ao arquivo no google drive, este é um arquivo auxiliar para que seja possível selecionar quais empresas serão consideradas comex sem definir no código python.
Este arquivo excel contém duas planilhas:
Filtro por nome: com uma coluna em que é usada para definir as palavras-chave que são usadas para buscar nas colunas razao_social e nome_fantasia, se alguma destas duas colunas conter alguma das palavras definidas, ela é considerada empresa comex.
Filtro por cnae: foi mantido as colunas classe sem alterar a formatação do código cnae, a coluna de descrição, e criada uma nova coluna nomeada “Seleção”, e nela criada uma regra para que possa ser preenchida apenas com “aceita” ou “rejeita” ou “ok”, foi definido por padrão todas as linhas preenchidas com “ok”. "aceita" são as atividades que são consideradas passíveis de COMEX. E "rejeita" é usado apenas para os casos que possam aparecer como COMEX por falso positivo, devido à busca por palavras-chave, por exemplo, a palavra “internacional” pode trazer empresas de COMEX, mas também pode retornar uma igreja ou restaurante. “ok” já definido por padrão são consideradas não comex, ao menos que sejam captadas pelo filtro por nome (palavra_chave).

Link do arquivo: para postagem no GitHub adaptei um arquivo local

- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
A etapa de análise dos CNAEs foi dividida em duas partes principais: o tratamento do CNAE Principal e o tratamento dos CNAEs Secundários. O objetivo foi contar a frequência de cada código e enriquecer a informação com suas respectivas descrições.

Para o cnae principal, foram mantidas apenas as linhas não nulas e não vazias. E a frequência de cada cnae foi armazenada na coluna "frequência", depois foi feita a mesclagem do código cnae usando a coluna "cod_cnae" como identificador.
Para a coluna de cnae secundário foram feitas a separação por "," pois nesta coluna algumas linhas continham mais de um cnae.
O que resultou em um dataframe com as seguintes colunas:
"código", "frequência", "descrição", "tipo", "classe"
No fim este dataframe não foi utilizado, porém está disponível para um possível uso futuro
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 

No código em python usando unicodedata NFD foi feita uma normalização das palavras-chave, e transformado tudo em minúsculo sem acento, para que alguma diferença de acentuação não afetasse o resultado. Depois identificado no arquivo em excel as atividades selecionadas com “aceita” e “rejeita” e utilizando o cnae_padronizado referente a elas para criar as variáveis do tipo lista cnaes_aceitos e cnaes_rejeitados.
Antes de filtrar, o codigo padroniza as colunas do df_final
Cria-se a coluna razao_social_norm aplicando a função normalizar_texto em razao_social.
Cria-se a coluna nome_fantasia_norm aplicando a função normalizar_texto em nome_fantasia.
Depois a filtragem é feita usando as duas regras, de aceitação e rejeição.

Aceitação: uma empresa é considerada aceita se cumprir qualquer uma das condições:
Filtro por nome: A versão normalizada da razão social ou nome fantasia contém alguma das palavras-chave definidas.
Filtro por cnae principal: O código cnae principal da empresa (os primeiros 5 dígitos, que representam a Classe) está na lista dos cnaes_aceitos.
Filtro por cnaes secundários: O texto dos cnaess secundários contém algum dos códigos definidos nos cnaes_aceitos (usa-se uma expressão regular especial para buscar o código correto dentro da lista de secundários).

Rejeição: uma empresa é considerada rejeitada se cumprir alguma das condições:
Filtro por cnae principal: O código cnae principal da empresa (os 5 dígitos da Classe) está na lista dos cnaes_rejeitados
Filtro por cnaes secundários: O texto dos cnaes secundários contém algum dos códigos definidos nos cnaes_rejeitados.
Após isso é criado um dataframe df_comex com as empresas aceitas, ou seja, consideradas passíveis de comex.

Em seguida é selecionado os CNPJ das empresas comex para criar um rótulo no df_final, uma coluna booleana é criada is_comex, se a empresas foi considerada comex = True, senão = False.

Após esta parte é feita uma preparação para o conteúdo que será usado na seção de metodologia do dashboard html. Esta etapa tem um propósito de recuperar os códigos CNAE e as palavras-chave que foram usados nas regras de Aceitação e Rejeição e prepará-los com suas descrições completas.
Filtragem: O df_filtro_cnae é filtrado novamente, selecionando apenas as linhas onde a coluna Seleção está marcada como 'ACEITA'.
Seleção de Colunas: As colunas Classe (o código CNAE de 5 dígitos) e Denominação (a descrição completa da atividade) são mantidas.
Criação da Lista: A lista lista_cnaes_aceitos é criada. Cada item é formatado como uma string que combina o código da classe e sua descrição completa (Ex: "4623-1 – Comércio atacadista de gado").
O processo se repete para rejeita.
Depois é resgatado a lista de palavras-chave usadas no arquivo excel e transformado para formato html.




Após isso é calculado o total de empresas ativas, o total de comex, o percentual de comex e a quantidade de empresas não comex. Também conta, das empresas consideradas comex, se foi por atividade principal ou por atividade secundária, e calcula o percentual de cada uma.
O código calcula a idade de cada empresa (tempo desde a dt_inicio_atv até a data atual), expressa em anos decimais, e usa uma função auxiliar (formatar_idade) para exibir o resultado em "Anos e Meses". Também identifica a empresa mais antiga e a mais recente.

- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
A seguir as empresas são classificadas em faixas de idade predefinidas (0-5, 5-10, 10-15, 15-20, 20-30 e 30+ anos).
A distribuição de empresas por faixa etária é apresentada, separadamente para o grupo Geral (Não-COMEX) e COMEX, mostrando em qual faixa a maior parte das empresas de cada grupo se concentra.
Mas no final esta parte acabou não sendo usada
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 

Também foi analisado o fluxo de novas aberturas de empresas ao longo do tempo.
Evolução Mensal (Últimos 12 Meses): O código conta o número de novas empresas abertas por mês, comparando o volume Total, COMEX e Não-COMEX
Crescimento Anual (Últimos 10 Anos): Repete a análise de volume de novas aberturas, mas por ano.
Após este passo é analisado o crescimento percentual do total de empresas no último ano, o código simula o total acumulado de empresas ativas a cada mês, calcula o crescimento total no período, mostrando o aumento percentual e a diferença absoluta da base de empresas.

Após calcular a idade de todas as empresas e estabelecer o perfil geral (média, mediana e extremos), o processo avança para análises comparativas e de fluxo de aberturas.
O sistema agrupa os dados pela coluna booleana is_comex e calcula, para cada grupo, a idade média e a idade mediana, além do desvio padrão (que mede a dispersão das idades). É calculada a diferença percentual entre a idade média do grupo COMEX e a do grupo Não-COMEX, resultando em uma conclusão direta (ex: "Empresas COMEX são em média X% mais antigas").

Após isso é feita a análise geográfica, com o objetivo de identificar os bairros com maior concentração de negócios em geral e, principalmente, aqueles que se destacam pela presença de empresas de comex. Para isso é feito um tratamento na coluna bairro, apenas para garantir a qualidade dos dados, removendo possíveis espações vazios antes e depois dos dados e transformado tudo para maiúsculo, então isso é salvo na coluna bairro_padronizado. Então é feito a contagem de empresas totais e comex em cada bairro.
As duas contagens são unidas em um único dataframe e calculado o percentual_comex para cada bairro ((total_comex / total_empresas) * 100).
Após isso é exibido a quantidade de empresas com a coluna bairro preenchida e vazia, e os bairros identificados, então é gerado três rankings: Top 20 bairros com mais empresas (geral) para identificar os 20 bairros com maior densidade empresarial. Top 20 bairro com mais empresas de comex, para identificar os 20 bairros com maior densidade de empresas comex. Top 15 bairros com maior concentração de comex percentual em que ordena os bairros pelo percentual_comex, mas aplica um filtro de mínimo de 10 empresas (total_empresas >= 10) para evitar distorções, onde um bairro com apenas 1 ou 2 empresas, se ambas forem COMEX, apareceria com 100% de concentração. Este ranking revela os bairros onde o perfil de COMEX é mais dominante.

Após este passo é iniciado a criação dos gráficos usando a biblioteca Plotly para transformar os dados de crescimento empresarial, calculados nas etapas anteriores, em gráficos interativos no formato HTML.

Para isso é criada uma pasta “graficos” para armazenar os arquivos html gerados. Para cada gráfico, os dados de contagem (anuais, mensais, acumulados) são extraídos dos dataframes do Polars (crescimento_anual, últimos_12_meses, df_crescimento_acumulado) e convertidos para o formato dataframe do pandas para a criação de gráficos com Plotly.
Entao são criados os graficos: “Crescimento Anual de Empresas (Últimos 10 Anos)”, “Evolução Mensal de Abertura de Empresas (Últimos 12 Meses)”, “Crescimento Acumulado Mensal (Últimos 12 Meses)”, em sequência os gráficos em base 100 (transforma uma série em índice base 100. O primeiro valor vira 100, os demais mostram variação percentual), e então criado os gráficos “Crescimento Anual de Empresas - Índice Base 100”, “Evolução Mensal de Abertura de Empresas - Índice Base 100”, “Crescimento Acumulado de Empresas - Índice Base 100”, e por fim os gráficos geográficos “Top 20 Bairros com Mais Empresas”, “Top 15 Bairros com Maior Concentração de Comex”, “Top 10 Bairros com Mais Empresas de Comex”. Todos os gráficos criados são salvos na pasta “gráficos” criada anteriormente.

Com todos os gráficos e variáveis já criadas é feito a preparação para criar o dashboard em html. Todas as métricas numéricas calculadas nas análises, e temporal e crescimento são consolidadas em um único dicionário chamado “stats”. Os critérios de filtragem (cnaes aceitos, rejeitados e palavras-chave) são reunidos nesse dicionário, utilizando a função formatar_lista_para_html para que sejam inseridos no relatório como listas formatadas em HTML.
O código define uma string longa que representa o esqueleto completo do dashboard em HTML, Ele utiliza placeholders ({stats['nome_da_métrica']}, {metodologia_html['nome_da_lista']}) que são preenchidos com os valores dos dicionários stats e metodologia_html, é feito desta forma para que quando chegue novos dados, estes novos dados ocupem estes espaços. Os gráficos gerados em html são incorporados no dashboard usando iframes. E após os gráficos é criada uma seção de metodologia para expor a metodologia usada e garantir transparência. Ao final o template html é preenchido e salvo como “dashboard.html” na pasta raiz, que conforme indicado no template irá importar o arquivo “style.css” que acompanha o projeto para realizar a formatação da página.

E para concluir é exportado para a pasta “outputs” no formato excel, o dataframe “df_final” que contém os dados originais tratados, já com a coluna “is_comex”, nomeada como “final aaaa-mm.xlsx”.

---------------------------------------------------------------------------------------------------------------------------
Como executar:

Requisitos: Possuir o python 3 instalado na máquina.
Conexão com internet para download dos dados.
Certifique-se de que o arquivo “processo.py” e o arquivo de estilos “style.css” estejam na mesma pasta.
O script já instala as bibliotecas necessárias sozinho.

Para executar, abra o terminal (ou Prompt de Comando/PowerShell) na pasta onde você salvou o projeto e execute o script:

python processo.py

Ao final da execução, os resultados gerados estarão disponíveis no arquivo “dashboard.html” salvo na pasta raiz

Acesse o arquivo html, ele irá abrir no navegador, pressione Ctrl + P, ou clique com botão direito do mouse e selecione Imprimir, em Destino, selecione “Salvar como PDF”, e em mais definições, selecione o tamanho do papel “A2”, desmarque “Cabeçalhos e rodapés” e “Gráficos de segundo plano” caso estejam selecionados. Então salve o arquivo.

No final, caso deseje, você pode excluir as seguintes pastas de arquivos gerados para armazenar os dados de CNPJ baixados:
“empresas/”, “estabelecimentos/”, “arquivos_unicos/”
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Caso opte por usar o arquivo processo.ipynb no formato jupyter notebook:

O projeto foi executado e testado utilizando a IDE VSCode.
Versão do Python: 3.10.11
Abra a IDE, selecione a pasta com os arquivos processo.ipynb e style.css.
Execute as células em ordem (de cima para baixo).

---------------------------------------------------------------------------------------------------------------------------

Observações:
Com velocidade de internet testada em aproximadamente 100 Mbps, o processo demorou cerca de 1 hora.

Foi criada uma conta google para colocar o arquivo auxiliar excel.

Este projeto foi desenvolvido como atividade de estágio na Fatec Cotia, e a publicação foi autorizada pelo responsável pelo projeto

---------------------------------------------------------------------------------------------------------------------------

Desenvolvedores do projeto:
Rafael Clem
Vinicius Francisco
