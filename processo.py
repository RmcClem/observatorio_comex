# Instalações

print("Iniciando...")

import subprocess
import sys

print("Instalando dependencias...")

def instalar_dependencias():
    pacotes = ['pandas', 'polars', 'requests', 'xlrd', 'openpyxl', 'plotly', 'xlsxwriter']
    for pacote in pacotes:
        subprocess.check_call([sys.executable, "-m", "pip", "install", pacote])
        
instalar_dependencias()

# Importações
import os
import re
import shutil
import requests
import zipfile
from datetime import datetime
import pandas as pd
import polars as pl
import unicodedata
import plotly.graph_objects as go


# Criação das pastas

print("Ciando pastas (se necessário)...")

diretorios = ["empresas", "estabelecimentos", "arquivos_unicos", "outputs"]
for dir_name in diretorios:
    os.makedirs(dir_name, exist_ok=True)


# Download automatico de dados

print("Download dos dados...")

base_url = "https://arquivos.receitafederal.gov.br/dados/cnpj/dados_abertos_cnpj/"
pastas = ["empresas", "estabelecimentos"]
arquivos_unicos = ["Cnaes"]
num_arq = 10
timeout_seconds = 60

# ---- data ----
hoje = datetime.today()
ano = hoje.year
mes = hoje.month
mes_atual = f"{ano}-{mes:02d}"

def mes_anterior(ano, mes):
    if mes == 1:
        return f"{ano-1}-12"
    else:
        return f"{ano}-{mes-1:02d}"

# regex para identificar arquivos já numerados: "0.EXT", "12.EXT"
numerado_re = re.compile(r"^(\d+)\.")

baixou = False
mes_para_baixar = mes_atual

for tentativa in range(2):  # tenta mês atual, depois mês anterior
    print(f"\n=> Tentativa: baixar dados de {mes_para_baixar} ...")
    sucesso_mes = True

    # Criar pastas para cada tipo neste mês (serão removidas se falhar)
    for tipo in pastas:
        pasta = f"{tipo}/{mes_para_baixar}"
        os.makedirs(pasta, exist_ok=True)
    
    # Criar pasta para arquivos únicos
    os.makedirs(f"arquivos_unicos/{mes_para_baixar}", exist_ok=True)

    try:
        # === DOWNLOAD DE ARQUIVOS ÚNICOS ===
        print(f"\n[Arquivos únicos] Baixando arquivos especiais")
        
        for arquivo_unico in arquivos_unicos:
            pasta = f"arquivos_unicos/{mes_para_baixar}"
            
            # Remove arquivos antigos desta pasta antes de baixar novamente
            if os.path.exists(pasta):
                for arquivo_antigo in os.listdir(pasta):
                    caminho_antigo = os.path.join(pasta, arquivo_antigo)
                    if os.path.isfile(caminho_antigo):
                        os.remove(caminho_antigo)
                        print(f"  -> Removido arquivo antigo: {arquivo_antigo}")
            
            url = f"{base_url}{mes_para_baixar}/{arquivo_unico}.zip"
            local_zip = os.path.join(pasta, f"{arquivo_unico}.zip")
            
            print(f"  -> Checando URL: {url}")
            try:
                r = requests.get(url, timeout=timeout_seconds)
                r.raise_for_status()
            except requests.RequestException as e:
                print(f"Falha no download: {e}")
                raise
            
            antes = set(os.listdir(pasta))
            
            with open(local_zip, "wb") as f:
                f.write(r.content)
            with zipfile.ZipFile(local_zip, "r") as zip_ref:
                zip_ref.extractall(pasta)
            os.remove(local_zip)
            
            depois = set(os.listdir(pasta))
            novos = sorted(list(depois - antes))
            
            if not novos:
                print(f"Nenhum arquivo extraído do zip {arquivo_unico} (pode ser vazio).")
            else:
                # renomeia os arquivos extraídos para o padrão desejado
                for novo in novos:
                    caminho_novo = os.path.join(pasta, novo)
                    # ignora diretórios (se houver)
                    if os.path.isdir(caminho_novo):
                        continue
                    
                    # Pega apenas a extensão após o ÚLTIMO ponto
                    # Ex: F.K03200$Z.D51011.CNAECSV -> extensão = "CNAECSV"
                    if "." in novo:
                        extensao = novo.split(".")[-1]
                    else:
                        extensao = ""
                    
                    # usa o nome base do arquivo único em minúsculo
                    destino_nome = f"{arquivo_unico.lower()}.{extensao}" if extensao else arquivo_unico.lower()
                    destino_path = os.path.join(pasta, destino_nome)
                    
                    # se destino já existir, sobrescreve
                    if os.path.exists(destino_path):
                        os.remove(destino_path)
                        print(f"Sobrescrevendo arquivo existente: {destino_nome}")
                    
                    os.rename(caminho_novo, destino_path)
                    print(f"{arquivo_unico} → renomeado {novo} → {destino_nome}")

        # === DOWNLOAD DE ARQUIVOS NUMERADOS ===
        for tipo in pastas:
            pasta = f"{tipo}/{mes_para_baixar}"

            print(f"\n[{tipo}] Baixando arquivos 0 a {num_arq-1}")

            for i in range(num_arq):
                url = f"{base_url}{mes_para_baixar}/{tipo.capitalize()}{i}.zip"
                local_zip = os.path.join(pasta, f"{i}.zip")

                print(f"  -> Checando URL: {url}")
                try:
                    r = requests.get(url, timeout=timeout_seconds)
                    r.raise_for_status()
                except requests.RequestException as e:
                    print(f"Falha no download: {e}")
                    raise  # dispara para o bloco externo tratar limpeza

                # antes da extração, lista os arquivos presentes
                antes = set(os.listdir(pasta))

                # salva e extrai
                with open(local_zip, "wb") as f:
                    f.write(r.content)
                with zipfile.ZipFile(local_zip, "r") as zip_ref:
                    zip_ref.extractall(pasta)
                os.remove(local_zip)

                # depois da extração, pega apenas os arquivos recém-criados
                depois = set(os.listdir(pasta))
                novos = sorted(list(depois - antes))

                if not novos:
                    print(f"Nenhum arquivo extraído do zip {i} (pode ser vazio).")
                else:
                    # renomeia os novos usando o índice `i` do arquivo original
                    for novo in novos:
                        caminho_novo = os.path.join(pasta, novo)
                        # ignora diretórios (se houver)
                        if os.path.isdir(caminho_novo):
                            continue
                        # se o arquivo já estiver no padrão numerado, pula
                        if numerado_re.match(novo):
                            # este caso é raro; significa que zip já trazia arquivo numerado
                            print(f"(pula) arquivo já numerado: {novo}")
                            continue
                        # pega extensão (tudo após o último ponto)
                        extensao = novo.split(".")[-1] if "." in novo else ""
                        destino_nome = f"{i}.{extensao}" if extensao else f"{i}"
                        destino_path = os.path.join(pasta, destino_nome)
                        
                        # se destino já existir, sobrescreve (comportamento esperado para re-downloads)
                        if os.path.exists(destino_path):
                            os.remove(destino_path)
                            print(f"Sobrescrevendo arquivo existente: {destino_nome}")
                        
                        os.rename(caminho_novo, destino_path)
                        print(f"{tipo} → renomeado {novo} → {destino_nome}")

        # se chegou aqui sem exceções, mês foi baixado com sucesso
        baixou = True
        break

    except Exception as e:
        # limpeza das pastas parciais do mês em tentativa
        print(f"\nErro ao baixar/extrair dados para {mes_para_baixar}: {e}")
        print("Limpando pastas parciais geradas")
        
        for tipo in pastas:
            pasta = f"{tipo}/{mes_para_baixar}"
            if os.path.exists(pasta):
                try:
                    shutil.rmtree(pasta)
                    print(f"Removida: {pasta}")
                except Exception as rm_e:
                    print(f"Falha ao remover {pasta}: {rm_e}")
        
        # Limpar pasta de arquivos únicos
        pasta_unicos = f"arquivos_unicos/{mes_para_baixar}"
        if os.path.exists(pasta_unicos):
            try:
                shutil.rmtree(pasta_unicos)
                print(f"Removida: {pasta_unicos}")
            except Exception as rm_e:
                print(f"Falha ao remover {pasta_unicos}: {rm_e}")
        
        # tenta mês anterior na próxima iteração
        mes_para_baixar = mes_anterior(ano, mes)
        continue

if not baixou:
    print("\nNenhum mês disponível (atual nem anterior).")
else:
    print(f"\nDownload e renomeação finalizados para: {mes_para_baixar}")
    for tipo in pastas:
        pasta = f"{tipo}/{mes_para_baixar}"
        print(f"  Conteúdo de {pasta}: {os.listdir(pasta)}")

    pasta_unicos = f"arquivos_unicos/{mes_para_baixar}"
    if os.path.exists(pasta_unicos):
        print(f"  Conteúdo de {pasta_unicos}: {os.listdir(pasta_unicos)}")


# Importação dos dados

# ==============================
# LEITURA E PROCESSAMENTO DE ESTABELECIMENTOS
# ==============================

#caso ja tenha os dados baixados e queira apenas executar, comente o trecho de download dos dados e defina o mes na variavel abaixo "aaaa-mm"
#mes_para_baixar="2025-11"

print("Lendo dados...")
df_estabelecimentos = pl.concat([
    pl.scan_csv(
        f"estabelecimentos/{mes_para_baixar}/{i}.ESTABELE",
        separator=";",
        encoding="utf8-lossy",
        has_header=False,
        schema_overrides={f"column_{j}": pl.String for j in range(31)}
    )
    .filter(
        pl.col("column_21").is_not_null() &
        (pl.col("column_21") == "6361") &
        (pl.col("column_6") == "02")
    )
    .select([
        "column_1", "column_2", "column_3", "column_4", "column_5", "column_6", "column_7",
        "column_11", "column_12", "column_13", "column_14", "column_15", "column_16",
        "column_17", "column_18", "column_19", "column_21", "column_22", "column_23", "column_28"
    ])
    .rename({
        "column_1": "cnpj_basico", "column_2": "cnpj_ordem", "column_3": "cnpj_dv",
        "column_4": "id_matriz_filial", "column_5": "nome_fantasia", "column_6": "sit_cadastral",
        "column_7": "dt_sit_cadastral", "column_11": "dt_inicio_atv", "column_12": "cnae_principal",
        "column_13": "cnae_secundario", "column_14": "tipo_logradouro", "column_15": "logradouro",
        "column_16": "numero", "column_17": "complemento", "column_18": "bairro", "column_19": "cep",
        "column_21": "municipio", "column_22": "ddd", "column_23": "telefone", "column_28": "email"
    })
    .with_columns([
        pl.col("email").str.to_lowercase(),
        pl.col("dt_sit_cadastral").str.to_date(format="%Y%m%d", strict=False).alias("dt_sit_cadastral"),
        pl.col("dt_inicio_atv").str.to_date(format="%Y%m%d", strict=False).alias("dt_inicio_atv")
    ])
    for i in range(10)
]).collect()

# ==============================
# LEITURA E PROCESSAMENTO DE EMPRESAS
# ==============================
df_empresas = pl.concat([
    pl.scan_csv(
        f"empresas/{mes_para_baixar}/{i}.EMPRECSV",
        separator=";",
        encoding="utf8-lossy",
        has_header=False,
        schema_overrides={f"column_{j}": pl.String for j in range(9)}
    )
    .select(["column_1", "column_2", "column_4", "column_5", "column_6"])
    .rename({
        "column_1": "cnpj_basico", "column_2": "razao_social",
        "column_4": "qualif_responsavel", "column_5": "capital_social",
        "column_6": "porte"
    })
    for i in range(10)
]).collect()

# ==============================
# JUNÇÃO E TRATAMENTO FINAL
# ==============================
ordem_colunas = [
    "cnpj_basico", "cnpj_ordem", "cnpj_dv", "id_matriz_filial", "razao_social",
    "nome_fantasia", "qualif_responsavel", "capital_social", "porte",
    "sit_cadastral", "dt_sit_cadastral", "dt_inicio_atv", "cnae_principal",
    "cnae_secundario", "tipo_logradouro", "logradouro", "numero",
    "complemento", "bairro", "cep", "municipio", "telefone_completo", "email"
]

df_final = (
    df_estabelecimentos
    .join(df_empresas, on="cnpj_basico", how="left")
    .with_columns([
        pl.col("sit_cadastral").replace(old=["02"], new=["ATIVA"]),
        pl.col("municipio").replace(old=["6361"], new=["COTIA"]),
        pl.col("porte").replace(
            old=["00", "01", "03", "05"],
            new=["NÃO INFORMADO", "MICRO EMPRESA", "EMPRESA DE PEQUENO PORTE", "DEMAIS"]
        ),
        pl.concat_str([pl.col("ddd"), pl.col("telefone")], separator="").alias("telefone_completo")
    ])
    .select(ordem_colunas)
)

# ==============================
# LEITURA DE CNAE E CLASSIFICAÇÃO
# ==============================
df_cnae = (
    pl.read_csv(f"arquivos_unicos/{mes_para_baixar}/cnaes.CNAECSV", separator=";", encoding="latin1", has_header=False)
    .rename({"column_1": "cod_cnae", "column_2": "descricao"})
    .with_columns([
        pl.col("descricao").cast(pl.Utf8),
        pl.col("cod_cnae").cast(pl.Utf8).str.zfill(7).alias("cod_cnae")
    ])
)

#Importação do arquivo excel
url_excel_filtro = "selecao_atv_nomes.xlsx"
cols_filtro_cnae = ["Classe", "Denominação", "Seleção"]
df_filtro_cnae = pd.read_excel(url_excel_filtro, sheet_name= "Filtro por cnae" ,usecols=cols_filtro_cnae)
df_filtro_cnae['cnae_padronizado'] = (
    df_filtro_cnae['Classe']
    .astype(str)
    .str.replace('.', '', regex=False)
    .str.replace('-', '', regex=False)
)

df_filtro_nome = pd.read_excel(url_excel_filtro, sheet_name= "Filtro por nome")


# joins

# ==============================
# CNAE PRINCIPAL
# ==============================
cnae_principal = (
    df_final
    .filter(pl.col("cnae_principal").is_not_null() & (pl.col("cnae_principal") != ""))
    .group_by("cnae_principal")
    .agg(pl.len().alias("frequencia"))
    .join(
        df_cnae.select(["cod_cnae", "descricao"]),
        left_on="cnae_principal",
        right_on="cod_cnae",
        how="left"
    )
    .with_columns(
        pl.when(pl.col("descricao").is_null())
        .then(pl.col("cnae_principal"))
        .otherwise(pl.col("descricao"))
        .alias("descricao_final")
    )
    .sort("frequencia", descending=True)
)

df_principal = pd.DataFrame({
    "codigo": cnae_principal["cnae_principal"].to_list(),
    "frequencia": cnae_principal["frequencia"].to_list(),
    "descricao": cnae_principal["descricao_final"].to_list()
})
df_principal["tipo"] = "principal"


# ==============================
# CNAE SECUNDÁRIO
# ==============================
cnae_secundarios = (
    df_final
    .select("cnae_secundario")
    .with_columns(pl.col("cnae_secundario").str.split(","))  # separa os códigos por vírgula
    .explode("cnae_secundario")  # transforma cada código em uma linha
    .with_columns(pl.col("cnae_secundario").str.strip_chars())  # remove espaços
    .filter(pl.col("cnae_secundario").is_not_null() & (pl.col("cnae_secundario") != ""))
    .group_by("cnae_secundario")
    .agg(pl.len().alias("frequencia"))
    .join(
        df_cnae.select(["cod_cnae", "descricao"]),
        left_on="cnae_secundario",
        right_on="cod_cnae",
        how="left"
    )
    .with_columns(
        pl.when(pl.col("descricao").is_null())
        .then(pl.col("cnae_secundario"))
        .otherwise(pl.col("descricao"))
        .alias("descricao_final")
    )
    .sort("frequencia", descending=True)
)

df_secundarios = pd.DataFrame({
    "codigo": cnae_secundarios["cnae_secundario"].to_list(),
    "frequencia": cnae_secundarios["frequencia"].to_list(),
    "descricao": cnae_secundarios["descricao_final"].to_list()
})
df_secundarios["tipo"] = "secundario"


# ==============================
# UNIÃO E AGREGAÇÃO
# ==============================
df_unico = pd.concat([df_principal, df_secundarios], ignore_index=True)

# Criar coluna de classe CNAE (5 primeiros dígitos)
df_unico["classe_cnae"] = df_unico["codigo"].astype(str).str[:5]

# Merge com a classificação detalhada
df_unico = df_unico.merge(
    df_filtro_cnae[["cnae_padronizado", "Denominação"]],
    left_on="classe_cnae",
    right_on="cnae_padronizado",
    how="left"
)

# Renomear e limpar colunas auxiliares
df_unico = (
    df_unico
    .rename(columns={"Denominação": "classe"})
    .drop(columns=["classe_cnae", "cnae_padronizado"])
)


# Filtra as empresas de comex

######################## VERSAO QUE FUNCIONOU #######################

print("Filtrando Comex...")

def normalizar_texto(texto):
    """
    Remove acentos, converte para minúsculas e remove espaços extras
    Exemplo: 'Importação' vai para 'importacao'
    """
    if texto is None or pd.isna(texto):
        return ""
    
    # Converter para string
    texto = str(texto)
    
    # Remover acentos (NFD = decompõe caracteres acentuados)
    texto_nfd = unicodedata.normalize('NFD', texto)
    texto_sem_acento = ''.join(char for char in texto_nfd if unicodedata.category(char) != 'Mn')
    
    # Converter para minúsculas e remover espaços extras
    texto_normalizado = texto_sem_acento.lower().strip()
    
    return texto_normalizado

# PROCESSAR LISTAS DE CNAEs

# CNAEs que devem ser ACEITOS (marcados como "Aceita")
cnaes_aceitos = df_filtro_cnae[
    df_filtro_cnae['Seleção'].str.upper() == 'ACEITA']['cnae_padronizado'].tolist()

# CNAEs que devem ser REJEITADOS (marcados como "Rejeita")
cnaes_rejeitados = df_filtro_cnae[
    df_filtro_cnae['Seleção'].str.upper() == 'REJEITA']['cnae_padronizado'].tolist()

# PROCESSAR PALAVRAS-CHAVE COM NORMALIZAÇÃO

# Extrair palavras-chave da planilha
palavras_chave_originais = df_filtro_nome['Nome Considerado em: Razão social ou nome fantasia'].dropna().tolist()

# Normalizar palavras-chave para busca
palavras_chave_normalizadas = []

for palavra in palavras_chave_originais:
    # Verificar se é um regex (contém caracteres especiais de regex)
    if any(char in str(palavra) for char in ['[', ']', '(', ')', '|', '\\', '^', '$', '*', '+', '?', '.']):
        # Se for regex, mantém como está (assume que usuário sabe o que está fazendo)
        palavras_chave_normalizadas.append(str(palavra))
    else:
        # Se não for regex, normaliza (remove acentos)
        palavra_normalizada = normalizar_texto(palavra)
        # Escapa caracteres especiais para usar como texto literal no regex
        palavra_escapada = re.escape(palavra_normalizada)
        palavras_chave_normalizadas.append(palavra_escapada)

# ==============================
# CRIAR COLUNAS NORMALIZADAS NO df_final
# ==============================

print("\nCriando colunas normalizadas para busca...")

df_final_normalizado = df_final.with_columns([
    pl.col("razao_social").map_elements(
        lambda x: normalizar_texto(x) if x is not None else "",
        return_dtype=pl.Utf8
    ).alias("razao_social_norm"),
    
    pl.col("nome_fantasia").map_elements(
        lambda x: normalizar_texto(x) if x is not None else "",
        return_dtype=pl.Utf8
    ).alias("nome_fantasia_norm")
])

# ==============================
# CRIAR REGEX PARA PALAVRAS-CHAVE
# ==============================

if palavras_chave_normalizadas:
    # Case insensitive não é necessário pois já normalizamos tudo para minúsculas
    regex_pattern = "|".join(palavras_chave_normalizadas)
else:
    regex_pattern = None

# ==============================
# CRIAR REGEX PARA CNAEs
# ==============================

# Regex para CNAEs ACEITOS nos secundários (busca apenas os 5 primeiros dígitos)
if cnaes_aceitos:
    regex_cnaes_aceitos = r"(?:^|,\s*)(" + "|".join(cnaes_aceitos) + r")(?:\d{2})?(?:,|$)"
else:
    regex_cnaes_aceitos = None

# Regex para CNAEs REJEITADOS nos secundários
if cnaes_rejeitados:
    regex_cnaes_rejeitados = r"(?:^|,\s*)(" + "|".join(cnaes_rejeitados) + r")(?:\d{2})?(?:,|$)"
else:
    regex_cnaes_rejeitados = None

# ==============================
# APLICAR FILTRO NO df_final
# ==============================

# Construir filtro de ACEITAÇÃO (condições OR)
filtro_aceita = pl.lit(False)  # Começa com False

# Adicionar filtro por palavras-chave (se houver) - AGORA USA COLUNAS NORMALIZADAS
if regex_pattern:
    filtro_aceita = (
        filtro_aceita | 
        pl.col("razao_social_norm").str.contains(regex_pattern) |
        pl.col("nome_fantasia_norm").str.contains(regex_pattern)
    )

# Adicionar filtro por CNAE principal (se houver CNAEs aceitos)
if cnaes_aceitos:
    filtro_aceita = (
        filtro_aceita |
        pl.col("cnae_principal").str.slice(0, 5).is_in(cnaes_aceitos)
    )

# Adicionar filtro por CNAEs secundários (se houver CNAEs aceitos)
if regex_cnaes_aceitos:
    filtro_aceita = (
        filtro_aceita |
        pl.col("cnae_secundario").str.contains(regex_cnaes_aceitos)
    )

# Construir filtro de REJEIÇÃO (condições OR)
filtro_rejeita = pl.lit(False)  # Começa com False

# Adicionar filtro de rejeição por CNAE principal (se houver CNAEs rejeitados)
if cnaes_rejeitados:
    filtro_rejeita = (
        filtro_rejeita |
        pl.col("cnae_principal").str.slice(0, 5).is_in(cnaes_rejeitados)
    )

# Adicionar filtro de rejeição por CNAEs secundários (se houver CNAEs rejeitados)
if regex_cnaes_rejeitados:
    filtro_rejeita = (
        filtro_rejeita |
        pl.col("cnae_secundario").str.contains(regex_cnaes_rejeitados)
    )

# Aplicar filtro: ACEITA (algum critério positivo) E NÃO REJEITA (nenhum critério negativo)
df_comex = df_final_normalizado.filter(filtro_aceita & ~filtro_rejeita)

# Remover colunas normalizadas auxiliares (se não precisar mais delas)
df_comex = df_comex.drop(["razao_social_norm", "nome_fantasia_norm"])



# Obter lista de CNPJs das empresas de comércio exterior
cnpj_comex = df_comex.select("cnpj_basico").unique()

# Marcar no df_final quais empresas são de Comex
df_final = df_final.with_columns(
    pl.col("cnpj_basico").is_in(cnpj_comex["cnpj_basico"].implode()).alias("is_comex")
)

# Atualiza df_comex para referência futura (somente empresas marcadas)
df_comex = df_final.filter(pl.col("is_comex"))



# ==============================
# PREPARAR CONTEÚDO PARA A SEÇÃO DE METODOLOGIA
# ==============================

# 1. CNAEs ACEITOS com descrição
df_aceitos = df_filtro_cnae[df_filtro_cnae['Seleção'].str.upper() == 'ACEITA'][['Classe', 'Denominação']]
lista_cnaes_aceitos = [
    f"{row['Classe']} – {row['Denominação']}"
    for _, row in df_aceitos.iterrows()
]

# 2. CNAEs REJEITADOS com descrição
df_rejeitados = df_filtro_cnae[df_filtro_cnae['Seleção'].str.upper() == 'REJEITA'][['Classe', 'Denominação']]
lista_cnaes_rejeitados = [
    f"{row['Classe']} – {row['Denominação']}"
    for _, row in df_rejeitados.iterrows()
]

# 3. Palavras-chave (já temos a lista original)
lista_palavras = palavras_chave_originais if palavras_chave_originais else ["(nenhuma palavra definida)"]

# Função para formatar listas longas em blocos de texto com quebras de linha
def formatar_lista_para_html(lista):
    if not lista:
        return "<em>Nenhum item definido</em>"
    return "<br>".join(f"• {item}" for item in lista)


# Análises estatísticas

# ==============================
# ESTATÍSTICAS GERAIS
# ==============================

print("Analisando dados...")

print("- " * 40)
print("ANÁLISE ESTATÍSTICA - EMPRESAS DE COMÉRCIO EXTERIOR EM COTIA")
print("- " * 40)

# Total de empresas em Cotia
total_empresas = df_final.height
print(f"\nTotal de empresas ativas em Cotia: {total_empresas:,}".replace(",", "."))

# Total de empresas de Comex
total_comex = df_comex.height
print(f"Total de empresas com atividades comex: {total_comex:,}".replace(",", "."))

# Percentual de empresas Comex
percentual_comex = (total_comex / total_empresas) * 100
print(f"Percentual de empresas com atividade Comex: {percentual_comex:.2f}%")

# Empresas NÃO-Comex
total_nao_comex = total_empresas - total_comex
print(f"Total de empresas sem atividade Comex: {total_nao_comex:,}".replace(",", "."))

print("\n" + "- " * 40)

# Comex como atividade PRINCIPAL vs SECUNDÁRIA
# ATUALIZADO: Agora usa os CNAEs aceitos da nova filtragem
if cnaes_aceitos:  # Verifica se há CNAEs aceitos configurados
    comex_principal = df_comex.filter(
        pl.col("cnae_principal").str.slice(0, 5).is_in(cnaes_aceitos)
    ).height
    
    comex_secundaria = total_comex - comex_principal
    
    print(f"\nEmpresas com Comex como atividade PRINCIPAL: {comex_principal:,} ({(comex_principal/total_comex)*100:.1f}%)".replace(",", "."))
    print(f"Empresas com Comex como atividade SECUNDÁRIA: {comex_secundaria:,} ({(comex_secundaria/total_comex)*100:.1f}%)".replace(",", "."))
else:
    print(f"\n  Análise de CNAE principal vs secundário não disponível (sem CNAEs configurados)")

# ==============================
# 2. ANÁLISE TEMPORAL
# ==============================
print("\n" + "- " * 40)
print("ANÁLISE TEMPORAL")
print("- " * 40 + "\n")

# Função auxiliar para formatar idade decimal em anos e meses
def formatar_idade(anos_decimal):
    if anos_decimal is None or pd.isna(anos_decimal):
        return "N/A"
    
    anos = int(anos_decimal)
    meses = int((anos_decimal - anos) * 12)
    
    if anos == 0:
        return f"{meses} {'mês' if meses == 1 else 'meses'}"
    elif meses == 0:
        return f"{anos} {'ano' if anos == 1 else 'anos'}"
    else:
        return f"{anos} {'ano' if anos == 1 else 'anos'} e {meses} {'mês' if meses == 1 else 'meses'}"

# 2.1 - Tempo médio de atividade (geral)
print("TEMPO MÉDIO DE ATIVIDADE:\n")

# Data atual
data_atual = datetime.today()

# Calcular idade em anos para todas as empresas com data válida
# ATUALIZADO: Usa df_final que já tem a coluna is_comex
df_com_idade = df_final.with_columns([
    (
        (pl.lit(data_atual.year) - pl.col("dt_inicio_atv").dt.year()) +
        ((pl.lit(data_atual.month) - pl.col("dt_inicio_atv").dt.month()) / 12) +
        ((pl.lit(data_atual.day) - pl.col("dt_inicio_atv").dt.day()) / 365)
    ).alias("idade_anos")
]).filter(
    pl.col("dt_inicio_atv").is_not_null() & (pl.col("idade_anos") >= 0)
)

# Estatísticas gerais
empresas_com_data = df_com_idade.height
empresas_sem_data = df_final.height - empresas_com_data

idade_media_geral = df_com_idade.select(pl.col("idade_anos").mean()).item()
idade_mediana_geral = df_com_idade.select(pl.col("idade_anos").median()).item()
idade_min = df_com_idade.select(pl.col("idade_anos").min()).item()
idade_max = df_com_idade.select(pl.col("idade_anos").max()).item()

print(f"Total de empresas com data de início: {empresas_com_data:,}".replace(",", "."))
print(f"Total de empresas SEM data de início: {empresas_sem_data:,}".replace(",", "."))
print(f"\nIdade média (GERAL): {formatar_idade(idade_media_geral)}")
print(f"Idade mediana (GERAL): {formatar_idade(idade_mediana_geral)}")
print(f"Empresa mais antiga: {formatar_idade(idade_max)}")
print(f"Empresa mais nova: {formatar_idade(idade_min)}")

print("\n" + "- " * 40)

# 2.2 - Comparação de idade média: Comex vs Não-Comex
# ATUALIZADO: Usa a coluna is_comex diretamente
print("\nCOMPARAÇÃO DE IDADE MÉDIA: COMEX vs NÃO-COMEX\n")

comparacao_idade = (
    df_com_idade
    .group_by("is_comex")
    .agg([
        pl.len().alias("quantidade"),
        pl.col("idade_anos").mean().alias("idade_media"),
        pl.col("idade_anos").median().alias("idade_mediana"),
        pl.col("idade_anos").std().alias("desvio_padrao")
    ])
    .sort("is_comex")
)

for row in comparacao_idade.iter_rows(named=True):
    tipo = "COMEX" if row["is_comex"] else "NÃO-COMEX"
    print(f"{tipo}:")
    print(f"Quantidade: {row['quantidade']:,}".replace(",", "."))
    print(f"Idade média: {formatar_idade(row['idade_media'])}")
    print(f"Idade mediana: {formatar_idade(row['idade_mediana'])}")
    print(f"Desvio padrão: {formatar_idade(row['desvio_padrao'])}")
    print()

# Calcular diferença percentual
idade_comex = comparacao_idade.filter(pl.col("is_comex")).select("idade_media").item()
idade_nao_comex = comparacao_idade.filter(~pl.col("is_comex")).select("idade_media").item()

# Calcular mín/máx por grupo usando df_com_idade
idade_extremos = (
    df_com_idade
    .group_by("is_comex")
    .agg([
        pl.col("idade_anos").min().alias("idade_min"),
        pl.col("idade_anos").max().alias("idade_max")
    ])
)

# Extrair valores para COMEX (is_comex = True)
comex_stats = comparacao_idade.filter(pl.col("is_comex")).to_dicts()[0]
comex_extremos = idade_extremos.filter(pl.col("is_comex")).to_dicts()[0]

idade_comex_media = comex_stats["idade_media"]
idade_comex_mediana = comex_stats["idade_mediana"]
idade_comex_min = comex_extremos["idade_min"]
idade_comex_max = comex_extremos["idade_max"]

diferenca_percentual = ((idade_comex - idade_nao_comex) / idade_nao_comex) * 100

if diferenca_percentual > 0:
    print(f"Empresas de Comex são em média {abs(diferenca_percentual):.1f}% MAIS ANTIGAS que empresas não-Comex")
else:
    print(f"Empresas de Comex são em média {abs(diferenca_percentual):.1f}% MAIS NOVAS que empresas não-Comex")

print("\n" + "- " * 40)

# 2.3 - Distribuição por faixas etárias
print("\nDISTRIBUIÇÃO POR FAIXAS ETÁRIAS:\n")

df_faixas = df_com_idade.with_columns(
    pl.when(pl.col("idade_anos") < 5).then(pl.lit("0-5 anos"))
    .when(pl.col("idade_anos") < 10).then(pl.lit("5-10 anos"))
    .when(pl.col("idade_anos") < 15).then(pl.lit("10-15 anos"))
    .when(pl.col("idade_anos") < 20).then(pl.lit("15-20 anos"))
    .when(pl.col("idade_anos") < 30).then(pl.lit("20-30 anos"))
    .otherwise(pl.lit("30+ anos"))
    .alias("faixa_etaria")
)

distribuicao_faixas = (
    df_faixas
    .group_by(["faixa_etaria", "is_comex"])
    .agg(pl.len().alias("quantidade"))
    .sort(["is_comex", "faixa_etaria"])
)

# Ordenação customizada para as faixas
ordem_faixas = ["0-5 anos", "5-10 anos", "10-15 anos", "15-20 anos", "20-30 anos", "30+ anos"]

print("GERAL:")
total_geral_com_data = df_faixas.filter(~pl.col("is_comex")).height
for faixa in ordem_faixas:
    qtd_row = distribuicao_faixas.filter(
        (~pl.col("is_comex")) & (pl.col("faixa_etaria") == faixa)
    )
    if qtd_row.height > 0:
        qtd = qtd_row.select("quantidade").item()
        pct = (qtd / total_geral_com_data) * 100
        print(f"  {faixa}: {qtd:,} ({pct:.1f}%)".replace(",", "."))

print("\nCOMEX:")
total_comex_com_data = df_faixas.filter(pl.col("is_comex")).height
for faixa in ordem_faixas:
    qtd_row = distribuicao_faixas.filter(
        (pl.col("is_comex")) & (pl.col("faixa_etaria") == faixa)
    )
    if qtd_row.height > 0:
        qtd = qtd_row.select("quantidade").item()
        pct = (qtd / total_comex_com_data) * 100
        print(f"  {faixa}: {qtd:,} ({pct:.1f}%)".replace(",", "."))

print("\n" + "- " * 40)

# 2.4 - Evolução mensal do número de empresas ativas
print("\nEVOLUÇÃO MENSAL DE ABERTURA DE EMPRESAS:\n")

# Criar coluna ano_mes
df_evolucao = df_com_idade.with_columns([
    pl.col("dt_inicio_atv").dt.strftime("%Y-%m").alias("ano_mes")
])

# Agrupar por ano_mes e is_comex
evolucao_mensal = (
    df_evolucao
    .group_by(["ano_mes", "is_comex"])
    .agg(pl.len().alias("quantidade"))
    .sort("ano_mes")
)

# Estatísticas dos últimos 12 meses
data_limite = datetime(data_atual.year - 1, data_atual.month, 1)
ano_mes_limite = data_limite.strftime("%Y-%m")

ultimos_12_meses = evolucao_mensal.filter(pl.col("ano_mes") >= ano_mes_limite)

print("ÚLTIMOS 12 MESES:")
print(f"{'Mês':<12} {'Total':<10} {'Comex':<10} {'Não-Comex':<10}")
print("-" * 42)

# Agrupar dados por mês
meses_unicos = sorted(ultimos_12_meses.select("ano_mes").unique().to_series().to_list())

for mes in meses_unicos:
    dados_mes = ultimos_12_meses.filter(pl.col("ano_mes") == mes)
    
    total_mes = dados_mes.select("quantidade").sum().item()
    
    comex_mes = dados_mes.filter(pl.col("is_comex")).select("quantidade")
    qtd_comex = comex_mes.sum().item() if comex_mes.height > 0 else 0
    
    nao_comex_mes = dados_mes.filter(~pl.col("is_comex")).select("quantidade")
    qtd_nao_comex = nao_comex_mes.sum().item() if nao_comex_mes.height > 0 else 0
    
    print(f"{mes:<12} {total_mes:<10} {qtd_comex:<10} {qtd_nao_comex:<10}")

print("\n" + "- " * 40)

# 2.5 - Taxa de crescimento anual
print("\nCRESCIMENTO ANUAL DE EMPRESAS:\n")

df_anual = df_com_idade.with_columns([
    pl.col("dt_inicio_atv").dt.year().alias("ano")
]).filter(
    pl.col("ano") >= (data_atual.year - 10)  # Últimos 10 anos
)

crescimento_anual = (
    df_anual
    .group_by(["ano", "is_comex"])
    .agg(pl.len().alias("quantidade"))
    .sort("ano")
)

print("ÚLTIMOS 10 ANOS:")
print(f"{'Ano':<8} {'Total':<10} {'Comex':<10} {'Não-Comex':<10}")
print("-" * 38)

anos_unicos = sorted(crescimento_anual.select("ano").unique().to_series().to_list())

for ano in anos_unicos:
    dados_ano = crescimento_anual.filter(pl.col("ano") == ano)
    
    total_ano = dados_ano.select("quantidade").sum().item()
    
    comex_ano = dados_ano.filter(pl.col("is_comex")).select("quantidade")
    qtd_comex = comex_ano.sum().item() if comex_ano.height > 0 else 0
    
    nao_comex_ano = dados_ano.filter(~pl.col("is_comex")).select("quantidade")
    qtd_nao_comex = nao_comex_ano.sum().item() if nao_comex_ano.height > 0 else 0
    
    print(f"{ano:<8} {total_ano:<10} {qtd_comex:<10} {qtd_nao_comex:<10}")

print("\n" + "- " * 40)

# 2.6 - Estatísticas de crescimento acumulado
print("\nESTATÍSTICAS DE CRESCIMENTO ACUMULADO (ÚLTIMOS 12 MESES):\n")

# Preparar dados para cálculo de crescimento acumulado
meses_nomes = ['Jan', 'Fev', 'Mar', 'Abr', 'Mai', 'Jun', 
               'Jul', 'Ago', 'Set', 'Out', 'Nov', 'Dez']

dados_crescimento_acumulado = []

for mes in meses_unicos:
    # Converter string ano-mes para data (último dia do mês)
    ano, mes_num = mes.split('-')
    ano = int(ano)
    mes_num = int(mes_num)
    
    # Último dia do mês
    if mes_num == 12:
        data_fim = datetime(ano + 1, 1, 1)
    else:
        data_fim = datetime(ano, mes_num + 1, 1)
    
    # Contar empresas que iniciaram atividade até esta data
    empresas_ate_mes = df_com_idade.filter(
        pl.col("dt_inicio_atv") < pl.lit(data_fim)
    )
    
    total_ate_mes = empresas_ate_mes.height
    
    comex_ate_mes = empresas_ate_mes.filter(pl.col("is_comex")).height
    nao_comex_ate_mes = empresas_ate_mes.filter(~pl.col("is_comex")).height
    
    # Formatar mês para exibição
    mes_formatado = f"{meses_nomes[mes_num-1]}/{str(ano)[2:]}"
    
    dados_crescimento_acumulado.append({
        'ano_mes': mes,
        'mes_formatado': mes_formatado,
        'total': total_ate_mes,
        'comex': comex_ate_mes,
        'nao_comex': nao_comex_ate_mes
    })

df_crescimento_acumulado = pd.DataFrame(dados_crescimento_acumulado)

# Calcular crescimento percentual mês a mês
df_crescimento_acumulado['crescimento_total_pct'] = df_crescimento_acumulado['total'].pct_change() * 100
df_crescimento_acumulado['crescimento_comex_pct'] = df_crescimento_acumulado['comex'].pct_change() * 100
df_crescimento_acumulado['crescimento_nao_comex_pct'] = df_crescimento_acumulado['nao_comex'].pct_change() * 100

# Estatísticas do crescimento
crescimento_total_periodo = (
    (df_crescimento_acumulado['total'].iloc[-1] - df_crescimento_acumulado['total'].iloc[0]) /
    df_crescimento_acumulado['total'].iloc[0] * 100
)

crescimento_comex_periodo = (
    (df_crescimento_acumulado['comex'].iloc[-1] - df_crescimento_acumulado['comex'].iloc[0]) /
    df_crescimento_acumulado['comex'].iloc[0] * 100
)

crescimento_nao_comex_periodo = (
    (df_crescimento_acumulado['nao_comex'].iloc[-1] - df_crescimento_acumulado['nao_comex'].iloc[0]) /
    df_crescimento_acumulado['nao_comex'].iloc[0] * 100
)

print(f"Crescimento TOTAL no período: {crescimento_total_periodo:+.2f}%")
print(f"De {df_crescimento_acumulado['total'].iloc[0]:,} para {df_crescimento_acumulado['total'].iloc[-1]:,} empresas".replace(",", "."))
print(f"Diferença: {(df_crescimento_acumulado['total'].iloc[-1] - df_crescimento_acumulado['total'].iloc[0]):+,} empresas".replace(",", "."))

print(f"\nCrescimento COMEX no período: {crescimento_comex_periodo:+.2f}%")
print(f"De {df_crescimento_acumulado['comex'].iloc[0]:,} para {df_crescimento_acumulado['comex'].iloc[-1]:,} empresas".replace(",", "."))
print(f"Diferença: {(df_crescimento_acumulado['comex'].iloc[-1] - df_crescimento_acumulado['comex'].iloc[0]):+,} empresas".replace(",", "."))

print(f"\nCrescimento NÃO-COMEX no período: {crescimento_nao_comex_periodo:+.2f}%")
print(f"De {df_crescimento_acumulado['nao_comex'].iloc[0]:,} para {df_crescimento_acumulado['nao_comex'].iloc[-1]:,} empresas".replace(",", "."))
print(f"Diferença: {(df_crescimento_acumulado['nao_comex'].iloc[-1] - df_crescimento_acumulado['nao_comex'].iloc[0]):+,} empresas".replace(",", "."))

# Crescimento médio mensal
crescimento_medio_mensal = df_crescimento_acumulado['crescimento_total_pct'].mean()
print(f"\nCrescimento médio mensal (Total): {crescimento_medio_mensal:.2f}%")


# Analise Geográfica

# ==============================
# ANÁLISE GEOGRÁFICA
# ==============================

print("- " * 40)
print("ANÁLISE GEOGRÁFICA")
print("- " * 40 + "\n")

# 4.1 - Limpeza e padronização dos bairros
print("DISTRIBUIÇÃO POR BAIRROS:\n")

# Padronizar nomes de bairros (uppercase, remover espaços extras)
df_geo = df_final.with_columns([
    pl.col("bairro").str.to_uppercase().str.strip_chars().alias("bairro_padronizado")
])

# Contar empresas por bairro (geral)
distribuicao_bairros_geral = (
    df_geo
    .filter(pl.col("bairro_padronizado").is_not_null() & (pl.col("bairro_padronizado") != ""))
    .group_by("bairro_padronizado")
    .agg(pl.len().alias("total_empresas"))
    .sort("total_empresas", descending=True)
)

# Contar empresas Comex por bairro
distribuicao_bairros_comex = (
    df_geo
    .filter(
        pl.col("is_comex") & 
        pl.col("bairro_padronizado").is_not_null() & 
        (pl.col("bairro_padronizado") != "")
    )
    .group_by("bairro_padronizado")
    .agg(pl.len().alias("total_comex"))
    .sort("total_comex", descending=True)
)

# Join para ter total e comex no mesmo dataframe
distribuicao_completa = (
    distribuicao_bairros_geral
    .join(distribuicao_bairros_comex, on="bairro_padronizado", how="left")
    .with_columns([
        pl.col("total_comex").fill_null(0),
        ((pl.col("total_comex").fill_null(0) / pl.col("total_empresas")) * 100).alias("percentual_comex")
    ])
    .sort("total_empresas", descending=True)
)

# Estatísticas gerais
total_empresas_com_bairro = distribuicao_bairros_geral.select("total_empresas").sum().item()
total_empresas_sem_bairro = df_final.height - total_empresas_com_bairro
total_bairros_unicos = distribuicao_completa.height

print(f"Total de empresas COM bairro cadastrado: {total_empresas_com_bairro:,}".replace(",", "."))
print(f"Total de empresas SEM bairro cadastrado: {total_empresas_sem_bairro:,}".replace(",", "."))
print(f"Total de bairros únicos identificados: {total_bairros_unicos:,}".replace(",", "."))

print("\n" + "- " * 40)

# 4.2 - Top 20 bairros com mais empresas (geral)
print("\nTOP 20 BAIRROS COM MAIS EMPRESAS (GERAL):\n")

print(f"{'Posição':<10} {'Bairro':<35} {'Total':<10} {'Comex':<10} {'% Comex':<10}")
print("- " * 40)

top_20_geral = distribuicao_completa.head(20)

for idx, row in enumerate(top_20_geral.iter_rows(named=True), 1):
    bairro = row["bairro_padronizado"][:33]  # Limitar tamanho
    total = row["total_empresas"]
    comex = row["total_comex"]
    pct = row["percentual_comex"]
    
    print(f"{idx:<10} {bairro:<35} {total:<10} {comex:<10} {pct:<10.1f}%")

print("\n" + "- " * 40)

# 4.3 - Top 20 bairros com mais empresas COMEX
print("\nTOP 20 BAIRROS COM MAIS EMPRESAS DE COMEX:\n")

distribuicao_comex_ordenada = distribuicao_completa.sort("total_comex", descending=True)

print(f"{'Posição':<10} {'Bairro':<35} {'Comex':<10} {'Total':<10} {'% Comex':<10}")
print("- " * 40)

top_20_comex = distribuicao_comex_ordenada.head(20)

for idx, row in enumerate(top_20_comex.iter_rows(named=True), 1):
    bairro = row["bairro_padronizado"][:33]
    comex = row["total_comex"]
    total = row["total_empresas"]
    pct = row["percentual_comex"]
    
    print(f"{idx:<10} {bairro:<35} {comex:<10} {total:<10} {pct:<10.1f}%")

print("\n" + "- " * 40)

# 4.4 - Bairros com maior concentração percentual de Comex
print("\nTOP 15 BAIRROS COM MAIOR CONCENTRAÇÃO DE COMEX (mínimo 10 empresas):\n")

bairros_concentracao = (
    distribuicao_completa
    .filter(pl.col("total_empresas") >= 10)  # Filtrar apenas bairros com pelo menos 10 empresas
    .sort("percentual_comex", descending=True)
    .head(15)
)

print(f"{'Posição':<10} {'Bairro':<35} {'% Comex':<12} {'Comex':<10} {'Total':<10}")
print("- " * 40)

for idx, row in enumerate(bairros_concentracao.iter_rows(named=True), 1):
    bairro = row["bairro_padronizado"][:33]
    pct = row["percentual_comex"]
    comex = row["total_comex"]
    total = row["total_empresas"]
    
    print(f"{idx:<10} {bairro:<35} {pct:<12.1f}% {comex:<10} {total:<10}")

print("\n" + "- " * 40)


# Graficos plotly html

# ==============================
# GRÁFICOS INTERATIVOS - ANÁLISE TEMPORAL
# ==============================

print("Ciando Gráficos...")

import plotly.graph_objects as go
from plotly.subplots import make_subplots

# Criar pasta para salvar os gráficos
os.makedirs("graficos", exist_ok=True)

# ==============================
# GRÁFICO 1: CRESCIMENTO ANUAL (ÚLTIMOS 10 ANOS)
# ==============================

print("Gerando gráfico 1/3: Crescimento Anual...")

# Preparar dados para o gráfico anual
dados_grafico_anual = []
for ano in anos_unicos:
    dados_ano = crescimento_anual.filter(pl.col("ano") == ano)
    
    total_ano = dados_ano.select("quantidade").sum().item()
    
    comex_ano = dados_ano.filter(pl.col("is_comex")).select("quantidade")
    qtd_comex = comex_ano.sum().item() if comex_ano.height > 0 else 0
    
    nao_comex_ano = dados_ano.filter(~pl.col("is_comex")).select("quantidade")
    qtd_nao_comex = nao_comex_ano.sum().item() if nao_comex_ano.height > 0 else 0
    
    dados_grafico_anual.append({
        'ano': ano,
        'total': total_ano,
        'comex': qtd_comex,
        'nao_comex': qtd_nao_comex
    })

df_grafico_anual = pd.DataFrame(dados_grafico_anual)

# Criar gráfico de linhas para crescimento anual
fig_anual = go.Figure()

# Linha Total
fig_anual.add_trace(go.Scatter(
    x=df_grafico_anual['ano'],
    y=df_grafico_anual['total'],
    mode='lines+markers',
    name='Total',
    line=dict(color='#1f77b4', width=3),
    marker=dict(size=8),
    hovertemplate='<b>Ano:</b> %{x}<br><b>Total:</b> %{y:,}<extra></extra>'
))

# Linha Comex
fig_anual.add_trace(go.Scatter(
    x=df_grafico_anual['ano'],
    y=df_grafico_anual['comex'],
    mode='lines+markers',
    name='Comex',
    line=dict(color='#2ca02c', width=3),
    marker=dict(size=8),
    hovertemplate='<b>Ano:</b> %{x}<br><b>Comex:</b> %{y:,}<extra></extra>'
))

# Linha Não-Comex
fig_anual.add_trace(go.Scatter(
    x=df_grafico_anual['ano'],
    y=df_grafico_anual['nao_comex'],
    mode='lines+markers',
    name='Não-Comex',
    line=dict(color='#ff7f0e', width=3),
    marker=dict(size=8),
    hovertemplate='<b>Ano:</b> %{x}<br><b>Não-Comex:</b> %{y:,}<extra></extra>'
))

# Configurar layout do gráfico anual
fig_anual.update_layout(
    title={
        'text': '<b>Crescimento Anual de Empresas - Últimos 10 Anos</b><br><sub>Cotia/SP</sub>',
        'x': 0.5,
        'xanchor': 'center',
        'font': {'size': 20}
    },
    xaxis_title='<b>Ano</b>',
    yaxis_title='<b>Número de Empresas Abertas</b>',
    hovermode='x unified',
    template='plotly_white',
    font=dict(size=12),
    showlegend=True,
    legend=dict(
        orientation="h",
        yanchor="bottom",
        y=1.02,
        xanchor="right",
        x=1,
        font=dict(size=12)
    ),
    autosize=True,
    xaxis=dict(
        dtick=1,
        tickmode='linear'
    ),
    yaxis=dict(
        gridcolor='lightgray',
        gridwidth=0.5
    )
)

# Salvar gráfico anual
fig_anual.write_html("graficos/crescimento_anual_empresas.html")
print("  Gráfico 'crescimento_anual_empresas.html' salvo")

# ==============================
# GRÁFICO 2: EVOLUÇÃO MENSAL (ÚLTIMOS 12 MESES)
# ==============================

print("Gerando gráfico 2/3: Evolução Mensal...")

# Preparar dados para o gráfico mensal
dados_grafico_mensal = []
for mes in meses_unicos:
    dados_mes = ultimos_12_meses.filter(pl.col("ano_mes") == mes)
    
    total_mes = dados_mes.select("quantidade").sum().item()
    
    comex_mes = dados_mes.filter(pl.col("is_comex")).select("quantidade")
    qtd_comex = comex_mes.sum().item() if comex_mes.height > 0 else 0
    
    nao_comex_mes = dados_mes.filter(~pl.col("is_comex")).select("quantidade")
    qtd_nao_comex = nao_comex_mes.sum().item() if nao_comex_mes.height > 0 else 0
    
    # Formatar mês para exibição (ex: "2024-10" -> "Out/24")
    ano_mes_split = mes.split('-')
    mes_num = int(ano_mes_split[1])
    ano_abrev = ano_mes_split[0][2:]
    
    mes_formatado = f"{meses_nomes[mes_num-1]}/{ano_abrev}"
    
    dados_grafico_mensal.append({
        'ano_mes': mes,
        'mes_formatado': mes_formatado,
        'total': total_mes,
        'comex': qtd_comex,
        'nao_comex': qtd_nao_comex
    })

df_grafico_mensal = pd.DataFrame(dados_grafico_mensal)

# Criar gráfico de linhas para evolução mensal
fig_mensal = go.Figure()

# Linha Total
fig_mensal.add_trace(go.Scatter(
    x=df_grafico_mensal['mes_formatado'],
    y=df_grafico_mensal['total'],
    mode='lines+markers',
    name='Total',
    line=dict(color='#1f77b4', width=3),
    marker=dict(size=10),
    hovertemplate='<b>Mês:</b> %{x}<br><b>Total:</b> %{y:,}<extra></extra>'
))

# Linha Comex
fig_mensal.add_trace(go.Scatter(
    x=df_grafico_mensal['mes_formatado'],
    y=df_grafico_mensal['comex'],
    mode='lines+markers',
    name='Comex',
    line=dict(color='#2ca02c', width=3),
    marker=dict(size=10),
    hovertemplate='<b>Mês:</b> %{x}<br><b>Comex:</b> %{y:,}<extra></extra>'
))

# Linha Não-Comex
fig_mensal.add_trace(go.Scatter(
    x=df_grafico_mensal['mes_formatado'],
    y=df_grafico_mensal['nao_comex'],
    mode='lines+markers',
    name='Não-Comex',
    line=dict(color='#ff7f0e', width=3),
    marker=dict(size=10),
    hovertemplate='<b>Mês:</b> %{x}<br><b>Não-Comex:</b> %{y:,}<extra></extra>'
))

# Configurar layout do gráfico mensal
fig_mensal.update_layout(
    title={
        'text': '<b>Evolução Mensal de Abertura de Empresas - Últimos 12 Meses</b><br><sub>Cotia/SP</sub>',
        'x': 0.5,
        'xanchor': 'center',
        'font': {'size': 20}
    },
    xaxis_title='<b>Mês</b>',
    yaxis_title='<b>Número de Empresas Abertas</b>',
    hovermode='x unified',
    template='plotly_white',
    font=dict(size=12),
    showlegend=True,
    legend=dict(
        orientation="h",
        yanchor="bottom",
        y=1.02,
        xanchor="right",
        x=1,
        font=dict(size=12)
    ),
    autosize=True,
    xaxis=dict(
        tickangle=-45
    ),
    yaxis=dict(
        gridcolor='lightgray',
        gridwidth=0.5
    )
)

# Salvar gráfico mensal
fig_mensal.write_html("graficos/evolucao_mensal_empresas.html")
print("  Gráfico 'evolucao_mensal_empresas.html' salvo")

# ==============================
# GRÁFICO 3: CRESCIMENTO ACUMULADO MENSAL (ÚLTIMOS 12 MESES)
# ==============================
print("Gerando gráfico 3/3: Crescimento Acumulado...")

# Criar gráfico com dois eixos Y
fig_crescimento = make_subplots(
    specs=[[{"secondary_y": True}]]
)

# Eixo principal (esquerdo): Número absoluto de empresas
fig_crescimento.add_trace(
    go.Scatter(
        x=df_crescimento_acumulado['mes_formatado'],
        y=df_crescimento_acumulado['total'],
        mode='lines+markers',
        name='Total',
        line=dict(color='#1f77b4', width=3),
        marker=dict(size=8),
        hovertemplate='<b>Mês:</b> %{x}<br><b>Total:</b> %{y:,}<br><extra></extra>'
    ),
    secondary_y=False
)

fig_crescimento.add_trace(
    go.Scatter(
        x=df_crescimento_acumulado['mes_formatado'],
        y=df_crescimento_acumulado['comex'],
        mode='lines+markers',
        name='Comex',
        line=dict(color='#2ca02c', width=3),
        marker=dict(size=8),
        hovertemplate='<b>Mês:</b> %{x}<br><b>Comex:</b> %{y:,}<br><extra></extra>'
    ),
    secondary_y=False
)

fig_crescimento.add_trace(
    go.Scatter(
        x=df_crescimento_acumulado['mes_formatado'],
        y=df_crescimento_acumulado['nao_comex'],
        mode='lines+markers',
        name='Não-Comex',
        line=dict(color='#ff7f0e', width=3),
        marker=dict(size=8),
        hovertemplate='<b>Mês:</b> %{x}<br><b>Não-Comex:</b> %{y:,}<br><extra></extra>'
    ),
    secondary_y=False
)

fig_crescimento.update_layout(
    title={
        'text': '<b>Crescimento acumulado de Empresas - Últimos 12 Meses</b><br><sub>Cotia/SP</sub>',
        'x': 0.5,
        'xanchor': 'center',
        'font': {'size': 20}
    },
    hovermode='x unified',
    template='plotly_white',
    font=dict(size=12),
    showlegend=True,
    legend=dict(
        orientation="h",  # Legendas horizontais
        yanchor="bottom",
        y=1.02,  # Logo abaixo do título
        xanchor="right",  # Centralizada
        x=0.94,  # No centro horizontal
        font=dict(size=12)
    ),
    autosize=True,
    margin=dict(l=80, r=20, t=100, b=80),  # Margem superior ajustada
    xaxis=dict(
        title='<b>Mês</b>',
        tickangle=-45,
        showgrid=True,
        gridcolor='lightgray',
        gridwidth=0.5
    )
)

# Configurar eixos Y
fig_crescimento.update_yaxes(
    title_text="<b>Número de Empresas Abertas</b>",
    secondary_y=False,
    gridcolor='lightgray',
    gridwidth=0.5
)

# Salvar gráfico
fig_crescimento.write_html("graficos/crescimento_acumulado_mensal.html")
print("  Gráfico 'crescimento_acumulado_mensal.html' salvo")


# ==============================
# GRÁFICOS INTERATIVOS - ÍNDICE BASE 100
# ==============================

# Função para calcular índice base 100
def calcular_indice_base_100(series):
    """
    Transforma uma série em índice base 100.
    O primeiro valor vira 100, os demais mostram variação percentual.
    """
    if len(series) == 0 or series.iloc[0] == 0:
        return pd.Series([100] * len(series))
    
    valor_base = series.iloc[0]
    return (series / valor_base) * 100

# ==============================
# GRÁFICO 1: CRESCIMENTO ANUAL (ÚLTIMOS 10 ANOS) - Índice Base 100
# ==============================

print("Gerando gráfico 1/3: Crescimento Anual (Índice Base 100)...")

# Preparar dados para o gráfico anual
dados_grafico_anual = []
for ano in anos_unicos:
    dados_ano = crescimento_anual.filter(pl.col("ano") == ano)
    
    total_ano = dados_ano.select("quantidade").sum().item()
    
    comex_ano = dados_ano.filter(pl.col("is_comex")).select("quantidade")
    qtd_comex = comex_ano.sum().item() if comex_ano.height > 0 else 0
    
    nao_comex_ano = dados_ano.filter(~pl.col("is_comex")).select("quantidade")
    qtd_nao_comex = nao_comex_ano.sum().item() if nao_comex_ano.height > 0 else 0
    
    dados_grafico_anual.append({
        'ano': ano,
        'total': total_ano,
        'comex': qtd_comex,
        'nao_comex': qtd_nao_comex
    })

df_grafico_anual = pd.DataFrame(dados_grafico_anual)

# Calcular índices base 100
df_grafico_anual['total_indice'] = calcular_indice_base_100(df_grafico_anual['total'])
df_grafico_anual['comex_indice'] = calcular_indice_base_100(df_grafico_anual['comex'])
df_grafico_anual['nao_comex_indice'] = calcular_indice_base_100(df_grafico_anual['nao_comex'])

fig_anual = go.Figure()

# Linha Total (índice)
fig_anual.add_trace(go.Scatter(
    x=df_grafico_anual['ano'],
    y=df_grafico_anual['total_indice'],
    mode='lines+markers',
    name='Total',
    line=dict(color='#1f77b4', width=3),
    marker=dict(size=8),
    customdata=list(zip(df_grafico_anual['total'], df_grafico_anual['total_indice'])),
    hovertemplate='<b>Ano:</b> %{x}<br>' +
                  '<b>Total:</b> %{customdata[0]:,} empresas<br>' +
                  '<b>Índice:</b> %{customdata[1]:.1f}<br>' +
                  '<extra></extra>'
))

# Linha Comex (índice)
fig_anual.add_trace(go.Scatter(
    x=df_grafico_anual['ano'],
    y=df_grafico_anual['comex_indice'],
    mode='lines+markers',
    name='Comex',
    line=dict(color='#2ca02c', width=3),
    marker=dict(size=8),
    customdata=list(zip(df_grafico_anual['comex'], df_grafico_anual['comex_indice'])),
    hovertemplate='<b>Ano:</b> %{x}<br>' +
                  '<b>Comex:</b> %{customdata[0]:,} empresas<br>' +
                  '<b>Índice:</b> %{customdata[1]:.1f}<br>' +
                  '<extra></extra>'
))

# Linha Não-Comex (índice)
fig_anual.add_trace(go.Scatter(
    x=df_grafico_anual['ano'],
    y=df_grafico_anual['nao_comex_indice'],
    mode='lines+markers',
    name='Não-Comex',
    line=dict(color='#ff7f0e', width=3),
    marker=dict(size=8),
    customdata=list(zip(df_grafico_anual['nao_comex'], df_grafico_anual['nao_comex_indice'])),
    hovertemplate='<b>Ano:</b> %{x}<br>' +
                  '<b>Não-Comex:</b> %{customdata[0]:,} empresas<br>' +
                  '<b>Índice:</b> %{customdata[1]:.1f}<br>' +
                  '<extra></extra>'
))

# Adicionar linha de referência em 100
fig_anual.add_hline(
    y=100, 
    line_dash="dash", 
    line_color="gray", 
    opacity=0.5,
    annotation_text="Base (100)",
    annotation_position="right"
)

# Configurar layout do gráfico anual
fig_anual.update_layout(
    title={
        'text': '<b>Crescimento Anual de Empresas - Índice Base 100</b><br>' +
                f'<sub>Base: {df_grafico_anual["ano"].iloc[0]} = 100 | Cotia/SP</sub>',
        'x': 0.5,
        'xanchor': 'center',
        'font': {'size': 20}
    },
    xaxis_title='<b>Ano</b>',
    yaxis_title='<b>Índice (Base 100)</b>',
    hovermode='x unified',
    template='plotly_white',
    font=dict(size=12),
    showlegend=True,
    legend=dict(
        orientation="h",
        yanchor="bottom",
        y=1.02,
        xanchor="right",
        x=1,
        font=dict(size=12)
    ),
    height=500,
    xaxis=dict(
        dtick=1,
        tickmode='linear'
    ),
    yaxis=dict(
        gridcolor='lightgray',
        gridwidth=0.5
    )
)

# Salvar gráfico anual
fig_anual.write_html("graficos/crescimento_anual_empresas_indice100.html")
print("  Gráfico 'crescimento_anual_empresas_indice100.html' salvo")

# ==============================
# GRÁFICO 2: EVOLUÇÃO MENSAL (ÚLTIMOS 12 MESES) - Índice Base 100
# ==============================

print("Gerando gráfico 2/3: Evolução Mensal (Índice Base 100)...")

dados_grafico_mensal = []
for mes in meses_unicos:
    dados_mes = ultimos_12_meses.filter(pl.col("ano_mes") == mes)
    
    total_mes = dados_mes.select("quantidade").sum().item()
    
    comex_mes = dados_mes.filter(pl.col("is_comex")).select("quantidade")
    qtd_comex = comex_mes.sum().item() if comex_mes.height > 0 else 0
    
    nao_comex_mes = dados_mes.filter(~pl.col("is_comex")).select("quantidade")
    qtd_nao_comex = nao_comex_mes.sum().item() if nao_comex_mes.height > 0 else 0
    
    # Formatar mês para exibição
    ano_mes_split = mes.split('-')
    mes_num = int(ano_mes_split[1])
    ano_abrev = ano_mes_split[0][2:]
    
    mes_formatado = f"{meses_nomes[mes_num-1]}/{ano_abrev}"
    
    dados_grafico_mensal.append({
        'ano_mes': mes,
        'mes_formatado': mes_formatado,
        'total': total_mes,
        'comex': qtd_comex,
        'nao_comex': qtd_nao_comex
    })

df_grafico_mensal = pd.DataFrame(dados_grafico_mensal)

# Calcular índices base 100
df_grafico_mensal['total_indice'] = calcular_indice_base_100(df_grafico_mensal['total'])
df_grafico_mensal['comex_indice'] = calcular_indice_base_100(df_grafico_mensal['comex'])
df_grafico_mensal['nao_comex_indice'] = calcular_indice_base_100(df_grafico_mensal['nao_comex'])

fig_mensal = go.Figure()

# Total
fig_mensal.add_trace(go.Scatter(
    x=df_grafico_mensal['mes_formatado'],
    y=df_grafico_mensal['total_indice'],
    mode='lines+markers',
    name='Total',
    line=dict(color='#1f77b4', width=3),
    marker=dict(size=10),
    customdata=list(zip(df_grafico_mensal['total'], df_grafico_mensal['total_indice'])),
    hovertemplate='<b>Mês:</b> %{x}<br>' +
                  '<b>Total:</b> %{customdata[0]:,} empresas<br>' +
                  '<b>Índice:</b> %{customdata[1]:.1f}<br>' +
                  '<extra></extra>'
))

# Comex
fig_mensal.add_trace(go.Scatter(
    x=df_grafico_mensal['mes_formatado'],
    y=df_grafico_mensal['comex_indice'],
    mode='lines+markers',
    name='Comex',
    line=dict(color='#2ca02c', width=3),
    marker=dict(size=10),
    customdata=list(zip(df_grafico_mensal['comex'], df_grafico_mensal['comex_indice'])),
    hovertemplate='<b>Mês:</b> %{x}<br>' +
                  '<b>Comex:</b> %{customdata[0]:,} empresas<br>' +
                  '<b>Índice:</b> %{customdata[1]:.1f}<br>' +
                  '<extra></extra>'
))

# Não-Comex
fig_mensal.add_trace(go.Scatter(
    x=df_grafico_mensal['mes_formatado'],
    y=df_grafico_mensal['nao_comex_indice'],
    mode='lines+markers',
    name='Não-Comex',
    line=dict(color='#ff7f0e', width=3),
    marker=dict(size=10),
    customdata=list(zip(df_grafico_mensal['nao_comex'], df_grafico_mensal['nao_comex_indice'])),
    hovertemplate='<b>Mês:</b> %{x}<br>' +
                  '<b>Não-Comex:</b> %{customdata[0]:,} empresas<br>' +
                  '<b>Índice:</b> %{customdata[1]:.1f}<br>' +
                  '<extra></extra>'
))

# Linha de referência
fig_mensal.add_hline(
    y=100, 
    line_dash="dash", 
    line_color="gray", 
    opacity=0.5,
    annotation_text="Base (100)",
    annotation_position="right"
)

# Layout
fig_mensal.update_layout(
    title={
        'text': '<b>Evolução Mensal de Abertura de Empresas - Índice Base 100</b><br>' +
                f'<sub>Base: {df_grafico_mensal["mes_formatado"].iloc[0]} = 100 | Cotia/SP</sub>',
        'x': 0.5,
        'xanchor': 'center',
        'font': {'size': 20}
    },
    xaxis_title='<b>Mês</b>',
    yaxis_title='<b>Índice (Base 100)</b>',
    hovermode='x unified',
    template='plotly_white',
    font=dict(size=12),
    showlegend=True,
    legend=dict(
        orientation="h",
        yanchor="bottom",
        y=1.02,
        xanchor="right",
        x=1,
        font=dict(size=12)
    ),
    autosize=True,
    xaxis=dict(
        tickangle=-45
    ),
    yaxis=dict(
        gridcolor='lightgray',
        gridwidth=0.5
    )
)

# Salvar
fig_mensal.write_html("graficos/evolucao_mensal_empresas_indice100.html")
print("  Gráfico 'evolucao_mensal_empresas_indice100.html' salvo")

# ==============================
# GRÁFICO 3: CRESCIMENTO ACUMULADO MENSAL - Índice Base 100
# ==============================

print("Gerando gráfico 3/3: Crescimento Acumulado (Índice Base 100)...")

# Copiar e preparar para plot
df_cres = df_crescimento_acumulado.copy()

# Calcular índices base 100
df_cres['total_indice'] = calcular_indice_base_100(df_cres['total'])
df_cres['comex_indice'] = calcular_indice_base_100(df_cres['comex'])
df_cres['nao_comex_indice'] = calcular_indice_base_100(df_cres['nao_comex'])

fig_crescimento = go.Figure()

# Total
fig_crescimento.add_trace(go.Scatter(
    x=df_cres['mes_formatado'],
    y=df_cres['total_indice'],
    mode='lines+markers',
    name='Total de Empresas',
    line=dict(color='#1f77b4', width=3),
    marker=dict(size=10),
    customdata=list(zip(df_cres['total'], df_cres['total_indice'], df_cres['crescimento_total_pct'])),
    hovertemplate='<b>Mês:</b> %{x}<br>' +
                  '<b>Total:</b> %{customdata[0]:,} empresas<br>' +
                  '<b>Índice:</b> %{customdata[1]:.1f}<br>' +
                  '<b>Crescimento:</b> %{customdata[2]:.2f}%<br>' +
                  '<extra></extra>'
))

# Comex
fig_crescimento.add_trace(go.Scatter(
    x=df_cres['mes_formatado'],
    y=df_cres['comex_indice'],
    mode='lines+markers',
    name='Empresas Comex',
    line=dict(color='#2ca02c', width=3),
    marker=dict(size=10),
    customdata=list(zip(df_cres['comex'], df_cres['comex_indice'])),
    hovertemplate='<b>Mês:</b> %{x}<br>' +
                  '<b>Comex:</b> %{customdata[0]:,} empresas<br>' +
                  '<b>Índice:</b> %{customdata[1]:.1f}<br>' +
                  '<extra></extra>'
))

# Não-Comex
fig_crescimento.add_trace(go.Scatter(
    x=df_cres['mes_formatado'],
    y=df_cres['nao_comex_indice'],
    mode='lines+markers',
    name='Empresas Não-Comex',
    line=dict(color='#ff7f0e', width=3),
    marker=dict(size=10),
    customdata=list(zip(df_cres['nao_comex'], df_cres['nao_comex_indice'])),
    hovertemplate='<b>Mês:</b> %{x}<br>' +
                  '<b>Não-Comex:</b> %{customdata[0]:,} empresas<br>' +
                  '<b>Índice:</b> %{customdata[1]:.1f}<br>' +
                  '<extra></extra>'
))

# Linha de referência
fig_crescimento.add_hline(
    y=100, 
    line_dash="dash", 
    line_color="gray", 
    opacity=0.5,
    annotation_text="Base (100)",
    annotation_position="right"
)

# Layout
fig_crescimento.update_layout(
    title={
        'text': '<b>Crescimento Acumulado de Empresas - Índice Base 100</b><br>' +
                f'<sub>Base: {df_cres["mes_formatado"].iloc[0]} = 100 | Cotia/SP</sub>',
        'x': 0.5,
        'xanchor': 'center',
        'font': {'size': 20}
    },
    xaxis_title='<b>Mês</b>',
    yaxis_title='<b>Índice (Base 100)</b>',
    hovermode='x unified',
    template='plotly_white',
    font=dict(size=12),
    showlegend=True,
    legend=dict(
        orientation="v",
        yanchor="top",
        y=1,
        xanchor="left",
        x=1.05,
        font=dict(size=11)
    ),
    autosize=True,
    xaxis=dict(
        tickangle=-45
    ),
    yaxis=dict(
        gridcolor='lightgray',
        gridwidth=0.5
    )
)

# Salvar gráfico
fig_crescimento.write_html("graficos/crescimento_acumulado_mensal_indice100.html")
print("  Gráfico 'crescimento_acumulado_mensal_indice100.html' salvo")

print("\nTodos os gráficos gerados com Índice Base 100!")
print("\n Como interpretar:")
print("   - Todas as linhas começam em 100 (valor base)")
print("   - Índice 150 = crescimento de 50% em relação ao início")
print("   - Índice 80 = redução de 20% em relação ao início")
print("   - Permite comparar taxas de crescimento independente dos valores absolutos")



# ==============================
# GRÁFICOS - ANÁLISE GEOGRÁFICA
# ==============================

import plotly.graph_objects as go

print("\n" + "=" * 80)
print("GERANDO GRÁFICOS DE ANÁLISE GEOGRÁFICA...")
print("=" * 80 + "\n")

# ==============================
# GRÁFICO 4: TOP 20 BAIRROS - COMPARATIVO
# ==============================

print("Gerando gráfico 4: Top 20 Bairros (Comparativo)...")

# Preparar dados
top_20_df = top_20_geral.to_pandas()

fig_bairros = go.Figure()

# Barra Total de Empresas
fig_bairros.add_trace(go.Bar(
    y=top_20_df['bairro_padronizado'],
    x=top_20_df['total_empresas'],
    name='Total de Empresas',
    orientation='h',
    marker=dict(color='#1f77b4'),
    hovertemplate='<b>%{y}</b><br>Total: %{x:,}<extra></extra>'
))

# Barra Empresas Comex
fig_bairros.add_trace(go.Bar(
    y=top_20_df['bairro_padronizado'],
    x=top_20_df['total_comex'],
    name='Empresas Comex',
    orientation='h',
    marker=dict(color='#2ca02c'),
    hovertemplate='<b>%{y}</b><br>Comex: %{x:,}<extra></extra>'
))

fig_bairros.update_layout(
    title={
        'text': '<b>Top 20 Bairros com Mais Empresas</b><br><sub>Comparativo: Total vs Comex - Cotia/SP</sub>',
        'x': 0.5,
        'xanchor': 'center',
        'font': {'size': 20}
    },
    xaxis_title='<b>Número de Empresas</b>',
    yaxis_title='<b>Bairro</b>',
    barmode='group',
    template='plotly_white',
    font=dict(size=11),
    showlegend=True,
    legend=dict(
        orientation="h",
        yanchor="bottom",
        y=1.02,
        xanchor="right",
        x=1
    ),
    autosize=True,
    yaxis=dict(autorange="reversed"),
    hovermode='y unified'
)

fig_bairros.write_html("graficos/top_20_bairros_comparativo.html")
print("  Gráfico 'top_20_bairros_comparativo.html' salvo")

# ==============================
# GRÁFICO 5: TOP 15 BAIRROS - PERCENTUAL COMEX
# ==============================

print("Gerando gráfico 5: Bairros com Maior Concentração Comex...")

# Preparar dados
concentracao_df = bairros_concentracao.to_pandas()

fig_concentracao = go.Figure()

# Barra de percentual
fig_concentracao.add_trace(go.Bar(
    y=concentracao_df['bairro_padronizado'],
    x=concentracao_df['percentual_comex'],
    orientation='h',
    marker=dict(
        color=concentracao_df['percentual_comex'],
        colorscale='Viridis',
        showscale=True,
        colorbar=dict(title="% Comex")
    ),
    text=concentracao_df['percentual_comex'].round(1).astype(str) + '%',
    textposition='outside',
    hovertemplate='<b>%{y}</b><br>Percentual Comex: %{x:.1f}%<br>Comex: %{customdata[0]}<br>Total: %{customdata[1]}<extra></extra>',
    customdata=concentracao_df[['total_comex', 'total_empresas']]
))

fig_concentracao.update_layout(
    title={
        'text': '<b>Top 15 Bairros com Maior Concentração de Comex</b><br><sub>Percentual de empresas Comex - Cotia/SP</sub>',
        'x': 0.5,
        'xanchor': 'center',
        'font': {'size': 20}
    },
    xaxis_title='<b>Percentual de Empresas Comex (%)</b>',
    yaxis_title='<b>Bairro</b>',
    template='plotly_white',
    font=dict(size=11),
    showlegend=False,
    autosize=True,
    yaxis=dict(autorange="reversed"),
    xaxis=dict(range=[0, concentracao_df['percentual_comex'].max() * 1.15])
)

fig_concentracao.write_html("graficos/bairros_concentracao_comex.html")
print("  Gráfico 'bairros_concentracao_comex.html' salvo")

# ==============================
# GRÁFICO 6: TOP 10 BAIRROS COMEX - BARRAS HORIZONTAIS (substitui pizza)
# ==============================

print("Gerando gráfico 6: Top 10 Bairros com Mais Empresas Comex (Barras)...")

# Preparar dados - Top 10 bairros com mais empresas Comex
top_10_comex_df = distribuicao_comex_ordenada.head(10).to_pandas()

# OPCIONAL: incluir "OUTROS"?
incluir_outros = False  # Mude para True se quiser incluir

if incluir_outros:
    total_comex_com_bairro = (
        distribuicao_bairros_comex.select(pl.col("total_comex")).sum().item()
        if "distribuicao_bairros_comex" in globals()
        else distribuicao_completa.select(pl.col("total_comex")).sum().item()
    )
    top_10_sum = int(top_10_comex_df['total_comex'].sum())
    outros_comex = max(0, int(total_comex_com_bairro) - top_10_sum)

    # Adicionar "OUTROS"
    bairros_lista = top_10_comex_df['bairro_padronizado'].tolist() + ['OUTROS']
    valores_lista = top_10_comex_df['total_comex'].tolist() + [outros_comex]
else:
    # Usar apenas Top 10
    bairros_lista = top_10_comex_df['bairro_padronizado'].tolist()
    valores_lista = top_10_comex_df['total_comex'].tolist()

# Verificar se há dados
if sum(valores_lista) == 0:
    print("Aviso: total de empresas COMEX é zero — gráfico de barras não será gerado.")
else:
    # Criar gráfico de barras horizontais
    fig_barras = go.Figure()

    fig_barras.add_trace(go.Bar(
        y=bairros_lista[::-1],  # Inverter para o maior ficar no topo
        x=valores_lista[::-1],
        orientation='h',
        marker=dict(
            color='#2ca02c',  # Cor consistente com Comex
            line=dict(color='#1a5f1a', width=0.5)
        ),
        text=valores_lista[::-1],
        textposition='auto',
        hovertemplate='<b>%{y}</b><br>Empresas Comex: %{x:,}<extra></extra>'
    ))

    fig_barras.update_layout(
        title={
            'text': '<b>Top 10 Bairros com Mais Empresas de Comex</b><br><sub>Cotia/SP</sub>',
            'x': 0.5,
            'xanchor': 'center',
            'font': {'size': 20}
        },
        xaxis_title='<b>Número de Empresas Comex</b>',
        yaxis_title='<b>Bairro</b>',
        template='plotly_white',
        font=dict(size=12),
        autosize=True,
        margin=dict(l=150, r=50, t=100, b=50),  # Margem extra para nomes longos
    )

    out_path = "graficos/top_10_bairros_comex_barras.html"
    fig_barras.write_html(out_path)
    print(f"  Gráfico '{out_path}' salvo")


# ==============================
# GERAR ARQUIVO HTML DO DASHBOARD COM DADOS ATUALIZADOS
# ==============================
print("Gerando arquivo dashboard html...")
# Formatar números com ponto (ex: 46892 -> "46.892")
def fmt(n):
    return f"{n:,}".replace(",", ".")

# Calcular todas as métricas (você já tem essas variáveis)
stats = {
    "total_empresas": total_empresas,
    "total_comex": total_comex,
    "percentual_comex": percentual_comex,
    "total_nao_comex": total_nao_comex,
    "comex_principal": comex_principal,
    "comex_secundaria": comex_secundaria,
    "pct_principal": (comex_principal / total_comex * 100) if total_comex > 0 else 0,
    "pct_secundaria": (comex_secundaria / total_comex * 100) if total_comex > 0 else 0,
    "empresas_com_data": empresas_com_data,
    "empresas_sem_data": empresas_sem_data,
    "idade_media_geral": formatar_idade(idade_media_geral),
    "idade_mediana_geral": formatar_idade(idade_mediana_geral),
    "idade_max": formatar_idade(idade_max),
    "idade_min": formatar_idade(idade_min),
    "idade_comex_media": formatar_idade(idade_comex_media),
    "idade_comex_mediana": formatar_idade(idade_comex_mediana),
    "idade_comex_max": formatar_idade(idade_comex_max),
    "idade_comex_min": formatar_idade(idade_comex_min),
    "crescimento_total_periodo": crescimento_total_periodo,
    "total_inicial": df_crescimento_acumulado['total'].iloc[0],
    "total_final": df_crescimento_acumulado['total'].iloc[-1],
    "dif_total": df_crescimento_acumulado['total'].iloc[-1] - df_crescimento_acumulado['total'].iloc[0],
    "crescimento_comex_periodo": crescimento_comex_periodo,
    "comex_inicial": df_crescimento_acumulado['comex'].iloc[0],
    "comex_final": df_crescimento_acumulado['comex'].iloc[-1],
    "dif_comex": df_crescimento_acumulado['comex'].iloc[-1] - df_crescimento_acumulado['comex'].iloc[0],
}

metodologia_html = {
    "cnaes_aceitos": formatar_lista_para_html(lista_cnaes_aceitos),
    "cnaes_rejeitados": formatar_lista_para_html(lista_cnaes_rejeitados),
    "palavras_chave": formatar_lista_para_html(lista_palavras),
    "total_aceitos": len(lista_cnaes_aceitos),
    "total_rejeitados": len(lista_cnaes_rejeitados),
    "total_palavras": len(lista_palavras),
    "mes": mes_para_baixar
}

# Template HTML (com placeholders)
html_template = f"""
<!DOCTYPE html>
<html lang="pt-BR">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0"/>
  <title>Dashboard - Empresas Comex em Cotia</title>
  <link rel="stylesheet" href="style.css">
</head>
<body>
  <div class="container">
    <header>
  <div class="header-content">
    <a href="https://fateccotia.cps.sp.gov.br/">
      <img src="https://bkpsitecpsnew.blob.core.windows.net/uploadsitecps/sites/115/2024/07/logo-fatec_cotia.png" 
           alt="Logo Fatec Cotia"
           class="logo">
      <div class="header-text">
    </a>
      <h1>Observatório Comex Cotia</h1>
      <p>{metodologia_html["mes"]}</>
    </div>
  </div>
</header>

    <!-- SEÇÃO: Análise Estatística -->
    <!--<h2>Análise Estatística</h2>-->
    
    <div class="highlight">
      <h3>Resumo Geral</h3>
      <div class="metric-card">
        <strong>Total de empresas ativas em Cotia:</strong> {fmt(stats['total_empresas'])}<br>
        <strong>Total de empresas com atividades passíveis de comex:</strong> {fmt(stats['total_comex'])}<br>
        <strong>Percentual de empresas com atividade passíveis de comex:</strong> {stats['percentual_comex']:.2f}%<br>
        <strong>Total de empresas sem atividade Comex:</strong> {fmt(stats['total_nao_comex'])}
      </div>
    </div>

    <div class="highlight">
      <h3>Comex como Atividade Principal vs Secundária</h3>
      <div class="metric-card">
        <strong>Empresas com Comex como atividade PRINCIPAL:</strong> {fmt(stats['comex_principal'])} ({stats['pct_principal']:.1f}%)<br>
        <strong>Empresas com Comex como atividade SECUNDÁRIA:</strong> {fmt(stats['comex_secundaria'])} ({stats['pct_secundaria']:.1f}%)
      </div>
    </div>

    <div class="highlight">
      <h3>Tempo Médio de Atividade</h3>

      <div class="age-comparison">
        <div class="age-card">
          <h4>Geral (Todas as empresas)</h4>
          <div class="metric-card">
            <strong>Idade média:</strong> {stats['idade_media_geral']}<br>
            <strong>Idade mediana:</strong> {stats['idade_mediana_geral']}<br>
            <strong>Empresa mais antiga:</strong> {stats['idade_max']}<br>
            <strong>Empresa mais nova:</strong> {stats['idade_min']}
          </div>
        </div>

        <div class="age-card">
          <h4>Empresas COMEX</h4>
          <div class="metric-card">
            <strong>Idade média:</strong> {stats['idade_comex_media']}<br>
            <strong>Idade mediana:</strong> {stats['idade_comex_mediana']}<br>
            <strong>Empresa mais antiga:</strong> {stats['idade_comex_max']}<br>
            <strong>Empresa mais nova:</strong> {stats['idade_comex_min']}
          </div>
        </div>
      </div>
    </div>

    <div class="highlight">
      <h3>Crescimento Acumulado (Últimos 12 Meses)</h3>

      <div class="growth-comparison">
        <div class="growth-card">
          <h4>Geral (Todas as empresas)</h4>
          <div class="metric-card">
            <strong>Crescimento:</strong> {stats['crescimento_total_periodo']:+.2f}%<br>
            De {fmt(stats['total_inicial'])} para {fmt(stats['total_final'])} empresas<br>
            Diferença: {fmt(stats['dif_total'])} empresas
          </div>
        </div>

        <div class="growth-card">
          <h4>Empresas COMEX</h4>
          <div class="metric-card comex-highlight">
            <strong>Crescimento:</strong> {stats['crescimento_comex_periodo']:+.2f}%<br>
            De {fmt(stats['comex_inicial'])} para {fmt(stats['comex_final'])} empresas<br>
            Diferença: {fmt(stats['dif_comex'])} empresas
          </div>
        </div>
      </div>
    </div>

    <!-- Gráficos -->
    <h2>Gráficos Temporais</h2>
    <div class="plot-container">
      <iframe src="graficos/crescimento_anual_empresas.html"></iframe>
    </div>
    <div class="plot-container">
      <iframe src="graficos/crescimento_acumulado_mensal.html"></iframe>
    </div>
    <div class="plot-container">
      <iframe src="graficos/evolucao_mensal_empresas.html"></iframe>
    </div>
    <hr class="linha">
    <div class="plot-container">
      <iframe src="graficos/crescimento_anual_empresas_indice100.html"></iframe>
    </div>
    <div class="plot-container">
      <iframe src="graficos/crescimento_acumulado_mensal_indice100.html"></iframe>
    </div>
    <div class="plot-container">
      <iframe src="graficos/evolucao_mensal_empresas_indice100.html"></iframe>
    </div>
    
    <h2>Análise Geográfica</h2>
    <div class="plot-container">
      <iframe src="graficos/top_20_bairros_comparativo.html"></iframe>
    </div>
    <div class="plot-container">
      <iframe src="graficos/bairros_concentracao_comex.html"></iframe>
    </div>
    <div class="plot-container">
      <iframe src="graficos/top_10_bairros_comex_barras.html"></iframe>
    </div>
    <h2>Metodologia</h2>

<div class="methodology-section">
  <p>
  Base de dados de utilizada disponível em: <a href="https://dados.gov.br/dados/conjuntos-dados/cadastro-nacional-da-pessoa-juridica---cnpj">Cadastro Nacional da Pessoa Jurídica - CNPJ</a>
  </p>
  <p>
    Este dashboard analisa empresas passíveis de Comércio Exterior em Cotia/SP com status de atividade "Aberta" no mês {metodologia_html["mes"]}, com base em dois critérios complementares:
    atividade econômica (CNAE) e palavras-chave em razão social ou nome fantasia.<br></br>
    A classificação é feita usando um arquivo de referência com regras manuais para garantir precisão e evitar falsos positivos.
  </p>
  
  <h3>Critérios de classificação</h3>
  <p>As empresas são consideradas Comex se atenderem ao menos um dos critérios abaixo e não estiverem em categorias rejeitadas:</p>

  <div class="criteria-grid">
    <div class="criterion accept">
      <h4>Aceita</h4>
      <p>Empresas com CNAE marcado como <strong>"Aceita"</strong> <strong>OU</strong> que contenham <strong>palavras-chave</strong> em razão social/nome fantasia.</p>
      <p><em>Exemplo de palavra-chave: "exportação", "importação", "internacional", "trade", etc.</em></p>
    </div>

    <div class="criterion reject">
      <h4>Rejeita</h4>
      <p>Empresas com CNAE marcado como <strong>"Rejeita"</strong> são <strong>excluídas</strong>, mesmo que atendam aos critérios positivos. Evitar falsos positivos.</p>
      <p><em>Exemplo real: empresa com nome "Igreja Internacional" (contém "internacional") mas CNAE religioso → <strong>rejeitada</strong>.</em></p>
    </div>

    <div class="criterion neutral">
      <h4>Ok</h4>
      <p>Empresas com CNAE marcado como <strong>"Ok"</strong> são consideradas <strong>não-Comex</strong> por padrão, a menos que atendam a critérios de nome.</p>
    </div>
  </div>

  <h3>Fluxo de classificação</h3>
  <ol>
    <li>Busca por palavras-chave em <em>razão social</em> e <em>nome fantasia</em> (normalizados: sem acentos, minúsculas);</li>
    <li>Verificação de CNAE principal e CNAEs secundários;</li>
    <li>Aplicação de regras: <strong>ACEITA</strong> se (palavra-chave ou CNAE aceito) e <strong>NÃO</strong> (CNAE rejeitado);</li>
    <li>Resultado: coluna <code>is_comex = True/False</code> no dataset final.</li>
  </ol>

  <h3>Palavras-chave utilizadas na busca</h3>
  <p>Foram usadas <strong>{metodologia_html['total_palavras']}</strong> palavras ou expressões para identificar empresas com atividade passíveis de comércio exterior:</p>
  <div class="criteria-list">{metodologia_html['palavras_chave']}</div>

  <h3>CNAEs considerados como <span class="tag accept">Aceita</span></h3>
  <p><strong>{metodologia_html['total_aceitos']}</strong> classes CNAE foram marcadas como atividade de Comex:</p>
  <div class="criteria-list">{metodologia_html['cnaes_aceitos']}</div>

  <h3>CNAEs marcados como <span class="tag reject">Rejeita</span></h3>
  <p><strong>{metodologia_html['total_rejeitados']}</strong> classes CNAE foram excluídas para evitar falsos positivos (ex: instituições religiosas, entidades sem fins lucrativos):</p>
  <div class="criteria-list">{metodologia_html['cnaes_rejeitados']}</div>
</div>
</div>
  </div>
</body>
</html>
"""

# Salvar o HTML
with open("dashboard.html", "w", encoding="utf-8") as f:
    f.write(html_template)

print("\nDashboard HTML gerado com sucesso: dashboard.html")


# Exportação dos arquivos

# se for necessario exportar os dados para o power bi
df_final.write_excel(f"outputs/final {mes_para_baixar}.xlsx")
print(f"Base final do mês {mes_para_baixar} salva em outputs")
print("FIM")
# df_cnae.write_csv("outputs/relacao_cnae.csv")
# df_unico.to_excel("outputs/frequencia_de_cnaes.xlsx")#to_excel pois é pandas
# df_comex.write_excel("outputs/empresas comex - razao_social, nome_fantasia, cnae.xlsx")#write_excel pois é polars