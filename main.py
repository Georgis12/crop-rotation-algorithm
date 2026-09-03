# -*- coding: utf-8 -*-
import pandas as pd
import time
import sys
import io

# Corrigir codificação para UTF-8 (compatível com vários ambientes)
try:
    # Python 3.7+ permite reconfigure do stream
    sys.stdout.reconfigure(encoding='utf-8')
except Exception:
    # fallback: apenas rewrap se houver atributo 'buffer'
    try:
        buf = getattr(sys.stdout, 'buffer', None)
        if buf is not None:
            sys.stdout = io.TextIOWrapper(buf, encoding='utf-8')
    except Exception:
        # se não for possível, continua com o stdout atual
        pass

EXCEL_PATH = "excel culturas_terraços_regras.xlsx"


def carregar_dados():
    sheets = pd.read_excel(EXCEL_PATH, sheet_name=None)

    print("Abas encontradas no Excel (antes de limpar):")
    for nome in sheets.keys():
        print(f"  Processando aba: {repr(nome)}")
        print(" -", repr(nome))

    sheets = {nome.strip(): df for nome, df in sheets.items()}

    print("\nAbas após limpeza:")
    for nome in sheets.keys():
        print(" -", nome)

    return sheets


import unicodedata


def obter_historico(terraco):
    """Retorna o histórico do terraço como lista de nomes de culturas limpas e em lower-case."""
    val = terraco.get("histórico_últimos_anos", "")
    if not isinstance(val, str):
        return []
    return [normalizar_nome_cultura(h) for h in val.split(",") if h.strip()]


def remover_acentos(texto):
    if not isinstance(texto, str):
        return ""
    texto = unicodedata.normalize("NFKD", texto)
    return "".join(c for c in texto if unicodedata.category(c) != "Mn")


def normalizar_nome_cultura(nome):
    if not isinstance(nome, str):
        return ""
    return remover_acentos(nome.strip().lower())


def encontrar_cultura_por_nome(culturas, nome):
    nome = normalizar_nome_cultura(nome)
    for c in culturas:
        if normalizar_nome_cultura(c.get("nome_da_cultura", "")) == nome:
            return c
    return None


def historico_para_familias(culturas, historico):
    familias = []
    for item in historico:
        cultura = encontrar_cultura_por_nome(culturas, item)
        if cultura is not None:
            familias.append(normalizar_nome_cultura(cultura.get("família_botânica", "")))
        else:
            familias.append(normalizar_nome_cultura(item))
    return familias


def historico_para_tipos_biodinamicos(culturas, historico):
    tipos = []
    for item in historico:
        cultura = encontrar_cultura_por_nome(culturas, item)
        if cultura is not None:
            tipos.append(normalizar_nome_cultura(cultura.get("tipo_biodinâmico", "")))
        else:
            tipos.append(normalizar_nome_cultura(item))
    return tipos


def parse_efeito_solo(cultura):
    """Retorna set de efeitos normalizados da cultura no solo."""
    raw = cultura.get("efeito_no_solo", "")
    if not isinstance(raw, str) or not raw.strip():
        return set()
    efeitos = set()
    for parte in raw.split(","):
        norm = remover_acentos(parte.strip().lower().replace("-", "_").replace(" ", "_"))
        if norm:
            efeitos.add(norm)
    return efeitos


def parse_exigencia(cultura):
    """Retorna nível de exigência nutricional: 'alta', 'media_alta', 'media', 'baixa_media', 'baixa'."""
    raw = cultura.get("exigência_nutrientes", "")
    if not isinstance(raw, str) or not raw.strip():
        return "media"
    norm = remover_acentos(raw.strip().lower().replace("-", "_"))
    mapa = {
        "alta": "alta",
        "media_alta": "media_alta",
        "media": "media",
        "baixa_media": "baixa_media",
        "baixa": "baixa",
    }
    return mapa.get(norm, "media")


def limpar_df(df):
    df = df.astype(str)
    for col in df.columns:
        print(f"    [LIMPAR] Coluna: {col}")
        df[col] = df[col].str.strip()
    return df


def separar_tabelas_regras(df):
    df = limpar_df(df)

    # As tabelas que realmente existem no teu Excel
    tabelas = [
        "Regras_Familias",
        "Incompatibilidade_Global",
        "Compatibilidade_Global"
    ]

    indices = {}

    # Detectar tabelas que têm título
    for tabela in tabelas:
        print(f"  [DETECTAR] Procurando tabela: {tabela}")
        for i in range(len(df)):
            if df.iloc[i, 0] == tabela:
                indices[tabela] = i
                print(f"    [LINHA {i}] ✓ Encontrada: {tabela}")
                break

    # A primeira tabela NÃO tem título → começa na linha 0
    indices["Regras_Gerais"] = 0

    # Ordenar pela posição
    tabelas_ordenadas = sorted(indices.items(), key=lambda x: x[1])

    resultados = {}

    for idx, (tabela, linha_inicio) in enumerate(tabelas_ordenadas):
        print(f"  [EXTRAIR] Processando tabela: '{tabela}' (linha {linha_inicio})")

        # Cabeçalho está na linha seguinte
        header = df.iloc[linha_inicio + 1].tolist()
        print(f"    [CABECALHO] {len(header)} colunas detectadas")

        # Descobrir onde termina
        if idx < len(tabelas_ordenadas) - 1:
            linha_fim = tabelas_ordenadas[idx + 1][1]
        else:
            linha_fim = len(df)

        # Extrair tabela
        tabela_df = df.iloc[linha_inicio + 2 : linha_fim].copy()
        tabela_df.columns = header

        tabela_df = tabela_df.replace("nan", pd.NA).dropna(how="all")
        print(f"    [RESULTADO] {len(tabela_df)} linhas ✓")

        resultados[tabela] = tabela_df

    return resultados

# ============================================================
# FILTROS OBRIGATÓRIOS DO ALGORITMO
# ============================================================

def filtrar_por_familia(cultura, historico, regras_gerais, culturas):
    """
    Evita repetir a mesma família antes de X anos.
    """
    regra = next((r for r in regras_gerais if "nao_repetir_familia" in r.values()), None)
    if not regra:
        return True  # se não achar a regra, não filtra

    anos_minimos = int(regra["valor"])

    # converte o histórico de cultivos para famílias
    historico_familias = historico_para_familias(culturas, historico)
    ultimas = historico_familias[-anos_minimos:]

    # comparar em lower-case para robustez
    familia = str(cultura.get("família_botânica", "")).strip().lower()
    return familia not in ultimas


def filtrar_incompatibilidade_global(cultura, regras_incomp_global):
    """
    Evita culturas globalmente incompatíveis (ex.: fenouil).
    """
    for regra in regras_incomp_global:
        if regra["cultura"].lower() == cultura["nome_da_cultura"].lower():
            return False
    return True


def filtrar_incompatibilidade_especifica(cultura, historico):
    """
    Evita incompatibilidades específicas (ex.: tomate depois de batata).
    """
    if not isinstance(cultura["incompatibilidade"], str):
        return True

    lista = [x.strip().lower() for x in cultura["incompatibilidade"].split(",")]

    for item in lista:
        for h in historico:
            if item in h.lower():
                return False

    return True


def regra_esta_ativa(regras_gerais, chave):
    return any(isinstance(r.get("regra"), str) and chave in r.get("regra") for r in regras_gerais)


def filtrar_familia_seguida(cultura, historico, culturas, familia_alvo):
    if not familia_alvo:
        return True

    familia_alvo = familia_alvo.strip().lower()
    historico_familias = historico_para_familias(culturas, historico)
    if not historico_familias:
        return True

    if cultura.get("família_botânica", "").strip().lower() != familia_alvo:
        return True

    return historico_familias[-1] != familia_alvo


def aplicar_filtros(cultura, terraço, regras_gerais, regras_incomp_global, culturas):
    historico = obter_historico(terraço)

    filtros = [
        filtrar_por_familia(cultura, historico, regras_gerais, culturas),
        filtrar_incompatibilidade_global(cultura, regras_incomp_global),
        filtrar_incompatibilidade_especifica(cultura, historico),
    ]

    if regra_esta_ativa(regras_gerais, "evitar_solanaceas_seguidas"):
        filtros.append(filtrar_familia_seguida(cultura, historico, culturas, "solanaceae"))

    if regra_esta_ativa(regras_gerais, "evitar_brassicas_seguidas"):
        filtros.append(filtrar_familia_seguida(cultura, historico, culturas, "brassicaceae"))

    return all(filtros)

# ============================================================
# SISTEMA DE PONTUAÇÃO
# ============================================================

def pontuar_biodinamica(cultura, historico, culturas=None):
    if len(historico) == 0:
        return 0

    ultimo = normalizar_nome_cultura(historico[-1])
    if culturas is not None:
        cultura_ultimo = encontrar_cultura_por_nome(culturas, ultimo)
        if cultura_ultimo is not None:
            ultimo = normalizar_nome_cultura(cultura_ultimo.get("tipo_biodinâmico", ""))

    tipo = cultura["tipo_biodinâmico"].strip().lower()

    ciclo = {
        "raiz": "folha",
        "folha": "flor",
        "flor": "fruto",
        "fruto": "raiz"
    }

    if ultimo in ciclo and tipo == ciclo[ultimo]:
        return 3  # bom
    return 0


def pontuar_compatibilidade_global(cultura, regras_comp_global):
    nome = cultura["nome_da_cultura"].lower()
    for regra in regras_comp_global:
        if regra["cultura"].lower() == nome:
            return 2
    return 0


def pontuar_compatibilidade_especifica(cultura, historico):
    if not isinstance(cultura["compatibilidade"], str):
        return 0

    comp = [x.strip().lower() for x in cultura["compatibilidade"].split(",")]

    pontos = 0
    for h in historico:
        if h.strip().lower() in comp:
            pontos += 1

    return pontos


def pontuar_diversidade(cultura, historico, culturas=None):
    familia = str(cultura.get("família_botânica", "")).strip().lower()
    if culturas is not None:
        historico_familias = historico_para_familias(culturas, historico)
    else:
        historico_familias = [h.strip().lower() for h in historico if isinstance(h, str)]
    if familia not in historico_familias:
        return 2
    return 0


def pontuar_leguminosa(cultura, historico, regras_gerais, culturas=None):
    # encontrar a regra correta
    regra = next(
        (r for r in regras_gerais if isinstance(r.get("regra"), str) and "inserir_leguminosa_cada" in r.get("regra")),
        None
    )
    if not regra:
        return 0

    try:
        ciclos = int(regra["valor"])
    except:
        return 0

    # se a cultura é leguminosa
    if cultura["família_botânica"].lower() == "fabaceae":
        return 2  # bônus

    if culturas is not None:
        historico_familias = historico_para_familias(culturas, historico)
    else:
        historico_familias = [h.strip().lower() for h in historico if isinstance(h, str)]

    ultimos = historico_familias[-ciclos:]
    if not any(h.lower() == "fabaceae" for h in ultimos):
        return 1

    return 0


def pontuar_solo_nutricional(cultura, historico, culturas):
    """
    Pontua a qualidade agronomica da sequencia solo/nutriçao.

    Principios:
    - Apos fixa_N  → alta exigencia: +3  (aproveita N fixado pela leguminosa)
    - Apos fixa_N  → media_alta:     +2
    - Apos fixa_N  → media:          +1
    - Apos esgota_NPK → fixa_N:      +2  (recupera solo com leguminosa)
    - Apos esgota_NPK → baixa exig:  +1  (deixa solo descansar)
    - Apos esgota_NPK → alta exig:   -2  (esgota mais ainda)
    - Dois esgota_NPK seguidos:       -1  (penalidade acumulada)
    - Apos melhora_estrutura:         +1
    - Apos repelente/reduz_nematoide: +1  (ambiente melhorado)
    - Candidato com repelente/nematoide: +1 (beneficio proativo para o proximo)
    """
    if not historico or culturas is None:
        return 0

    ultima = encontrar_cultura_por_nome(culturas, normalizar_nome_cultura(historico[-1]))
    if ultima is None:
        return 0

    efeito_ant  = parse_efeito_solo(ultima)
    exig_cand   = parse_exigencia(cultura)
    efeito_cand = parse_efeito_solo(cultura)

    pontos = 0

    if "fixa_n" in efeito_ant:
        bonus = {"alta": 3, "media_alta": 2, "media": 1}.get(exig_cand, 0)
        pontos += bonus

    if "esgota_npk" in efeito_ant:
        if "fixa_n" in efeito_cand:
            pontos += 2
        elif exig_cand in ("baixa", "baixa_media"):
            pontos += 1
        elif exig_cand in ("media_alta", "alta"):
            pontos -= 2
        if "esgota_npk" in efeito_cand:
            pontos -= 1

    if "melhora_estrutura" in efeito_ant:
        pontos += 1

    if "repelente" in efeito_ant or "reduz_nematoide" in efeito_ant:
        pontos += 1

    if "repelente" in efeito_cand or "reduz_nematoide" in efeito_cand:
        pontos += 1

    return pontos


def calcular_pontuacao(cultura, terraço, culturas, regras_gerais, regras_comp_global):
    historico = obter_historico(terraço)

    return (
        pontuar_biodinamica(cultura, historico, culturas)
        + pontuar_compatibilidade_global(cultura, regras_comp_global)
        + pontuar_compatibilidade_especifica(cultura, historico)
        + pontuar_diversidade(cultura, historico, culturas)
        + pontuar_leguminosa(cultura, historico, regras_gerais, culturas)
        + pontuar_solo_nutricional(cultura, historico, culturas)
    )

# ============================================================
# VERIFICAR COMPATIBILIDADE ENTRE CULTURAS
# ============================================================

def culturas_sao_compativeis(culturas, regras_incomp_global):
    nomes = [c["nome_da_cultura"].lower() for c in culturas]

    # 1. incompatibilidades globais (ex.: fenouil)
    for regra in regras_incomp_global:
        if regra["cultura"].lower() in nomes:
            return False

    # 2. incompatibilidades específicas (coluna "incompatibilidade")
    for c in culturas:
        if isinstance(c["incompatibilidade"], str):
            ruins = [x.strip().lower() for x in c["incompatibilidade"].split(",")]
            for r in ruins:
                if r in nomes:
                    return False

    # 3. evitar mesma família repetida
    familias = [c["família_botânica"] for c in culturas]
    if len(familias) != len(set(familias)):
        return False

    return True

# ============================================================
# SUGERIR CONSÓRCIO DE ATÉ 4 CULTURAS
# ============================================================

from itertools import combinations

def sugerir_consortio(terraco_id, culturas, terracos, regras_gerais, regras_incomp_global, regras_comp_global, n=4):
    terraço = next(t for t in terracos if int(t["id_terraço"]) == terraco_id)
    historico = obter_historico(terraço)
    historico_familias = historico_para_familias(culturas, historico)

    # 1. aplicar filtros
    culturas_filtradas = [
        c for c in culturas
        if aplicar_filtros(c, terraço, regras_gerais, regras_incomp_global, culturas)
    ]
    print(f"  [CONSORTIO] Terraço {terraco_id}: {len(culturas_filtradas)} culturas viáveis")

    if len(culturas_filtradas) < n:
        n = len(culturas_filtradas)

    # 2. Usar estratégia gulosa: selecionar os n melhores com máxima compatibilidade
    melhor = None
    melhor_pontuacao = -999
    
    # Computar pontuação para todas as culturas filtradas uma vez
    # Também calcular diversidade de família (quanto menos a família está no histórico, melhor)
    pontos_cache = {}
    diversidade_cache = {}
    
    for c in culturas_filtradas:
        pontos_cache[id(c)] = calcular_pontuacao(c, terraço, culturas, regras_gerais, regras_comp_global)
        fam = normalizar_nome_cultura(c.get("família_botânica", ""))
        diversidade_cache[id(c)] = 1 if fam not in historico_familias else 0

    # Selecionar gulosa: começar com a melhor cultura
    selecionadas = []
    candidatas = culturas_filtradas[:]
    
    for _ in range(n):
        if not candidatas:
            break
        
        melhor_c = None
        melhor_score = -999
        melhor_diversidade = -999
        
        for c in candidatas:
            test_list = selecionadas + [c]
            if culturas_sao_compativeis(test_list, regras_incomp_global):
                score = pontos_cache.get(id(c), 0)
                div = diversidade_cache.get(id(c), 0)
                # Usar pontuação principal, com diversidade de família como tie-breaker
                if score > melhor_score or (score == melhor_score and div > melhor_diversidade):
                    melhor_score = score
                    melhor_diversidade = div
                    melhor_c = c
        
        if melhor_c:
            selecionadas.append(melhor_c)
            candidatas.remove(melhor_c)
        else:
            break
    
    if selecionadas:
        melhor = selecionadas
        melhor_pontuacao = sum(pontos_cache.get(id(c), 0) for c in selecionadas)
    
    print(f"    ✓ Melhor consórcio: {melhor_pontuacao} pontos ({len(selecionadas)} culturas)")
    return melhor, melhor_pontuacao

# ============================================================
# APOIO PARA PLANO ANUAL
# ============================================================

from itertools import combinations

# -----------------------------
# PARSE DE DURAÇÃO
# -----------------------------
def parse_duracao(cultura):
    val = cultura.get("tempo_de_cultivo_dias")
    if isinstance(val, str) and "-" in val:
        a, b = val.split("-")
        return int(a), int(b)
    else:
        d = int(val)
        return d, d

def cultura_e_perene(cultura):
    # usa ocupa_solo_anos como critério primário (vem direto do Excel)
    import math
    anos = cultura.get("ocupa_solo_anos")
    if isinstance(anos, (int, float)) and not math.isnan(anos) and anos > 0:
        return True
    return False


# -----------------------------
# ESTAÇÃO DO ANO (EUROPA)
# -----------------------------
def estacao_do_dia(dia):
    """Retorna a estação do ano para o hemisfério norte (França), contando a partir de 1 jan.
    Solstícios/equinócios aproximados:
      inverno : 1 jan – 20 mar  (dia   0 –  78)  e  21 dez – 31 dez (dia 354 – 364)
      primavera: 21 mar – 20 jun (dia  79 – 170)
      verão    : 21 jun – 22 set (dia 171 – 264)
      outono   : 23 set – 20 dez (dia 265 – 353)
    """
    dia = dia % 365  # normaliza para multi-ano
    if 79 <= dia < 171:
        return "primavera"
    elif 171 <= dia < 265:
        return "verao"
    elif 265 <= dia < 354:
        return "outono"
    else:                   # 0-78 (jan-mar) e 354-364 (dez)
        return "inverno"


# -----------------------------
# FILTROS DE ÉPOCA E DURAÇÃO
# -----------------------------
def filtrar_por_epoca(cultura, epoca):
    ep = cultura.get("época_preferencial")
    if not isinstance(ep, str):
        return True
    # normaliza acentos e separa por - / , para suportar formatos variados do Excel
    ep = remover_acentos(ep.lower())
    partes = [p.strip() for p in ep.replace("/", "-").replace(",", "-").split("-") if p.strip()]
    return epoca in partes

def filtrar_por_duracao(cultura, dias_restantes):
    _, max_d = parse_duracao(cultura)
    return max_d <= dias_restantes


# -----------------------------
# FILTROS AGRONÔMICOS (SEM RESTRIÇÃO DE FAMÍLIA ENTRE CICLOS)
# -----------------------------
def aplicar_filtros_contexto(cultura, historico, regras_gerais, regras_incomp_global):
    # incompatibilidade global
    if not filtrar_incompatibilidade_global(cultura, regras_incomp_global):
        return False

    # incompatibilidade específica com histórico
    if not filtrar_incompatibilidade_especifica(cultura, historico):
        return False

    return True


# -----------------------------
# PONTUAÇÃO USANDO HISTÓRICO LOCAL
# -----------------------------
def calcular_pontuacao_contexto(cultura, historico, regras_gerais, regras_comp_global, culturas=None):
    return (
        pontuar_biodinamica(cultura, historico)
        + pontuar_compatibilidade_global(cultura, regras_comp_global)
        + pontuar_compatibilidade_especifica(cultura, historico)
        + pontuar_diversidade(cultura, historico)
        + pontuar_leguminosa(cultura, historico, regras_gerais)
        + pontuar_solo_nutricional(cultura, historico, culturas)
    )


# -----------------------------
# GERAR CONSÓRCIO PARA UM CICLO
# -----------------------------
def gerar_consortio_para_ciclo(culturas_validas, historico, regras_gerais, regras_incomp_global, regras_comp_global, dias_restantes):
    # tentar 4 → 3 → 2 → 1
    for tamanho in [4, 3, 2, 1]:
        melhor = None
        melhor_pontos = -999
        melhor_duracao = None

        for combo in combinations(culturas_validas, tamanho):
            combo = list(combo)
            combo_names = [c["nome_da_cultura"] for c in combo]
            print(f"      [TAM {tamanho}] Testando: {' + '.join(combo_names)}")

            # não repetir família dentro do consórcio
            familias = [c["família_botânica"] for c in combo]
            if len(familias) != len(set(familias)):
                print(f"        ✗ Família repetida")
                continue

            # compatibilidade entre culturas
            if not culturas_sao_compativeis(combo, regras_incomp_global):
                print(f"        ✗ Incompatível")
                continue

            # pontuação total
            total = sum(
                calcular_pontuacao_contexto(c, historico, regras_gerais, regras_comp_global)
                for c in combo
            )

            # duração do ciclo = menor max_dias do grupo
            max_dias = []
            for c in combo:
                _, mx = parse_duracao(c)
                max_dias.append(mx)
            duracao = min(max_dias)

            if duracao > dias_restantes:
                print(f"        ✗ Duração ({duracao} dias) > {dias_restantes} disponíveis")
                continue

            print(f"        ✓ Válido: {total} pontos, {duracao} dias")
            if total > melhor_pontos:
                melhor = combo
                melhor_pontos = total
                melhor_duracao = duracao
                print(f"          [NOVO MELHOR] {melhor_pontos} pontos!")

        if melhor is not None:
            return melhor, melhor_duracao, melhor_pontos

    return None, None, None


# -----------------------------
# GERAR CICLO COMPLETO
# -----------------------------
def gerar_ciclo_para_terraco(culturas, historico, regras_gerais, regras_incomp_global, regras_comp_global, dia_atual):
    epoca = estacao_do_dia(dia_atual)
    dias_restantes = 365 - dia_atual

    # filtrar culturas possíveis
    candidatas = []
    for c in culturas:
        if not aplicar_filtros_contexto(c, historico, regras_gerais, regras_incomp_global):
            continue
        if not filtrar_por_epoca(c, epoca):
            continue
        if not filtrar_por_duracao(c, dias_restantes):
            continue
        candidatas.append(c)

    if not candidatas:
        return None

    # gerar consórcio
    cons, duracao, pontos = gerar_consortio_para_ciclo(
        candidatas, historico, regras_gerais, regras_incomp_global, regras_comp_global, dias_restantes
    )

    if not cons:
        return None

    return {
        "inicio": dia_atual,
        "fim": dia_atual + duracao,
        "consorcio": cons,
        "pontos": pontos,
    }


# -----------------------------
# GERAR PLANO ANUAL COMPLETO
# -----------------------------
def gerar_plano_anual_para_terraco(terraco, culturas, regras_gerais, regras_incomp_global, regras_comp_global):
    historico = obter_historico(terraco)

    dia_atual = 0
    plano = []

    while dia_atual < 365:
        ciclo = gerar_ciclo_para_terraco(
            culturas, historico, regras_gerais, regras_incomp_global, regras_comp_global, dia_atual
        )

        if not ciclo:
            break

        plano.append(ciclo)
        dia_atual = ciclo["fim"]

        # atualizar histórico com famílias do ciclo (lower-case)
        for c in ciclo["consorcio"]:
            historico.append(str(c.get("família_botânica", "")).strip().lower())

    return plano

# ============================================================
# FUNÇÃO FINAL: SUGERIR CULTURA
# ============================================================

def sugerir_cultura(terraco_id, culturas, terracos, regras_gerais, regras_incomp_global, regras_comp_global):
    # pegar o terraço correto
    terraço = next(t for t in terracos if int(t["id_terraço"]) == terraco_id)

    historico = obter_historico(terraço)

    # aplicar filtros
    culturas_filtradas = [
        c for c in culturas
        if aplicar_filtros(c, terraço, regras_gerais, regras_incomp_global, culturas)
    ]

    # calcular pontuação
    ranking = sorted(
        culturas_filtradas,
        key=lambda c: calcular_pontuacao(c, terraço, culturas, regras_gerais, regras_comp_global),
        reverse=True
    )

    melhor = ranking[0]
    pontos = calcular_pontuacao(melhor, terraço, culturas, regras_gerais, regras_comp_global)

    # gerar explicação
    explicacao = []

    # biodinâmica
    if pontuar_biodinamica(melhor, historico, culturas) > 0:
        explicacao.append("✔ Alternância biodinâmica adequada")

    # leguminosa
    if pontuar_leguminosa(melhor, historico, regras_gerais, culturas) > 0:
        explicacao.append("✔ Necessidade de leguminosa atendida")

    # compatibilidade global
    if pontuar_compatibilidade_global(melhor, regras_comp_global) > 0:
        explicacao.append("✔ Bons companheiros no sistema")

    # diversidade
    if pontuar_diversidade(melhor, historico, culturas) > 0:
        explicacao.append("✔ Aumenta diversidade do terraço")

    return melhor["nome_da_cultura"], pontos, explicacao

# ============================================================
# PRÉ-CÁLCULO DE COMPATIBILIDADES (OTIMIZAÇÃO B)
# ============================================================

def construir_tabelas_compatibilidade(culturas, regras_comp_global, regras_incomp_global):
    nomes = [c["nome_da_cultura"] for c in culturas]

    compat = {n: set() for n in nomes}
    incompat = {n: set() for n in nomes}

    # --- detectar automaticamente as colunas ---
    def detectar_colunas(regra):
            chaves = list(regra.keys())
            # pega as duas primeiras colunas que têm string
            cols = [k for k in chaves if isinstance(regra[k], str)]
            if len(cols) >= 2:
                return cols[0], cols[1]
            return None, None

        # --- compatibilidade global ---
    if regras_comp_global:
            colA, colB = detectar_colunas(regras_comp_global[0])
            if colA and colB:
                for r in regras_comp_global:
                    a = str(r[colA]).strip()
                    b = str(r[colB]).strip()
                    if a in compat:
                        compat[a].add(b)
                    if b in compat:
                        compat[b].add(a)

        # --- incompatibilidade global ---
    if regras_incomp_global:
            colA, colB = detectar_colunas(regras_incomp_global[0])
            if colA and colB:
                for r in regras_incomp_global:
                    a = str(r[colA]).strip()
                    b = str(r[colB]).strip()
                    if a in incompat:
                        incompat[a].add(b)
                    if b in incompat:
                        incompat[b].add(a)

    return compat, incompat

# ============================================================
# MOTOR CONTÍNUO DE 4 SLOTS POR TERRAÇO
# ============================================================

ANOS_SEM_REPETIR_CULTURA   = 3  # regra: não repetir a mesma cultura por N anos
MAX_TERRACOS_MESMA_CULTURA = 6  # regra: no máximo N terraços com a mesma cultura ao mesmo tempo
MAX_PERENES_POR_TERRACO    = 2  # regra: no máximo N slots perenes por terraço ao mesmo tempo

def _inicializar_dias_ultimo_plantio(terraco, slots):
    """Constrói {nome_cultura: dia_ultimo_plantio} a partir do histórico e do consórcio inicial."""
    historico_orig = obter_historico(terraco)
    dias = {}
    n = len(historico_orig)
    for i, nome in enumerate(historico_orig):
        # aproximação: cada entrada = 90 dias antes da anterior, da mais antiga para a mais recente
        dias[normalizar_nome_cultura(nome)] = -(n - i) * 90
    for s in slots:
        if s["cultura"]:
            dias[normalizar_nome_cultura(s["cultura"]["nome_da_cultura"])] = 0
    return dias

from itertools import combinations

# -----------------------------
# ESCOLHER MELHOR CULTURA PARA UM SLOT
# -----------------------------
def escolher_melhor_cultura_para_slot(
    dia_plantio,
    culturas,
    slots,
    regras_gerais,
    regras_incomp_global,
    regras_comp_global,
    historico_familias,
    compat_global,
    incompat_global,
    total_dias=365,
    dias_ultimo_plantio=None,
    slots_globais=None
):
    epoca = estacao_do_dia(dia_plantio)
    familias_atuais = [
        s["cultura"]["família_botânica"]
        for s in slots
        if s["cultura"] is not None
    ]
    janela_repeticao = ANOS_SEM_REPETIR_CULTURA * 365

    melhor = None
    melhor_pontos = -999

    # contar quantos slots do terraço já estão com culturas perenes
    perenes_atuais = sum(1 for s in slots if s["cultura"] is not None and s["perene"])

    for c in culturas:
        # não repetir família dentro do consórcio atual
        if c["família_botânica"] in familias_atuais:
            continue

        # regra: limitar número de perenes por terraço (garante rotação de anuais)
        if cultura_e_perene(c) and perenes_atuais >= MAX_PERENES_POR_TERRACO:
            continue

        # época preferencial
        if not filtrar_por_epoca(c, epoca):
            continue

        # dias restantes até o fim do período de simulação
        dias_restantes = total_dias - dia_plantio
        if not filtrar_por_duracao(c, dias_restantes):
            continue

        # regra: não repetir a mesma cultura dentro da janela de N anos
        if dias_ultimo_plantio is not None:
            nome_norm = normalizar_nome_cultura(c["nome_da_cultura"])
            ultimo = dias_ultimo_plantio.get(nome_norm, -(janela_repeticao + 1))
            if dia_plantio - ultimo < janela_repeticao:
                continue

        # regra: não repetir a mesma família antes do período mínimo definido no Excel
        if not filtrar_por_familia(c, historico_familias, regras_gerais, culturas):
            continue

        # incompatibilidades específicas com o histórico do terraço
        if not filtrar_incompatibilidade_especifica(c, historico_familias):
            continue

        # evitar solanáceas consecutivas
        if regra_esta_ativa(regras_gerais, "evitar_solanaceas_seguidas"):
            if not filtrar_familia_seguida(c, historico_familias, culturas, "solanaceae"):
                continue

        # evitar brássicas consecutivas
        if regra_esta_ativa(regras_gerais, "evitar_brassicas_seguidas"):
            if not filtrar_familia_seguida(c, historico_familias, culturas, "brassicaceae"):
                continue

        # regra: no máximo MAX_TERRACOS_MESMA_CULTURA com a mesma planta ao mesmo tempo
        if slots_globais is not None:
            nome_norm = normalizar_nome_cultura(c["nome_da_cultura"])
            _, max_d = parse_duracao(c)
            dia_fim_est = dia_plantio + max_d
            concorrentes = sum(
                1 for (cn, di, df) in slots_globais
                if cn == nome_norm and di < dia_fim_est and df > dia_plantio
            )
            if concorrentes >= MAX_TERRACOS_MESMA_CULTURA:
                continue

        # compatibilidade com as outras culturas do terraço
        outras = [s["cultura"] for s in slots if s["cultura"] is not None]
        ok = True
        for outra in outras:
            nome_c = c["nome_da_cultura"]
            nome_o = outra["nome_da_cultura"]
            if nome_o in incompat_global[nome_c]:
                ok = False
                break

        if not ok:
            continue

        # pontuação (SEM compatibilidade específica — removida!)
        pontos = (
            pontuar_biodinamica(c, historico_familias, culturas)
            + pontuar_compatibilidade_global(c, regras_comp_global)
            + pontuar_diversidade(c, historico_familias, culturas)
            + pontuar_leguminosa(c, historico_familias, regras_gerais, culturas)
            + pontuar_solo_nutricional(c, historico_familias, culturas)
        )

        if pontos > melhor_pontos:
            melhor = c
            melhor_pontos = pontos

    return melhor, melhor_pontos



# -----------------------------
# INICIAR 4 SLOTS COM CONSÓRCIO INICIAL
# -----------------------------
def iniciar_slots_para_terraco(terraco_id, culturas, terracos, regras_gerais, regras_incomp_global, regras_comp_global, total_dias=365):
    terraco = next(t for t in terracos if int(t["id_terraço"]) == terraco_id)
    historico_culturas = obter_historico(terraco)

    consorcio_inicial, pontos = sugerir_consortio(
        terraco_id,
        culturas,
        terracos,
        regras_gerais,
        regras_incomp_global,
        regras_comp_global,
        n=4
    )

    slots = []
    for c in consorcio_inicial:
        _, max_d = parse_duracao(c)
        fim = min(total_dias, max_d)  # plantio no dia 0
        slots.append({
            "cultura": c,
            "fim": fim,
            "perene": cultura_e_perene(c),
        })
        historico_culturas.append(normalizar_nome_cultura(c.get("nome_da_cultura", "")))

    return slots, historico_culturas


# -----------------------------
# SIMULAÇÃO CONTÍNUA ATÉ O DIA 365
# -----------------------------
def simular_plano_continuo_para_terraco(
    terraco,
    culturas,
    terracos,
    regras_gerais,
    regras_incomp_global,
    regras_comp_global,
    dias_descanso,
    compat_global,
    incompat_global,
    total_dias=365
):

    terraco_id = int(terraco["id_terraço"])

    # iniciar 4 slots com consórcio inicial (mantemos comportamento existente)
    slots, historico_familias = iniciar_slots_para_terraco(
        terraco_id,
        culturas,
        terracos,
        regras_gerais,
        regras_incomp_global,
        regras_comp_global,
        total_dias=total_dias
    )
    dias_ultimo_plantio = _inicializar_dias_ultimo_plantio(terraco, slots)

    eventos = []
    dia_atual = 0
    ciclo_num = 0

    # registrar estado inicial em dia 0
    eventos.append({
        "dia": 0,
        "culturas": [s["cultura"]["nome_da_cultura"] if s["cultura"] else None for s in slots]
    })

    while True:
        # próximo fim de cultura entre os slots ativos
        fins_ativos = [s["fim"] for s in slots if s["cultura"] is not None]
        if not fins_ativos:
            print("  [CONTINUO] Nenhum slot ativo - finalizando")
            break

        proximo_fim = min(fins_ativos)
        if proximo_fim >= total_dias:
            print(f"  [CONTINUO] Próximo fim (dia {proximo_fim}) >= {total_dias} - finalizando")
            break

        dia_atual = proximo_fim
        ciclo_num += 1
        print(f"  [CONTINUO] Dia {dia_atual} (Ano {dia_atual // 365 + 1}) - Ciclo {ciclo_num}", flush=True)

        if ciclo_num > 10000:
            print("  [CICLO-DEBUG] Exceeded 10000 cycles — aborting to avoid infinite loop.", flush=True)
            break

        # registrar estado antes da substituição
        eventos.append({
            "dia": dia_atual,
            "culturas": [s["cultura"]["nome_da_cultura"] if s["cultura"] else None for s in slots]
        })

        slots_modificados = 0

        # processar slots que terminam nesse dia
        for idx, s in enumerate(slots):
            if s["cultura"] is None:
                continue

            if s["fim"] == dia_atual:
                cultura_atual = s["cultura"]["nome_da_cultura"]
                print(f"    [SLOT {idx+1}] {cultura_atual} terminou")
                # se for perene, renova o ciclo em vez de substituir
                if s["perene"]:
                    _, max_d = parse_duracao(s["cultura"])
                    novo_fim = min(total_dias, dia_atual + max_d)
                    slots[idx]["fim"] = novo_fim
                    slots_modificados += 1
                    print(f"      [PERENE] Renovando até dia {novo_fim}")
                    continue

                # dia de plantio da próxima cultura
                dia_plantio = dia_atual + dias_descanso
                print(f"      [DESCANSO] {dias_descanso} dias - Plantio em: dia {dia_plantio}")
                if dia_plantio >= total_dias:
                    slots[idx]["cultura"] = None
                    slots_modificados += 1
                    print(f"      [FIM DO PERÍODO] Sem tempo para nova cultura")
                    continue

                # escolher nova cultura para esse slot
                print(f"      [SELECIONANDO] Buscar cultura para dia {dia_plantio}...")
                nova, _ = escolher_melhor_cultura_para_slot(
                    dia_plantio, culturas, slots, regras_gerais, regras_incomp_global,
                    regras_comp_global, historico_familias, compat_global, incompat_global,
                    total_dias=total_dias, dias_ultimo_plantio=dias_ultimo_plantio
                )

                if not nova:
                    slots[idx]["cultura"] = None
                    slots_modificados += 1
                    print(f"      [SEM OPCAO] Nenhuma cultura disponível")
                    continue

                _, max_d = parse_duracao(nova)
                fim_novo = min(total_dias, dia_plantio + max_d)
                slots[idx] = {"cultura": nova, "fim": fim_novo, "perene": cultura_e_perene(nova)}
                nome_norm = normalizar_nome_cultura(nova["nome_da_cultura"])
                historico_familias.append(nome_norm)
                dias_ultimo_plantio[nome_norm] = dia_plantio
                slots_modificados += 1
                print(f"      ↻ Plantado {nova['nome_da_cultura']} (fim: dia {fim_novo})")

        # tentar preencher slots que ficaram vazios
        for idx, s in enumerate(slots):
            if s["cultura"] is not None or dia_atual >= total_dias:
                continue
            nova, _ = escolher_melhor_cultura_para_slot(
                dia_atual, culturas, slots, regras_gerais, regras_incomp_global,
                regras_comp_global, historico_familias, compat_global, incompat_global,
                total_dias=total_dias, dias_ultimo_plantio=dias_ultimo_plantio
            )
            if nova:
                _, max_d = parse_duracao(nova)
                fim_novo = min(total_dias, dia_atual + max_d)
                slots[idx] = {"cultura": nova, "fim": fim_novo, "perene": cultura_e_perene(nova)}
                nome_norm = normalizar_nome_cultura(nova["nome_da_cultura"])
                historico_familias.append(nome_norm)
                dias_ultimo_plantio[nome_norm] = dia_atual
                slots_modificados += 1
                print(f"      [SLOT {idx+1} VAZIO] Preenchido com {nova['nome_da_cultura']} (fim: dia {fim_novo})")

        if all(s["cultura"] is None for s in slots):
            print("  [CONTINUO] Todos os slots vazios - finalizando")
            break

        if slots_modificados == 0:
            print("  [CONTINUO] Nenhum slot foi modificado - finalizando para evitar loop infinito")
            break

    return eventos


# -----------------------------
# Versões alternativas/auxiliares para execução em lote e inicialização gulosa
# -----------------------------
def iniciar_slots_greedy_para_terraco(terraco, culturas, terracos, regras_gerais, regras_incomp_global, regras_comp_global, compat_global, incompat_global, n_slots=4):
    # inicia slots selecionando iterativamente a melhor cultura por slot (greedy)
    slots = []
    historico_familias = obter_historico(terraco)

    for _ in range(n_slots):
        dia_plantio = 0
        nova, _ = escolher_melhor_cultura_para_slot(
            dia_plantio,
            culturas,
            slots,
            regras_gerais,
            regras_incomp_global,
            regras_comp_global,
            historico_familias,
            compat_global,
            incompat_global
        )

        if not nova:
            slots.append({"cultura": None, "fim": None, "perene": False})
            continue

        _, max_d = parse_duracao(nova)
        fim = min(365, max_d)
        slots.append({"cultura": nova, "fim": fim, "perene": cultura_e_perene(nova)})
        historico_familias.append(nova.get("família_botânica", ""))

    return slots, historico_familias


def _tentar_preencher_slot(dia, slots, culturas, regras_gerais, regras_incomp_global,
                           regras_comp_global, historico_familias, compat_global, incompat_global,
                           total_dias, dias_ultimo_plantio, slots_globais):
    """Tenta escolher uma cultura para um slot vazio em até quatro tentativas progressivamente
    mais permissivas, para evitar longos períodos sem planta:
      1ª — todos os filtros + slots_globais (regras completas)
      2ª — histórico limitado (relaxa incompatibilidade específica) + slots_globais
      3ª — histórico limitado + sem restrição cross-terraço (slots_globais ignorado)
      4ª — sem regra de 3 anos + histórico mínimo + sem restrição cross-terraço
           (último recurso: garante que o slot nunca fica vazio por falta de opções)"""
    for hist, sg, dup in (
        (historico_familias,      slots_globais, dias_ultimo_plantio),  # completo
        (historico_familias[-8:], slots_globais, dias_ultimo_plantio),  # histórico curto
        (historico_familias[-8:], None,          dias_ultimo_plantio),  # sem cross-terraço
        (historico_familias[-4:], None,          None),                 # sem regra 3 anos
    ):
        nova, _ = escolher_melhor_cultura_para_slot(
            dia, culturas, slots, regras_gerais, regras_incomp_global,
            regras_comp_global, hist, compat_global, incompat_global,
            total_dias=total_dias, dias_ultimo_plantio=dup,
            slots_globais=sg
        )
        if nova:
            return nova
    return None


def simular_continuo_com_slots(slots, historico_familias, terraco, culturas, terracos, regras_gerais, regras_incomp_global, regras_comp_global, dias_descanso, compat_global, incompat_global, total_dias=365, slots_globais=None):
    eventos = []
    dia_atual = 0
    ciclo_num = 0
    dias_ultimo_plantio = _inicializar_dias_ultimo_plantio(terraco, slots)

    # estado inicial (dia 0)
    eventos.append({
        "dia": 0,
        "culturas": [s["cultura"]["nome_da_cultura"] if s["cultura"] else None for s in slots]
    })

    while True:
        fins_ativos = [s["fim"] for s in slots if s["cultura"] is not None]
        if not fins_ativos:
            break

        proximo_fim_natural = min(fins_ativos)
        if proximo_fim_natural >= total_dias:
            break

        # se há slots vazios e o próximo evento natural é distante, adiantar a iteração
        # para tentar preencher os slots mais cedo (máximo a cada 30 dias)
        tem_vazios = any(s["cultura"] is None for s in slots)
        if tem_vazios and proximo_fim_natural - dia_atual > 30:
            proximo_fim = dia_atual + 30
        else:
            proximo_fim = proximo_fim_natural

        if proximo_fim >= total_dias:
            break

        dia_atual = proximo_fim
        ciclo_num += 1

        slots_modificados = 0

        # processar slots que chegaram ao fim
        for idx, s in enumerate(slots):
            if s["cultura"] is None or s["fim"] != dia_atual:
                continue

            if s["perene"]:
                _, max_d = parse_duracao(s["cultura"])
                slots[idx]["fim"] = min(total_dias, dia_atual + max_d)
                slots_modificados += 1
                continue

            dia_plantio = dia_atual + dias_descanso
            if dia_plantio >= total_dias:
                slots[idx]["cultura"] = None
                slots_modificados += 1
                continue

            nova = _tentar_preencher_slot(
                dia_plantio, slots, culturas, regras_gerais, regras_incomp_global,
                regras_comp_global, historico_familias, compat_global, incompat_global,
                total_dias, dias_ultimo_plantio, slots_globais
            )

            if not nova:
                slots[idx]["cultura"] = None
                slots_modificados += 1
                continue

            _, max_d = parse_duracao(nova)
            fim_novo = min(total_dias, dia_plantio + max_d)
            slots[idx] = {"cultura": nova, "fim": fim_novo, "perene": cultura_e_perene(nova)}
            nome_norm = normalizar_nome_cultura(nova["nome_da_cultura"])
            historico_familias.append(nome_norm)
            dias_ultimo_plantio[nome_norm] = dia_plantio
            if slots_globais is not None:
                slots_globais.append((nome_norm, dia_plantio, fim_novo))
            slots_modificados += 1

        # tentar preencher slots que ficaram vazios (inclusive de iterações anteriores)
        for idx, s in enumerate(slots):
            if s["cultura"] is not None or dia_atual >= total_dias:
                continue
            nova = _tentar_preencher_slot(
                dia_atual, slots, culturas, regras_gerais, regras_incomp_global,
                regras_comp_global, historico_familias, compat_global, incompat_global,
                total_dias, dias_ultimo_plantio, slots_globais
            )
            if nova:
                _, max_d = parse_duracao(nova)
                fim_novo = min(total_dias, dia_atual + max_d)
                slots[idx] = {"cultura": nova, "fim": fim_novo, "perene": cultura_e_perene(nova)}
                nome_norm = normalizar_nome_cultura(nova["nome_da_cultura"])
                historico_familias.append(nome_norm)
                dias_ultimo_plantio[nome_norm] = dia_atual
                if slots_globais is not None:
                    slots_globais.append((nome_norm, dia_atual, fim_novo))
                slots_modificados += 1

        # gravar evento APÓS processamento (barras do Gantt ficam com duração correta)
        eventos.append({
            "dia": dia_atual,
            "culturas": [s["cultura"]["nome_da_cultura"] if s["cultura"] else None for s in slots]
        })

        if all(s["cultura"] is None for s in slots):
            break

    return eventos


def simular_plano_continuo_para_todos_terracos(terracos, culturas, regras_gerais, regras_incomp_global, regras_comp_global, dias_descanso, compat_global, incompat_global, total_dias=365):
    resultados = {}
    slots_globais = []  # lista compartilhada entre terraços: (nome_norm, dia_inicio, dia_fim)

    for t in terracos:
        terraco_id = int(t["id_terraço"])
        slots, historico_familias = iniciar_slots_para_terraco(
            terraco_id,
            culturas,
            terracos,
            regras_gerais,
            regras_incomp_global,
            regras_comp_global,
            total_dias=total_dias
        )

        # registrar culturas iniciais no índice global
        for s in slots:
            if s["cultura"] is not None:
                nome_norm = normalizar_nome_cultura(s["cultura"]["nome_da_cultura"])
                slots_globais.append((nome_norm, 0, s["fim"]))

        eventos = simular_continuo_com_slots(
            slots, historico_familias, t, culturas, terracos,
            regras_gerais, regras_incomp_global, regras_comp_global,
            dias_descanso, compat_global, incompat_global,
            total_dias=total_dias, slots_globais=slots_globais
        )

        resultados[terraco_id] = eventos

    return resultados



# ============================================================
# EXPORTAR RESULTADOS DO MOTOR CONTÍNUO
# ============================================================

def exportar_para_excel(planos_continuos, nome_arquivo="plano_rotacao_continuo.xlsx"):
    def montar_linha(terraco_id, ev):
        dia_abs = ev["dia"]
        row = {
            "Terrasse": terraco_id,
            "Année": dia_abs // 365 + 1,
            "Jour de l'Année": dia_abs % 365,
            "Jour Absolu": dia_abs,
        }
        for i, nome in enumerate(ev["culturas"], 1):
            row[f"Emplacement {i}"] = nome if nome else "(vide)"
        return row

    linhas_resumo = [
        montar_linha(tid, ev)
        for tid, eventos in planos_continuos.items()
        for ev in eventos
    ]

    with pd.ExcelWriter(nome_arquivo, engine="openpyxl") as writer:
        pd.DataFrame(linhas_resumo).to_excel(writer, sheet_name="Toutes les Terrasses", index=False)

        for terraco_id, eventos in planos_continuos.items():
            linhas = [montar_linha(terraco_id, ev) for ev in eventos]
            df = pd.DataFrame(linhas).drop(columns=["Terrasse"])
            sheet_name = f"Terrasse {terraco_id}"
            df.to_excel(writer, sheet_name=sheet_name, index=False)

    print(f"\n✓ Résultat exporté vers : {nome_arquivo}")


# ============================================================
# GANTT CHART AGRÍCOLA
# ============================================================

from datetime import date, timedelta

GANTT_BASE_DATE = date(2025, 1, 1)


def _extrair_barras_gantt(eventos, total_dias):
    """Converte lista de eventos (estados) em barras contínuas (inicio, fim, cultura)."""
    if not eventos:
        return []

    barras = []
    abertos = {}  # slot_idx -> (dia_inicio, nome_cultura)

    for slot_idx, nome in enumerate(eventos[0]["culturas"]):
        if nome is not None:
            abertos[slot_idx] = (eventos[0]["dia"], nome)

    for ev in eventos[1:]:
        for slot_idx, nome in enumerate(ev["culturas"]):
            prev = abertos.get(slot_idx)
            if prev is not None and nome != prev[1]:
                barras.append({"slot": slot_idx + 1, "cultura": prev[1],
                                "inicio": prev[0], "fim": ev["dia"]})
                if nome is not None:
                    abertos[slot_idx] = (ev["dia"], nome)
                else:
                    del abertos[slot_idx]
            elif prev is None and nome is not None:
                abertos[slot_idx] = (ev["dia"], nome)

    for slot_idx, (inicio, nome) in abertos.items():
        barras.append({"slot": slot_idx + 1, "cultura": nome,
                        "inicio": inicio, "fim": total_dias})
    return barras


def gerar_gantt(planos_continuos, total_dias, nome_arquivo="gantt_rotacao.html"):
    try:
        import plotly.express as px
    except ImportError:
        print("  [GANTT] plotly não instalado. Execute: pip install plotly")
        return

    # Montar tabela de barras
    registros = []
    for terraco_id, eventos in sorted(planos_continuos.items()):
        for b in _extrair_barras_gantt(eventos, total_dias):
            inicio_dt = GANTT_BASE_DATE + timedelta(days=int(b["inicio"]))
            fim_dt    = GANTT_BASE_DATE + timedelta(days=int(b["fim"]))
            ano_inicio = b["inicio"] // 365 + 1
            ano_fim    = b["fim"]    // 365 + 1
            registros.append({
                "Référence":  f"T{terraco_id:02d} · Emplacement {b['slot']}",
                "Terrasse":   f"Terrasse {terraco_id}",
                "Emplacement": f"Emplacement {b['slot']}",
                "Culture":    b["cultura"],
                "Début":      inicio_dt,
                "Fin":        fim_dt,
                "Période":    f"Année {ano_inicio}–{ano_fim}" if ano_inicio != ano_fim else f"Année {ano_inicio}",
                "Durée":      f"{b['fim'] - b['inicio']} jours",
            })

    if not registros:
        print("  [GANTT] Aucune donnée à exporter.")
        return

    df = pd.DataFrame(registros)

    # Paleta de cores consistente por cultura
    culturas_unicas = sorted(df["Culture"].unique())
    paleta = (px.colors.qualitative.Dark24 + px.colors.qualitative.Light24)
    cor_por_cultura = {c: paleta[i % len(paleta)] for i, c in enumerate(culturas_unicas)}

    fig = px.timeline(
        df,
        x_start="Début",
        x_end="Fin",
        y="Référence",
        color="Culture",
        color_discrete_map=cor_por_cultura,
        hover_data={
            "Terrasse": True, "Emplacement": True, "Culture": True,
            "Période": True, "Durée": True,
            "Début": True, "Fin": True, "Référence": False,
        },
        title="Plan de Rotation des Cultures — Moteur Continu",
    )

    # Linhas verticais marcando transição de anos
    n_anos = total_dias // 365
    for ano in range(1, n_anos):
        data_limite = str(GANTT_BASE_DATE + timedelta(days=ano * 365))
        fig.add_vline(
            x=data_limite,
            line_dash="dash", line_color="rgba(80,80,80,0.5)", line_width=1,
        )
        fig.add_annotation(
            x=data_limite, y=1.01, yref="paper",
            text=f"Année {ano + 1}", showarrow=False,
            font=dict(size=11, color="gray"),
            xanchor="left",
        )

    n_linhas = df["Référence"].nunique()
    fig.update_yaxes(autorange="reversed", categoryorder="category ascending", title="", tickfont=dict(size=10))
    fig.update_xaxes(title="Date", tickformat="%b %Y")
    fig.update_layout(
        height=max(700, n_linhas * 24 + 180),
        title_font_size=17,
        legend_title_text="Culture",
        legend=dict(orientation="v", x=1.01, y=1, font=dict(size=10)),
        margin=dict(l=140, r=200, t=80, b=60),
        plot_bgcolor="white",
        paper_bgcolor="white",
    )

    fig.write_html(nome_arquivo, include_plotlyjs="cdn")
    print(f"\n✓ Gantt exporté vers : {nome_arquivo}")


# ============================================================
# GANTT IMPRIMÍVEL (PDF, uma página por terraço)
# ============================================================

def gerar_gantt_imprimivel(planos_continuos, total_dias, nome_arquivo="gantt_rotacao_imprimivel.pdf"):
    try:
        import matplotlib
        matplotlib.use("Agg")
        import matplotlib.pyplot as plt
        import matplotlib.patches as mpatches
        from matplotlib.backends.backend_pdf import PdfPages
    except ImportError:
        print("  [GANTT] matplotlib não instalado. Execute: pip install matplotlib")
        return

    # ── paleta de cores consistente entre páginas ──────────────────────────
    todos_culturas_set = set()
    barras_por_terraco = {}
    for terraco_id, eventos in sorted(planos_continuos.items()):
        barras = _extrair_barras_gantt(eventos, total_dias)
        barras_por_terraco[terraco_id] = barras
        for b in barras:
            todos_culturas_set.add(b["cultura"])

    culturas_ordenadas = sorted(todos_culturas_set)
    paleta_mpl = (
        list(plt.cm.tab20.colors)
        + list(plt.cm.tab20b.colors)
        + list(plt.cm.tab20c.colors)
    )
    cor_por_cultura = {
        c: paleta_mpl[i % len(paleta_mpl)]
        for i, c in enumerate(culturas_ordenadas)
    }

    def _luminancia(rgb):
        r, g, b = rgb[:3]
        return 0.299 * r + 0.587 * g + 0.114 * b

    n_anos = total_dias // 365

    # ── PDF ────────────────────────────────────────────────────────────────
    with PdfPages(nome_arquivo) as pdf:
        for terraco_id, barras in barras_por_terraco.items():
            fig, ax = plt.subplots(figsize=(16.5, 5.8))
            fig.patch.set_facecolor("white")
            ax.set_facecolor("#f7f7f7")

            culturas_presentes = set()

            for b in barras:
                y = b["slot"] - 1          # slot 1 → y=0, slot 4 → y=3
                cor = cor_por_cultura.get(b["cultura"], (0.7, 0.7, 0.7))
                dur = b["fim"] - b["inicio"]

                ax.barh(y, dur, left=b["inicio"],
                        height=0.65, color=cor,
                        edgecolor="white", linewidth=0.7, zorder=3)

                culturas_presentes.add(b["cultura"])

                # nome dentro da barra (só se houver espaço mínimo)
                if dur >= 20:
                    cor_txt = "white" if _luminancia(cor) < 0.45 else "#1a1a1a"
                    ax.text(
                        b["inicio"] + dur / 2, y,
                        b["cultura"],
                        ha="center", va="center",
                        fontsize=6.5, fontweight="bold",
                        color=cor_txt, clip_on=True, zorder=4,
                    )

            # ── faixas de fundo e separadores de ano ──────────────────────
            for ano in range(1, n_anos + 1):
                x_ano = ano * 365
                if ano % 2 == 0:
                    ax.axvspan((ano - 1) * 365, x_ano,
                               color="#ececec", zorder=1)
                if ano < n_anos:
                    ax.axvline(x=x_ano, color="#888888",
                               linestyle="--", linewidth=1.0, zorder=5)
                ax.text((ano - 1) * 365 + 365 / 2, 3.58,
                        f"Année {ano}",
                        ha="center", va="bottom",
                        fontsize=9, fontweight="bold", color="#444444")

            # ── linhas e ticks mensais ─────────────────────────────────────
            MESES_PT = ["Jan","Fév","Mar","Avr","Mai","Juin",
                        "Juil","Août","Sep","Oct","Nov","Déc"]
            tick_dias, tick_labels = [], []
            d_mes = GANTT_BASE_DATE
            while (d_mes - GANTT_BASE_DATE).days < total_dias:
                dia_off = (d_mes - GANTT_BASE_DATE).days
                tick_dias.append(dia_off)
                tick_labels.append(MESES_PT[d_mes.month - 1])
                # avançar para o 1º dia do mês seguinte
                d_mes = (d_mes.replace(day=1) + timedelta(days=32)).replace(day=1)

            for dia in tick_dias:
                if dia > 0 and dia % 365 != 0:   # não sobrepor linhas de ano
                    ax.axvline(x=dia, color="#cccccc",
                               linewidth=0.4, zorder=2)

            # ── eixos ──────────────────────────────────────────────────────
            ax.set_xlim(0, total_dias)
            ax.set_ylim(-0.5, 4.1)
            ax.set_yticks([0, 1, 2, 3])
            ax.set_yticklabels(
                ["Emplacement 1", "Emplacement 2", "Emplacement 3", "Emplacement 4"],
                fontsize=10, fontweight="bold"
            )
            ax.set_xticks(tick_dias)
            ax.set_xticklabels(tick_labels, fontsize=6, ha="center")
            ax.tick_params(axis="x", length=3, color="lightgray")

            ax.yaxis.grid(True, color="white", linewidth=1.5, zorder=1)
            for spine in ax.spines.values():
                spine.set_edgecolor("lightgray")

            # ── título ─────────────────────────────────────────────────────
            ax.set_title(
                f"Terrasse {terraco_id}  —  Rotation des Cultures  ({n_anos} Ans)",
                fontsize=13, fontweight="bold", pad=12, loc="left"
            )

            # ── legenda abaixo do gráfico (não tapa as barras) ────────────
            patches = [
                mpatches.Patch(color=cor_por_cultura[c], label=c)
                for c in sorted(culturas_presentes)
            ]
            ax.legend(
                handles=patches,
                loc="upper center",
                bbox_to_anchor=(0.5, -0.14),
                fontsize=7, title="Cultures", title_fontsize=8,
                ncol=min(8, len(patches)),
                framealpha=0.95, edgecolor="lightgray",
                borderaxespad=0,
            )

            plt.tight_layout(pad=1.0)
            pdf.savefig(fig, bbox_inches="tight", dpi=150)
            plt.close(fig)

        # metadados do PDF
        d = pdf.infodict()
        d["Title"] = "Plan de Rotation des Cultures — Ferme Mauro Colagreco"
        d["Author"] = "Moteur Continu de Rotation"

    print(f"\n✓ Gantt imprimable exporté vers : {nome_arquivo}")


# ============================
# BLOCO PRINCIPAL
# ============================
if __name__ == "__main__":

    sheets = carregar_dados()

    df_culturas = sheets["Lista de culturas"]
    print("\n=== TABELA: CULTURAS ===")
    print(df_culturas.head())

    df_terracos = sheets["Terraços"]
    print("\n=== TABELA: TERRAÇOS ===")
    print(df_terracos.head())

    df_regras = sheets["Regras de rotações"]
    print("\n=== ABA DE REGRAS (BRUTA) ===")
    print(df_regras.head(20))

    tabelas_regras = separar_tabelas_regras(df_regras)

    print("\n=== TABELAS SEPARADAS ===")
    for nome, tabela in tabelas_regras.items():
        print(f"\n--- {nome} ---")
        print(tabela)

    # Criar estruturas de dados
    culturas = df_culturas.to_dict(orient="records")
    terracos = df_terracos.to_dict(orient="records")

    # Reconstruir Regras_Gerais com colunas claras
    df_rg = tabelas_regras["Regras_Gerais"]

    regras_gerais = []
    for _, row in df_rg.iterrows():
        regras_gerais.append({
            "regra": row.iloc[0],
            "valor": row.iloc[1],
            "unidade": row.iloc[2],
            "descricao": row.iloc[3],
        })

    regras_familias = tabelas_regras["Regras_Familias"].to_dict(orient="records")
    regras_incomp_global = tabelas_regras["Incompatibilidade_Global"].to_dict(orient="records")
    regras_comp_global = tabelas_regras["Compatibilidade_Global"].to_dict(orient="records")

    # ============================================================
    # PRÉ-CÁLCULO DE COMPATIBILIDADES (TEM QUE VIR AQUI!)
    # ============================================================
    compat_global, incompat_global = construir_tabelas_compatibilidade(
        culturas,
        regras_comp_global,
        regras_incomp_global
    )
    print("\n=== ESTRUTURAS DE DADOS CRIADAS ===")
    print(f"Total de culturas: {len(culturas)}")
    print(f"Total de terraços: {len(terracos)}")
    print(f"Regras gerais: {len(regras_gerais)} itens")
    print(f"Regras por família: {len(regras_familias)} itens")
    print(f"Incompatibilidades globais: {len(regras_incomp_global)} itens")
    print(f"Compatibilidades globais: {len(regras_comp_global)} itens")

    

    # ============================================================
    # TESTE DOS FILTROS — TERRAÇO 1
    # # ============================================================
    terraço_teste = terracos[0]

    culturas_filtradas = [
        c for c in culturas
        if aplicar_filtros(c, terraço_teste, regras_gerais, regras_incomp_global, culturas)
    ]

    print("\n=== CULTURAS APÓS FILTROS (TERRAÇO 1) ===")
    print(len(culturas_filtradas), "culturas possíveis")

    # ============================================================
    # SUGESTÃO FINAL PARA O TERRAÇO 1
    # ============================================================
    print("\n=== SUGESTÃO FINAL PARA O TERRAÇO 1 ===")

    nome, pontos, explicacao = sugerir_cultura(
        1,
        culturas,
        terracos,
        regras_gerais,
        regras_incomp_global,
        regras_comp_global
    )

    print(f"Cultura sugerida: {nome} ({pontos} pontos)")
    print("Motivos:")
    for e in explicacao:
        print(" -", e)

    # ============================================================
    # CONSÓRCIO RECOMENDADO PARA O TERRAÇO 1
    # ============================================================
    print("\n=== CONSÓRCIO RECOMENDADO PARA O TERRAÇO 1 ===")
    consorcio, pontos = sugerir_consortio(
        1,
        culturas,
        terracos,
        regras_gerais,
        regras_incomp_global,
        regras_comp_global,
        n=4
    )

    if consorcio:
        print(f"Pontuação total: {pontos}")
        print("Culturas recomendadas:")
        for c in consorcio:
            print(" -", c["nome_da_cultura"])
    else:
        print("Nenhum consórcio viável encontrado.")

    # ============================================================
    # CONSÓRCIOS PARA TODOS OS TERRAÇOS
    # ============================================================
    print("\n=== CONSÓRCIOS PARA TODOS OS TERRAÇOS ===")

    consorcios_fazenda = []

    for idx, t in enumerate(terracos, 1):
        terraco_id = int(t["id_terraço"])
        print(f"[PROCESSANDO] Terraço {terraco_id} ({idx}/{len(terracos)})")

        consorcio, pontos = sugerir_consortio(
            terraco_id,
            culturas,
            terracos,
            regras_gerais,
            regras_incomp_global,
            regras_comp_global,
            n=4
        )

        consorcios_fazenda.append({
            "terraço": terraco_id,
            "pontos": pontos,
            "culturas": [c["nome_da_cultura"] for c in consorcio] if consorcio else []
        })

    for item in consorcios_fazenda:
        print(f"\nTerraço {item['terraço']} — {item['pontos']} pontos")
        for c in item["culturas"]:
            print(" -", c)

    # ============================================================
    # PLANO ANUAL PARA O TERRAÇO 1
    # ============================================================
    #print("\n=== PLANO ANUAL PARA O TERRAÇO 1 ===")
    #terraco1 = next(t for t in terracos if int(t["id_terraço"]) == 1)

    #plano_anual = gerar_plano_anual_para_terraco(
        #terraco1,
        #culturas,
        #regras_gerais,
        #regras_incomp_global,
        #regras_comp_global,
    #)

    #for i, ciclo in enumerate(plano_anual, start=1):
        #print(f"\nCiclo {i}: dias {ciclo['inicio']}–{ciclo['fim']} (pontos: {ciclo['pontos']})")
        #print("Culturas:")
        #for c in ciclo["consorcio"]:
            #print(" -", c["nome_da_cultura"])
    # ============================================================
    # PLANO CONTÍNUO PARA O TERRAÇO 1
    # ============================================================
    print("\n=== PLANO CONTÍNUO PARA O TERRAÇO 1 (4 slots, +7 dias de descanso) ===")
    print(f"\n[INICIANDO] Simulação contínua... start_time={time.strftime('%Y-%m-%d %H:%M:%S')}", flush=True)
    terraco1 = next(t for t in terracos if int(t["id_terraço"]) == 1)

    eventos = simular_plano_continuo_para_terraco(
        terraco1,
        culturas,
        terracos,
        regras_gerais,
        regras_incomp_global,
        regras_comp_global,
        dias_descanso=7,
        compat_global=compat_global,
        incompat_global=incompat_global
    )

    for ev in eventos:
        print(f"\nDia {ev['dia']}:")
        for nome in ev["culturas"]:
            print(" -", nome if nome else "[slot vazio]")

    # ============================================================
    # PLANO CONTÍNUO PARA TODOS OS TERRAÇOS
    # ============================================================
    print("\n=== PLANO CONTÍNUO PARA TODOS OS TERRAÇOS (4 slots, +7 dias de descanso) ===")
    print(f"\n[INICIANDO] Simulação contínua para todos os terraços... start_time={time.strftime('%Y-%m-%d %H:%M:%S')}", flush=True)

    planos_continuos = simular_plano_continuo_para_todos_terracos(
        terracos,
        culturas,
        regras_gerais,
        regras_incomp_global,
        regras_comp_global,
        dias_descanso=7,
        compat_global=compat_global,
        incompat_global=incompat_global
    )

    for terraco_id, eventos in planos_continuos.items():
        print(f"\n--- Terraço {terraco_id} ({len(eventos)} eventos) ---")
        for ev in eventos:
            print(f"Dia {ev['dia']}:")
            for nome in ev['culturas']:
                print(" -", nome if nome else "[slot vazio]")

    # ============================================================
    # EXPORTAR PLANO 1 ANO PARA EXCEL
    # ============================================================
    exportar_para_excel(planos_continuos, nome_arquivo="plano_rotacao_1_ano.xlsx")

    # ============================================================
    # PLANO CONTÍNUO 4 ANOS PARA TODOS OS TERRAÇOS
    # ============================================================
    print("\n=== PLANO CONTÍNUO 4 ANOS PARA TODOS OS TERRAÇOS (4 slots, +7 dias de descanso) ===")
    print(f"\n[INICIANDO] Simulação 4 anos... start_time={time.strftime('%Y-%m-%d %H:%M:%S')}", flush=True)

    planos_4_anos = simular_plano_continuo_para_todos_terracos(
        terracos,
        culturas,
        regras_gerais,
        regras_incomp_global,
        regras_comp_global,
        dias_descanso=7,
        compat_global=compat_global,
        incompat_global=incompat_global,
        total_dias=4 * 365
    )

    exportar_para_excel(planos_4_anos, nome_arquivo="plano_rotacao_4_anos.xlsx")
    gerar_gantt(planos_4_anos, total_dias=4 * 365, nome_arquivo="gantt_rotacao_4_anos.html")
    gerar_gantt_imprimivel(planos_4_anos, total_dias=4 * 365, nome_arquivo="gantt_rotacao_4_anos.pdf")
    print(f"\n[FIM] Simulação 4 anos concluída. start_time={time.strftime('%Y-%m-%d %H:%M:%S')}", flush=True)


            








    





