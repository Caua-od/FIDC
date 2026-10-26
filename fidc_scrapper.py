#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
fnet_scraper_fidc.py
=====================

Baixa do Fundos.NET (B3/CVM) documentos dos ultimos N meses de todas as
classes/fundos de FIDC marcados como ativos no arquivo de registro de
classes da CVM.

TIPOS DE DOCUMENTO SUPORTADOS
------------------------------
- Informe Mensal Estruturado  (nivel classe, periodico mensal)
- Regulamento                 (nivel classe, avulso)
- Fato Relevante              (nivel classe, avulso)
- Atos de Deliberacao do Adm. (nivel classe, avulso)
- Relatorio de Ag. de Rating  (nivel fundo, avulso)

Cada tipo tem sua propria pasta de saida e janela de busca configuravel
na secao CONFIG abaixo.

COMO FUNCIONA
-------------
1. Le o arquivo "classe" (csv, separador ';', export CVM) dentro de PASTA_DADOS.
2. Filtra FIDCs ativos / pre-operacionais.
3. FASE 1 - para cada CNPJ de classe:
     ancora a sessao, lista todos os documentos e filtra/baixa os tipos
     de nivel "classe".
4. FASE 2 - para cada CNPJ de fundo (unico):
     ancora a sessao, lista todos os documentos e filtra/baixa os tipos
     de nivel "fundo".
5. Documentos periodicos (mensal): filtra por janela de meses, deduplica
   por competencia (MM/YYYY), mantendo a versao mais recente (maior id).
6. Documentos avulsos: pega os N mais recentes (por id), deduplicando
   por data de referencia.

NOME DOS ARQUIVOS
------------------
    <cnpj14>_<Prefixo>_<data-referencia>.<ext>
    ex: 37406818000119_InformeMensal_09-2026.xml
        37406818000119_Regulamento_01-10-2026.xml

ROBUSTEZ
--------
- Retry com backoff exponencial, circuit breaker, erros.log.
- Re-rodar e seguro: arquivos existentes sao pulados.

REQUISITOS
----------
    pip install requests
"""

from __future__ import annotations

import csv
import os
import random
import re
import sys
import time
import unicodedata
from dataclasses import dataclass
from datetime import date
from typing import Optional

import requests

# ============================================================================
# CONFIG
# ============================================================================

PASTA_DADOS = r"C:\Users\olive\Downloads\Franklin project\Data\Scrapping"

# Pasta onde ficam os registros de cadastro de fundo/classe da CVM (arquivo
# .csv/.txt com "classe" no nome). Pode ser diferente de PASTA_DADOS.
PASTA_REGISTRO_CLASSE = r"C:\Users\olive\Downloads\Franklin project\Data\registro_fundo_classe"

# ---------- Configuracao por tipo de documento ------------------------------
# Ajuste max_recentes / meses_retroativos de cada tipo diretamente aqui.
#
#   periodicidade = "mensal"  -> filtra por janela de meses (meses_retroativos),
#                                deduplica por competencia MM/YYYY, pega todos
#                                dentro da janela.
#   periodicidade = "avulso"  -> pega os N documentos mais recentes
#                                (max_recentes), sem filtro de data.
#
#   campo_filtro = "tipo"      -> compara nome_filtro com tipoDocumento
#   campo_filtro = "categoria" -> compara nome_filtro com categoriaDocumento
#
#   nivel = "classe" -> ancora sessao pelo CNPJ da classe
#   nivel = "fundo"  -> ancora sessao pelo CNPJ do fundo


@dataclass
class TipoDocConfig:
    nome_filtro: str       # substring normalizada p/ filtrar
    campo_filtro: str      # "tipo" ou "categoria"
    pasta: str             # subpasta dentro de PASTA_DADOS
    prefixo_arquivo: str   # prefixo no nome do arquivo salvo
    max_recentes: int = 2  # (avulso) quantos docs mais recentes manter
    meses_retroativos: int = 2  # (mensal) janela de meses pra tras
    nivel: str = "classe"
    periodicidade: str = "avulso"


TIPOS_DOCUMENTO: list[TipoDocConfig] = [
    TipoDocConfig(
        nome_filtro="informe mensal estruturado",
        campo_filtro="tipo",
        pasta="xml",
        prefixo_arquivo="InformeMensal",
        meses_retroativos=2,        # <-- ajuste aqui
        nivel="classe",
        periodicidade="mensal",
    ),
    TipoDocConfig(
        nome_filtro="regulamento",
        campo_filtro="categoria",
        pasta="regulamentos",
        prefixo_arquivo="Regulamento",
        max_recentes=2,             # <-- ajuste aqui
        nivel="classe",
        periodicidade="avulso",
    ),
    TipoDocConfig(
        nome_filtro="fato relevante",
        campo_filtro="categoria",
        pasta="fatos_relevantes",
        prefixo_arquivo="FatoRelevante",
        max_recentes=2,             # <-- ajuste aqui
        nivel="classe",
        periodicidade="avulso",
    ),
    TipoDocConfig(
        nome_filtro="atos de deliberacao do administrador",
        campo_filtro="categoria",
        pasta="atos_deliberacao",
        prefixo_arquivo="AtosDeliberacao",
        max_recentes=2,             # <-- ajuste aqui
        nivel="classe",
        periodicidade="avulso",
    ),
    TipoDocConfig(
        nome_filtro="relatorio de agencia de rating",
        campo_filtro="tipo",
        pasta="rating",
        prefixo_arquivo="Rating",
        max_recentes=2,             # <-- ajuste aqui
        nivel="fundo",
        periodicidade="avulso",
    ),
]

# ---------- Protecao de escopo (so mexe nas pastas deste projeto) -----------
# Nenhuma escrita/criacao de pasta e feita fora disso. O script NUNCA apaga
# ou sobrescreve arquivo que ele mesmo nao tenha criado nesta execucao.
NOME_LOG_ERROS = "erros.log"
_SUBPASTAS_PERMITIDAS = {tc.pasta for tc in TIPOS_DOCUMENTO}


def _validar_escrita(caminho: str) -> None:
    """Garante que 'caminho' esta dentro de uma subpasta conhecida de
    PASTA_DADOS (ou e o log de erros na raiz dela). Aborta a execucao
    (RuntimeError) se nao estiver -- nunca escreve fora do escopo."""
    raiz = os.path.realpath(PASTA_DADOS)
    alvo = os.path.realpath(caminho)

    if os.path.commonpath([raiz, alvo]) != raiz:
        raise RuntimeError(
            f"[PROTECAO] Tentativa de escrever fora de PASTA_DADOS, bloqueada: {alvo}"
        )

    relativo = os.path.relpath(alvo, raiz)
    partes = relativo.split(os.sep)
    primeiro_nivel = partes[0]

    permitido = primeiro_nivel in _SUBPASTAS_PERMITIDAS or (
        len(partes) == 1 and primeiro_nivel == NOME_LOG_ERROS
    )
    if not permitido:
        raise RuntimeError(
            f"[PROTECAO] Caminho fora das subpastas conhecidas {sorted(_SUBPASTAS_PERMITIDAS)} "
            f"ou do log de erros, bloqueado: {alvo}"
        )


# ---------- Filtros do CSV --------------------------------------------------
TIPO_CLASSE_ALVO = "fidc"
SITUACOES_ATIVAS = {"em funcionamento normal", "fase pre-operacional"}

# ---------- Controle de execucao --------------------------------------------
LIMITE_FUNDOS: Optional[int] = None
DELAY_ENTRE_ENTIDADES_SEG = (0.3, 0.8)

# ---------- HTTP ------------------------------------------------------------
MAX_TENTATIVAS = 4
BACKOFF_INICIAL_SEG = 1.5
TIMEOUT_SEG = 25
FALHAS_SEGUIDAS_PARA_PAUSA = 5
PAUSA_LONGA_SEG = 45

USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
)

BASE_URL = "https://fnet.bmfbovespa.com.br/fnet/publico"
URL_PAGINA_FUNDO = f"{BASE_URL}/abrirGerenciadorDocumentosCVM"
URL_LISTAGEM = f"{BASE_URL}/pesquisarGerenciadorDocumentosDados"
URL_DOWNLOAD = f"{BASE_URL}/downloadDocumento"

# ============================================================================
# Helpers de texto
# ============================================================================


def normalizar(texto: str) -> str:
    if texto is None:
        return ""
    sem_acento = unicodedata.normalize("NFKD", texto).encode("ascii", "ignore").decode("ascii")
    return " ".join(sem_acento.lower().split())


def so_digitos(texto: str) -> str:
    return re.sub(r"\D", "", texto or "")


def data_ref_para_nome(data_ref: str) -> str:
    """'09/2026' -> '09-2026', '01/10/2026' -> '01-10-2026'."""
    return data_ref.strip().replace("/", "-")


# ============================================================================
# Leitura do CSV de classes
# ============================================================================


@dataclass
class FundoClasse:
    cnpj: str           # CNPJ da classe (14 digitos)
    cnpj_fundo: str     # CNPJ do fundo pai (14 digitos)
    nome: str
    tipo_classe: str
    situacao: str


def encontrar_arquivo_classe(pasta: str) -> str:
    if not os.path.isdir(pasta):
        raise FileNotFoundError(f"Pasta nao encontrada: {pasta}")
    candidatos = []
    for raiz, _subpastas, arquivos in os.walk(pasta):
        for n in arquivos:
            if "classe" in normalizar(n) and n.lower().endswith((".csv", ".txt")):
                candidatos.append(os.path.join(raiz, n))
    if not candidatos:
        raise FileNotFoundError(
            f"Nenhum .csv/.txt com 'classe' no nome em {pasta} (busca recursiva). "
            f"Arquivos na raiz: {os.listdir(pasta)}"
        )
    if len(candidatos) > 1:
        print(f"[AVISO] Varios candidatos: {candidatos}. Usando {candidatos[0]}.")
    return candidatos[0]


def ler_csv_com_fallback(caminho: str) -> tuple[list[str], list[list[str]]]:
    for enc in ("utf-8-sig", "cp1252", "latin-1"):
        try:
            with open(caminho, "r", encoding=enc, newline="") as f:
                amostra = f.read(4096)
                f.seek(0)
                delim = ";" if amostra.count(";") >= amostra.count(",") else ","
                linhas = list(csv.reader(f, delimiter=delim))
            print(f"[INFO] CSV lido com encoding '{enc}' (delim '{delim}').")
            return linhas[0], linhas[1:]
        except UnicodeDecodeError:
            continue
    raise UnicodeDecodeError("csv", b"", 0, 1, "Nao decodificou o arquivo classe")


def _encontrar_coluna(col_map: dict[str, int], nomes: list[str]) -> Optional[int]:
    """Retorna o indice da primeira coluna cujo nome (normalizado) bate."""
    for n in nomes:
        for k, v in col_map.items():
            if normalizar(k) == normalizar(n):
                return v
    return None


def carregar_fidcs_ativos(caminho_classe: str) -> list[FundoClasse]:
    header, linhas = ler_csv_com_fallback(caminho_classe)
    col = {nome.strip(): i for i, nome in enumerate(header)}

    obrigatorias = ["CNPJ_Classe", "Denominacao_Social", "Tipo_Classe", "Situacao"]
    faltando = [c for c in obrigatorias if c not in col]
    if faltando:
        raise ValueError(f"Colunas esperadas nao encontradas: {faltando}. Disponiveis: {list(col.keys())}")

    # Coluna do CNPJ do fundo pai (tenta varios nomes comuns)
    idx_cnpj_fundo = _encontrar_coluna(col, ["CNPJ_Fundo", "CNPJ", "CNPJ_Companhia"])
    if idx_cnpj_fundo is None:
        print(
            "[AVISO] Coluna CNPJ do fundo nao encontrada no CSV "
            f"(tentei CNPJ_Fundo / CNPJ / CNPJ_Companhia). "
            "Documentos de nivel 'fundo' serao ignorados."
        )

    vistos: dict[str, FundoClasse] = {}
    for linha in linhas:
        if len(linha) <= max(col.values()):
            continue

        tipo_classe = linha[col["Tipo_Classe"]]
        situacao = linha[col["Situacao"]]

        if TIPO_CLASSE_ALVO not in normalizar(tipo_classe):
            continue
        if normalizar(situacao) not in SITUACOES_ATIVAS:
            continue

        cnpj = so_digitos(linha[col["CNPJ_Classe"]]).zfill(14)
        if len(cnpj) != 14:
            continue

        cnpj_fundo = ""
        if idx_cnpj_fundo is not None and idx_cnpj_fundo < len(linha):
            cnpj_fundo = so_digitos(linha[idx_cnpj_fundo]).zfill(14)
            if len(cnpj_fundo) != 14:
                cnpj_fundo = ""

        nome = linha[col["Denominacao_Social"]].strip()
        vistos[cnpj] = FundoClasse(
            cnpj=cnpj, cnpj_fundo=cnpj_fundo, nome=nome,
            tipo_classe=tipo_classe.strip(), situacao=situacao.strip(),
        )

    return sorted(vistos.values(), key=lambda f: f.cnpj)


# ============================================================================
# Janela de competencia (mensal)
# ============================================================================


def mes_cutoff_absoluto(hoje: date, meses_retroativos: int) -> int:
    total = hoje.year * 12 + (hoje.month - 1)
    return total - meses_retroativos


def referencia_mensal_para_abs(data_ref: str) -> Optional[int]:
    """'09/2026' -> ano*12+mes. Retorna None se formato != MM/YYYY."""
    m = re.fullmatch(r"(\d{2})/(\d{4})", (data_ref or "").strip())
    if not m:
        return None
    mes, ano = int(m.group(1)), int(m.group(2))
    if not (1 <= mes <= 12):
        return None
    return ano * 12 + (mes - 1)


# ============================================================================
# HTTP com retry/backoff
# ============================================================================


class FalhaDefinitiva(Exception):
    pass


def requisitar(
    sessao: requests.Session,
    url: str,
    *,
    params: Optional[dict] = None,
    esperar_json: bool = False,
    contexto: str = "",
) -> requests.Response:
    ultimo_erro: Optional[Exception] = None
    espera = BACKOFF_INICIAL_SEG

    for t in range(1, MAX_TENTATIVAS + 1):
        try:
            resp = sessao.get(url, params=params, timeout=TIMEOUT_SEG)
        except requests.exceptions.RequestException as e:
            ultimo_erro = e
            print(f"    [tentativa {t}/{MAX_TENTATIVAS}] erro de conexao em {contexto}: {e}")
        else:
            bloqueado = resp.status_code in (403, 429, 503)
            tipo_errado = esperar_json and "json" not in resp.headers.get("content-type", "").lower()

            if resp.status_code == 200 and not tipo_errado:
                return resp

            if bloqueado:
                print(f"    [tentativa {t}/{MAX_TENTATIVAS}] bloqueio/rate-limit ({resp.status_code}) em {contexto}")
            elif tipo_errado:
                print(f"    [tentativa {t}/{MAX_TENTATIVAS}] resposta nao-JSON em {contexto}")
            else:
                print(f"    [tentativa {t}/{MAX_TENTATIVAS}] HTTP {resp.status_code} em {contexto}")
            ultimo_erro = RuntimeError(f"status={resp.status_code}")

        if t < MAX_TENTATIVAS:
            time.sleep(espera + random.uniform(0, 0.5))
            espera *= 2

    raise FalhaDefinitiva(f"Falhou apos {MAX_TENTATIVAS} tentativas em {contexto}: {ultimo_erro}")


# ============================================================================
# Extensao do arquivo baixado
# ============================================================================


def extensao_do_download(resp: requests.Response) -> str:
    cd = resp.headers.get("content-disposition", "")
    m = re.search(r'filename\*?=(?:UTF-8\'\')?"?([^";]+)"?', cd)
    if m:
        _, ext = os.path.splitext(m.group(1))
        if ext:
            return ext
    ct = resp.headers.get("content-type", "").split(";")[0].strip().lower()
    mapa = {"text/xml": ".xml", "application/xml": ".xml", "application/pdf": ".pdf",
            "application/zip": ".zip", "text/html": ".html"}
    return mapa.get(ct, ".bin")


# ============================================================================
# Listagem e filtragem de documentos
# ============================================================================


@dataclass
class DocumentoAlvo:
    id: int
    cnpj: str
    data_ref_nome: str   # formatado p/ nome do arquivo ("09-2026" ou "01-10-2026")
    prefixo: str         # prefixo do tipo (ex: "InformeMensal")


_CHAVES_TIPO = (
    "tipoDocumento", "descricaoTipo", "tipo", "dsTipo",
    "descricaoTipoDocumento", "nomeTipoDocumento",
)
_CHAVES_CATEGORIA = (
    "categoriaDocumento", "descricaoCategoria", "categoria", "dsCategoria",
    "descricaoCategoriaDocumento", "nomeCategoriaDocumento",
)
_DIAG_CAMPOS_EMITIDO = False


def _campo_do_registro(reg: dict, campo: str) -> str:
    """Extrai o valor normalizado do campo do JSON, tentando variantes de nome.
    Se nenhuma chave conhecida for encontrada, faz busca fuzzy por chaves que
    contenham 'tipo' ou 'categor' no nome."""
    global _DIAG_CAMPOS_EMITIDO
    chaves = _CHAVES_TIPO if campo == "tipo" else _CHAVES_CATEGORIA

    # 1) Tentativa direta nas chaves conhecidas
    for k in chaves:
        v = reg.get(k)
        if v:
            return normalizar(v)

    # 2) Fallback: qualquer chave cujo nome contenha a palavra-chave
    palavra = "tipo" if campo == "tipo" else "categor"
    for k, v in reg.items():
        if palavra in k.lower() and isinstance(v, str) and v.strip():
            if not _DIAG_CAMPOS_EMITIDO:
                print(
                    f"  [DIAG] Campo '{campo}' encontrado via fallback: "
                    f"chave='{k}', valor='{v}'",
                    flush=True,
                )
            return normalizar(v)

    return ""


_AVISO_STATUS_EMITIDO = False
_CAMPOS_STATUS = ("situacaoDocumento", "descricaoStatus", "status", "statusDocumento")


_CODIGOS_ATIVO = {"a", "ac"}  # codigos curtos que significam "ativo"


def _status_ativo(reg: dict) -> bool:
    """True se o documento esta com status Ativo.

    Aceita tanto o texto completo ('Ativo', 'Ativo com visualizacao', ...)
    quanto codigos curtos ('A', 'AC') usados pela API.

    Se nenhuma das chaves conhecidas existir no registro, NAO filtra
    (assume ativo) e avisa uma vez.
    """
    global _AVISO_STATUS_EMITIDO
    if not any(c in reg for c in _CAMPOS_STATUS):
        if not _AVISO_STATUS_EMITIDO:
            print(
                "[AVISO] Campo de status do documento nao encontrado no JSON "
                f"(tentei {_CAMPOS_STATUS}). Filtro de status DESATIVADO - "
                "confira o nome do campo na resposta da API."
            )
            _AVISO_STATUS_EMITIDO = True
        return True

    # Testa TODOS os campos individualmente em vez de parar no primeiro truthy
    for campo in _CAMPOS_STATUS:
        valor = normalizar(reg.get(campo, "") or "")
        if not valor:
            continue
        # Texto completo: "ativo", "ativo com visualizacao", etc.
        if valor.startswith("ativo"):
            return True
        # Codigo curto: "a", "ac"
        if valor in _CODIGOS_ATIVO:
            return True

    # Se chegou aqui, todos os campos existentes tem valores != ativo
    return False


def buscar_todos_registros(sessao: requests.Session, cnpj: str) -> list[dict]:
    """Ancora a sessao no CNPJ e retorna todos os registros de documentos
    (ate 5 paginas de 200)."""
    requisitar(sessao, URL_PAGINA_FUNDO, params={"cnpjFundo": cnpj},
               contexto=f"pagina do CNPJ {cnpj}")

    registros: list[dict] = []
    tam = 200
    inicio = 0

    for _ in range(5):
        resp = requisitar(
            sessao, URL_LISTAGEM,
            params={
                "d": 1, "s": inicio, "l": tam,
                "o[0][dataReferencia]": "desc",
                "idCategoriaDocumento": 0, "idTipoDocumento": 0,
                "idEspecieDocumento": 0, "isSession": "true",
                "_": int(time.time() * 1000),
            },
            esperar_json=True,
            contexto=f"listagem {cnpj} (offset {inicio})",
        )
        try:
            dados = resp.json()
        except ValueError as e:
            raise FalhaDefinitiva(f"JSON invalido na listagem de {cnpj}: {e}")

        lote = dados.get("data", [])
        if not lote:
            break
        registros.extend(lote)
        inicio += tam
        if inicio >= dados.get("recordsFiltered", 0):
            break

    return registros


def filtrar_documentos(
    registros: list[dict],
    tc: TipoDocConfig,
    cnpj: str,
    hoje: date,
) -> list[DocumentoAlvo]:
    """Filtra registros pelo TipoDocConfig e retorna DocumentoAlvo's selecionados."""

    if tc.periodicidade == "mensal":
        return _filtrar_mensal(registros, tc, cnpj, hoje)
    else:
        return _filtrar_avulso(registros, tc, cnpj)


def _filtrar_mensal(
    registros: list[dict], tc: TipoDocConfig, cnpj: str, hoje: date,
) -> list[DocumentoAlvo]:
    cutoff = mes_cutoff_absoluto(hoje, tc.meses_retroativos)
    por_comp: dict[int, DocumentoAlvo] = {}

    for reg in registros:
        if tc.nome_filtro not in _campo_do_registro(reg, tc.campo_filtro):
            continue
        if not _status_ativo(reg):
            continue
        doc_id = reg.get("id")
        if doc_id is None:
            continue

        data_ref = (reg.get("dataReferencia") or "").strip()
        comp_abs = referencia_mensal_para_abs(data_ref)
        if comp_abs is None or comp_abs < cutoff:
            continue

        candidato = DocumentoAlvo(
            id=doc_id, cnpj=cnpj,
            data_ref_nome=data_ref_para_nome(data_ref),
            prefixo=tc.prefixo_arquivo,
        )
        existente = por_comp.get(comp_abs)
        if existente is None or candidato.id > existente.id:
            por_comp[comp_abs] = candidato

    return sorted(por_comp.values(), key=lambda d: d.id, reverse=True)


def _filtrar_avulso(
    registros: list[dict], tc: TipoDocConfig, cnpj: str,
) -> list[DocumentoAlvo]:
    por_data: dict[str, DocumentoAlvo] = {}

    for reg in registros:
        if tc.nome_filtro not in _campo_do_registro(reg, tc.campo_filtro):
            continue
        if not _status_ativo(reg):
            continue
        doc_id = reg.get("id")
        if doc_id is None:
            continue

        data_ref = (reg.get("dataReferencia") or "").strip()
        if not data_ref:
            continue

        candidato = DocumentoAlvo(
            id=doc_id, cnpj=cnpj,
            data_ref_nome=data_ref_para_nome(data_ref),
            prefixo=tc.prefixo_arquivo,
        )
        existente = por_data.get(data_ref)
        if existente is None or candidato.id > existente.id:
            por_data[data_ref] = candidato

    resultado = sorted(por_data.values(), key=lambda d: d.id, reverse=True)
    return resultado[: tc.max_recentes]


# ============================================================================
# Download
# ============================================================================


def baixar_documento(
    sessao: requests.Session, doc: DocumentoAlvo, pasta_saida: str,
) -> tuple[str, str]:
    """Retorna (status, nome_arquivo) com status 'baixado' ou 'existente'.
    Lanca FalhaDefinitiva se falhar."""
    prefixo = f"{doc.cnpj}_{doc.prefixo}_{doc.data_ref_nome}"
    existentes = [n for n in os.listdir(pasta_saida) if n.startswith(prefixo)]
    if existentes:
        return "existente", existentes[0]

    resp = requisitar(
        sessao, URL_DOWNLOAD, params={"id": doc.id},
        contexto=f"download doc {doc.id} ({doc.cnpj}, {doc.data_ref_nome})",
    )
    ext = extensao_do_download(resp)
    nome_arquivo = f"{prefixo}{ext}"
    caminho_final = os.path.join(pasta_saida, nome_arquivo)
    caminho_tmp = caminho_final + ".part"

    _validar_escrita(caminho_final)

    with open(caminho_tmp, "wb") as f:
        f.write(resp.content)
    os.replace(caminho_tmp, caminho_final)
    return "baixado", nome_arquivo


# ============================================================================
# Processamento de uma entidade (classe ou fundo)
# ============================================================================


def processar_entidade(
    sessao: requests.Session,
    cnpj: str,
    nome: str,
    tipos: list[TipoDocConfig],
    hoje: date,
    pasta_base: str,
    estat: dict,
    erros: list,
    label: str,
) -> bool:
    """Processa todos os tipos de documento para um CNPJ.
    Retorna True se houve erro definitivo (pra circuit breaker)."""
    print(f"\n  {label} {cnpj} - {nome}", flush=True)

    try:
        registros = buscar_todos_registros(sessao, cnpj)
    except FalhaDefinitiva as e:
        print(f"    ERRO ao listar documentos: {e}", flush=True)
        erros.append((cnpj, nome, f"listagem: {e}"))
        estat["entidades_erro"] += 1
        return True  # falha

    print(f"    {len(registros)} registro(s) de documento encontrados no FNET", flush=True)

    # ---- Diagnostico (1a entidade com registros): mostra chaves reais do JSON
    global _DIAG_CAMPOS_EMITIDO
    if registros and not _DIAG_CAMPOS_EMITIDO:
        r0 = registros[0]
        print(f"    [DIAG] Chaves do 1o registro: {sorted(r0.keys())}", flush=True)
        # Mostra valores de campos que parecem ser tipo/categoria/status
        for k, v in sorted(r0.items()):
            kl = k.lower()
            if any(p in kl for p in ("tipo", "categor", "status", "situacao", "especie")):
                print(f"    [DIAG]   {k} = {v!r}", flush=True)
        _DIAG_CAMPOS_EMITIDO = True

    teve_doc = False
    teve_erro = False
    baixados_aqui = 0
    existentes_aqui = 0

    for tc in tipos:
        pasta_saida = os.path.join(pasta_base, tc.pasta)
        _validar_escrita(pasta_saida)
        os.makedirs(pasta_saida, exist_ok=True)

        docs = filtrar_documentos(registros, tc, cnpj, hoje)
        if not docs:
            print(f"    [{tc.prefixo_arquivo}] nenhum documento elegivel", flush=True)
            continue
        teve_doc = True
        print(f"    [{tc.prefixo_arquivo}] {len(docs)} documento(s) selecionado(s)", flush=True)

        for doc in docs:
            try:
                resultado, nome_arquivo = baixar_documento(sessao, doc, pasta_saida)
            except FalhaDefinitiva as e:
                print(f"      ERRO [{tc.prefixo_arquivo}] {doc.data_ref_nome} (id={doc.id}): {e}", flush=True)
                erros.append((cnpj, nome, f"{tc.prefixo_arquivo} {doc.data_ref_nome}: {e}"))
                teve_erro = True
                continue

            if resultado == "baixado":
                print(f"      BAIXADO [{tc.prefixo_arquivo}] {doc.data_ref_nome} -> {nome_arquivo}", flush=True)
                estat["baixados"] += 1
                baixados_aqui += 1
            else:
                print(f"      JA EXISTE [{tc.prefixo_arquivo}] {doc.data_ref_nome} -> {nome_arquivo}", flush=True)
                estat["existentes"] += 1
                existentes_aqui += 1

    print(f"    resumo: {baixados_aqui} baixado(s), {existentes_aqui} ja existente(s)"
          f"{', com erro' if teve_erro else ''}", flush=True)

    if teve_erro:
        estat["entidades_erro"] += 1
    elif teve_doc:
        estat["entidades_ok"] += 1
    else:
        estat["entidades_sem_doc"] += 1

    return teve_erro


# ============================================================================
# Main
# ============================================================================


def main() -> None:
    inicio_execucao = time.time()
    hoje = date.today()

    tipos_classe = [t for t in TIPOS_DOCUMENTO if t.nivel == "classe"]
    tipos_fundo = [t for t in TIPOS_DOCUMENTO if t.nivel == "fundo"]

    print("=" * 78)
    print("FNET scraper - Documentos FIDC")
    print("=" * 78)
    print(f"Hoje: {hoje.isoformat()}")
    print(f"Tipos de documento: {len(TIPOS_DOCUMENTO)}")
    for tc in TIPOS_DOCUMENTO:
        if tc.periodicidade == "mensal":
            janela = f"ultimos {tc.meses_retroativos} meses"
        else:
            janela = f"{tc.max_recentes} mais recentes"
        print(f"  - {tc.prefixo_arquivo:20s} (nivel {tc.nivel:6s}, {janela}) -> {tc.pasta}/")

    caminho_classe = encontrar_arquivo_classe(PASTA_REGISTRO_CLASSE)
    print(f"\nArquivo de classes: {caminho_classe}")

    pasta_base = PASTA_DADOS  # pastas de saida sempre aqui, independente de onde o csv foi achado
    fundos = carregar_fidcs_ativos(caminho_classe)
    if LIMITE_FUNDOS is not None:
        fundos = fundos[: LIMITE_FUNDOS]

    print(f"FIDCs ativos/pre-op: {len(fundos)}")
    print("-" * 78)

    sessao = requests.Session()
    sessao.headers.update({
        "User-Agent": USER_AGENT,
        "Referer": URL_PAGINA_FUNDO,
        "X-Requested-With": "XMLHttpRequest",
        "Accept": "application/json, text/javascript, */*; q=0.01",
    })

    erros: list[tuple[str, str, str]] = []
    estat = {"entidades_ok": 0, "entidades_sem_doc": 0, "entidades_erro": 0,
             "baixados": 0, "existentes": 0}
    falhas_seguidas = 0

    def _checar_circuit_breaker():
        nonlocal falhas_seguidas
        falhas_seguidas += 1
        if falhas_seguidas >= FALHAS_SEGUIDAS_PARA_PAUSA:
            print(f"    [AVISO] {falhas_seguidas} falhas seguidas -- "
                  f"pausando {PAUSA_LONGA_SEG}s")
            time.sleep(PAUSA_LONGA_SEG)
            falhas_seguidas = 0

    try:
        # ---- FASE 1: documentos de nivel classe ----------------------------
        if tipos_classe:
            print(f"\n=== FASE 1: documentos de nivel CLASSE ({len(fundos)} classes) ===\n")
            for i, fundo in enumerate(fundos, 1):
                teve_erro = processar_entidade(
                    sessao, fundo.cnpj, fundo.nome, tipos_classe,
                    hoje, pasta_base, estat, erros,
                    label=f"[classe {i}/{len(fundos)}]",
                )
                if teve_erro:
                    _checar_circuit_breaker()
                else:
                    falhas_seguidas = 0
                time.sleep(random.uniform(*DELAY_ENTRE_ENTIDADES_SEG))

        # ---- FASE 2: documentos de nivel fundo -----------------------------
        if tipos_fundo:
            # Deduplica por CNPJ do fundo
            fundos_unicos: dict[str, FundoClasse] = {}
            for f in fundos:
                if f.cnpj_fundo and f.cnpj_fundo not in fundos_unicos:
                    fundos_unicos[f.cnpj_fundo] = f
            lista_fundos = sorted(fundos_unicos.values(), key=lambda x: x.cnpj_fundo)

            if not lista_fundos:
                print("\n[AVISO] Nenhum CNPJ de fundo disponivel no CSV "
                      "-- pulando documentos de nivel fundo.")
            else:
                print(f"\n=== FASE 2: documentos de nivel FUNDO "
                      f"({len(lista_fundos)} fundos unicos) ===\n")
                falhas_seguidas = 0
                for i, fundo in enumerate(lista_fundos, 1):
                    teve_erro = processar_entidade(
                        sessao, fundo.cnpj_fundo, fundo.nome, tipos_fundo,
                        hoje, pasta_base, estat, erros,
                        label=f"[fundo {i}/{len(lista_fundos)}]",
                    )
                    if teve_erro:
                        _checar_circuit_breaker()
                    else:
                        falhas_seguidas = 0
                    time.sleep(random.uniform(*DELAY_ENTRE_ENTIDADES_SEG))

    except KeyboardInterrupt:
        print("\n[INTERROMPIDO] Salvando progresso...")

    # ---- Resumo ------------------------------------------------------------
    if erros:
        caminho_erros = os.path.join(pasta_base, NOME_LOG_ERROS)
        _validar_escrita(caminho_erros)
        with open(caminho_erros, "w", encoding="utf-8") as f:
            for cnpj, nome, motivo in erros:
                f.write(f"{cnpj}\t{nome}\t{motivo}\n")
        print(f"\n{len(erros)} erro(s) em: {caminho_erros}")

    duracao = time.time() - inicio_execucao
    print("-" * 78)
    print("RESUMO")
    print(f"  entidades processadas OK ......... {estat['entidades_ok']}")
    print(f"  entidades sem documento .......... {estat['entidades_sem_doc']}")
    print(f"  entidades com algum erro ......... {estat['entidades_erro']}")
    print(f"  arquivos baixados ................ {estat['baixados']}")
    print(f"  arquivos ja existentes (pulados) . {estat['existentes']}")
    print(f"  tempo total ...................... {duracao / 60:.1f} min")
    print("=" * 78)


if __name__ == "__main__":
    main()