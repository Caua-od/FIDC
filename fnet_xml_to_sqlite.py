#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
fnet_xml_to_sqlite.py
======================

Le os arquivos XML de Informe Mensal de FIDC baixados pelo scraper
(fnet_scraper_fidc.py) e armazena os dados em um banco SQLite local,
criando uma estrutura tabular facil de consultar.

ESTRATEGIA
----------
O XML do Informe Mensal FIDC (ICVM 489/576) e complexo e muda entre
versoes. Em vez de mapear cada tag numa coluna fixa, o parser extrai
**todos** os pares tag-valor recursivamente, guardando:

  - documentos     : metadados do arquivo (cnpj, competencia, etc.)
  - campos         : chave-valor (path completo da tag + valor)
  - cedentes       : lista de cedentes (tags CEDENT e CEDENT_CRED_EXISTE,
                      com a coluna 'origem' indicando de qual secao vieram)
  - devedores      : lista de devedores/sacados (tag DEV_SACADOS dentro
                      de LISTA_DEV_SACADOS) -- bloco OPCIONAL do informe;
                      e normal ficar vazia para boa parte dos documentos

Isso permite:
  * Funcionar com qualquer versao do XML (489, 576, futuras)
  * Zero modelagem manual por campo
  * Queries SQL sobre qualquer campo:
      SELECT valor FROM campos
      WHERE nome_campo = 'APLIC_ATIVO.VL_CARTEIRA'
        AND documento_id = X

RE-EXECUCAO SEGURA
------------------
Cada arquivo e identificado por hash SHA-256. Se o hash ja existe
no banco, o arquivo e ignorado. Novas execucoes processam somente
os XMLs ainda nao registrados.

BACKUP ROTATIVO
----------------
Antes de processar novos XMLs, o script faz uma copia de seguranca do
estado atual do banco em tres formatos, na pasta "backups" (dentro da
pasta do banco):

    backups/sqlite/fidc_informes_AAAAMMDD_HHMMSS.db
    backups/csv/AAAAMMDD_HHMMSS/<tabela>.csv
    backups/parquet/AAAAMMDD_HHMMSS/<tabela>.parquet

Mantem apenas as 3 copias mais recentes de cada formato; ao gerar uma
4a, a mais antiga e apagada. O SQLite continua sendo a fonte principal
de consulta (usada pelo backend/HTML); CSV e Parquet servem apenas
como copias de seguranca legiveis por outras ferramentas.

Se o pacote opcional "pyarrow" nao estiver instalado, o backup em
Parquet e pulado (com aviso) -- o restante do fluxo continua normal.

REQUISITOS
----------
    Nenhum obrigatorio alem da stdlib (sqlite3, xml.etree, hashlib, etc.)
    Opcional: pip install pyarrow   (para backup em formato Parquet)

USO
---
    python fnet_xml_to_sqlite.py [--pasta-xml CAMINHO] [--banco CAMINHO]

    Sem argumentos, usa os caminhos padrao definidos na secao CONFIG.
"""

from __future__ import annotations

import argparse
import csv as csv_mod
import hashlib
import os
import re
import shutil
import sqlite3
import sys
import traceback
import unicodedata
from datetime import datetime
from xml.etree import ElementTree as ET

try:
    import pyarrow as pa
    import pyarrow.parquet as pq
    _PYARROW_DISPONIVEL = True
except ImportError:
    _PYARROW_DISPONIVEL = False


# ============================================================================
# CONFIG
# ============================================================================

# Pasta onde o scraper salva os XMLs de Informe Mensal
PASTA_XML_PADRAO = r"C:\Users\olive\Downloads\Franklin project\Data\Scrapping\xml"

# Caminho do banco SQLite de saida
BANCO_PADRAO = r"C:\Users\olive\Downloads\Franklin project\Data\Scrapping\Compilado xml\fidc_informes.db"

# Quantas copias de backup manter por formato (sqlite/csv/parquet)
MAX_BACKUPS_POR_FORMATO = 3

# Nome das tabelas a incluir no backup CSV/Parquet
TABELAS_BACKUP = ["documentos", "campos", "cedentes", "devedores"]

# Tags que contem listas repetitivas (cedentes / devedores)
# e que devem ir para tabelas proprias em vez de 'campos'
TAGS_LISTA_CEDENTES = {"CEDENT", "CEDENT_CRED_EXISTE"}
# DEV_SACADOS (dentro de LISTA_DEV_SACADOS) e um bloco OPCIONAL do informe:
# alguns fundos/competencias nao o preenchem, entao e normal a tabela
# 'devedores' ficar vazia para boa parte dos documentos.
TAGS_LISTA_DEVEDORES = {"DEV_SACADOS"}

# Tags a ignorar na extracao recursiva. CAB_INFORM NAO entra aqui: seus
# campos principais (competencia, CNPJs, nome da classe, versao) ja viram
# colunas proprias em 'documentos' (ver inserir_documento), mas o restante
# (ex: PR_ENTRE_RESGATE, CLASS_UNICA, NR_CNPJ_ADM) ainda e util e deve ser
# preservado em 'campos' em vez de descartado.
TAGS_IGNORAR: set[str] = set()

# Limite de profundidade na recursao (protecao contra XML malformado)
MAX_PROFUNDIDADE = 30

# Log de erros
NOME_LOG_ERROS = "erros_parser.log"


# ============================================================================
# Helpers
# ============================================================================


def normalizar_texto(texto: str) -> str:
    """Remove acentos e normaliza espacos."""
    if not texto:
        return ""
    sem_acento = unicodedata.normalize("NFKD", texto).encode("ascii", "ignore").decode("ascii")
    return " ".join(sem_acento.lower().split())


def so_digitos(texto: str) -> str:
    return re.sub(r"\D", "", texto or "")


def calcular_hash(caminho: str) -> str:
    """SHA-256 do conteudo do arquivo."""
    h = hashlib.sha256()
    with open(caminho, "rb") as f:
        while True:
            bloco = f.read(8192)
            if not bloco:
                break
            h.update(bloco)
    return h.hexdigest()


def extrair_cnpj_do_nome(nome_arquivo: str) -> str:
    """Extrai CNPJ (14 digitos) do nome no padrao do scraper:
    <cnpj14>_InformeMensal_<data>.xml"""
    partes = nome_arquivo.split("_")
    if partes:
        candidato = so_digitos(partes[0])
        if len(candidato) == 14:
            return candidato
    return ""


def extrair_competencia_do_nome(nome_arquivo: str) -> str:
    """Extrai competencia do nome no padrao do scraper:
    <cnpj>_InformeMensal_09-2026.xml -> 09/2026"""
    match = re.search(r"_(\d{2})-(\d{4})\.", nome_arquivo)
    if match:
        return f"{match.group(1)}/{match.group(2)}"
    # Tenta formato DD-MM-YYYY
    match = re.search(r"_(\d{2})-(\d{2})-(\d{4})\.", nome_arquivo)
    if match:
        return f"{match.group(1)}/{match.group(2)}/{match.group(3)}"
    return ""


def _tag_local(tag: str) -> str:
    """Remove namespace de uma tag XML: {ns}Nome -> Nome"""
    if "}" in tag:
        return tag.split("}", 1)[1]
    return tag


# ============================================================================
# Schema SQLite
# ============================================================================

SQL_CRIAR_TABELAS = """
CREATE TABLE IF NOT EXISTS documentos (
    id              INTEGER PRIMARY KEY AUTOINCREMENT,
    nome_arquivo    TEXT    NOT NULL,
    hash_arquivo    TEXT    NOT NULL UNIQUE,
    cnpj            TEXT    NOT NULL,
    cnpj_fundo      TEXT    NOT NULL DEFAULT '',
    nome_classe     TEXT    NOT NULL DEFAULT '',
    competencia     TEXT    NOT NULL DEFAULT '',
    tipo_condominio TEXT    NOT NULL DEFAULT '',
    fdo_exclusivo   TEXT    NOT NULL DEFAULT '',
    versao_layout   TEXT    NOT NULL DEFAULT '',
    dt_processamento TEXT   NOT NULL,
    caminho_xml     TEXT    NOT NULL DEFAULT ''
);

CREATE TABLE IF NOT EXISTS campos (
    id              INTEGER PRIMARY KEY AUTOINCREMENT,
    documento_id    INTEGER NOT NULL,
    secao           TEXT    NOT NULL DEFAULT '',
    nome_campo      TEXT    NOT NULL,
    valor           TEXT    NOT NULL DEFAULT '',
    FOREIGN KEY (documento_id) REFERENCES documentos(id)
);

CREATE TABLE IF NOT EXISTS cedentes (
    id              INTEGER PRIMARY KEY AUTOINCREMENT,
    documento_id    INTEGER NOT NULL,
    origem          TEXT    NOT NULL DEFAULT '',
    nr_pf_pj        TEXT    NOT NULL DEFAULT '',
    pr_participacao TEXT    NOT NULL DEFAULT '',
    FOREIGN KEY (documento_id) REFERENCES documentos(id)
);

CREATE TABLE IF NOT EXISTS devedores (
    id              INTEGER PRIMARY KEY AUTOINCREMENT,
    documento_id    INTEGER NOT NULL,
    nr_pf_pj        TEXT    NOT NULL DEFAULT '',
    valor           TEXT    NOT NULL DEFAULT '',
    pr_participacao TEXT    NOT NULL DEFAULT '',
    FOREIGN KEY (documento_id) REFERENCES documentos(id)
);

CREATE INDEX IF NOT EXISTS idx_campos_doc      ON campos(documento_id);
CREATE INDEX IF NOT EXISTS idx_campos_nome     ON campos(nome_campo);
CREATE INDEX IF NOT EXISTS idx_campos_secao    ON campos(secao);
CREATE INDEX IF NOT EXISTS idx_cedentes_doc    ON cedentes(documento_id);
CREATE INDEX IF NOT EXISTS idx_devedores_doc   ON devedores(documento_id);
CREATE INDEX IF NOT EXISTS idx_documentos_cnpj ON documentos(cnpj);
CREATE INDEX IF NOT EXISTS idx_documentos_comp ON documentos(competencia);
CREATE INDEX IF NOT EXISTS idx_documentos_hash ON documentos(hash_arquivo);
"""


# ============================================================================
# Parser XML
# ============================================================================


class DadosExtraidos:
    """Container para os dados extraidos de um unico XML."""

    def __init__(self) -> None:
        self.cabecalho: dict[str, str] = {}
        self.campos: list[tuple[str, str, str]] = []   # (secao, path, valor)
        self.cedentes: list[dict[str, str]] = []
        self.devedores: list[dict[str, str]] = []


def _extrair_cabecalho(raiz: ET.Element) -> dict[str, str]:
    """Extrai campos do CAB_INFORM (cabecalho do documento)."""
    cab: dict[str, str] = {}

    for elem in raiz.iter():
        tag = _tag_local(elem.tag)
        if tag == "CAB_INFORM":
            for filho in elem:
                t = _tag_local(filho.tag)
                v = (filho.text or "").strip()
                cab[t] = v
            break

    return cab


def _extrair_recursivo(
    elem: ET.Element,
    dados: DadosExtraidos,
    caminho: str = "",
    profundidade: int = 0,
) -> None:
    """Percorre a arvore XML recursivamente extraindo campos folha."""
    if profundidade > MAX_PROFUNDIDADE:
        return

    tag = _tag_local(elem.tag)
    caminho_atual = f"{caminho}.{tag}" if caminho else tag

    # Pular o cabecalho (ja extraido separadamente)
    if tag in TAGS_IGNORAR:
        return

    # Tratar listas de cedentes (CEDENT dentro de LISTA_CEDENT, ou
    # CEDENT_CRED_EXISTE dentro de LISTA_CEDENT_CRED_EXISTE)
    if tag in TAGS_LISTA_CEDENTES:
        cedente: dict[str, str] = {}
        for filho in elem:
            t = _tag_local(filho.tag)
            v = (filho.text or "").strip()
            if "PF_PJ" in t or "CNPJ" in t or "CPF" in t:
                cedente["nr_pf_pj"] = v
            elif "PR_" in t:
                cedente["pr_participacao"] = v
        if cedente:
            cedente["origem"] = tag
            dados.cedentes.append(cedente)
        return

    # Tratar listas de devedores/sacados (bloco opcional -- muitos
    # documentos nao o preenchem, por isso a tabela 'devedores' fica
    # vazia com frequencia)
    if tag in TAGS_LISTA_DEVEDORES:
        devedor: dict[str, str] = {}
        for filho in elem:
            t = _tag_local(filho.tag)
            v = (filho.text or "").strip()
            if "PF_PJ" in t or "CNPJ" in t or "CPF" in t or "SACADO" in t:
                if "VL_" not in t and "PR_" not in t:
                    devedor["nr_pf_pj"] = v
                elif "VL_" in t:
                    devedor["valor"] = v
                elif "PR_" in t:
                    devedor["pr_participacao"] = v
            elif "VL_" in t:
                devedor["valor"] = v
            elif "PR_" in t:
                devedor["pr_participacao"] = v
        if devedor:
            dados.devedores.append(devedor)
        return

    # Se o elemento tem texto (folha), salvar como campo
    texto = (elem.text or "").strip()
    if texto and not list(elem):
        # Determinar a secao (primeiro nivel abaixo de LISTA_INFORM)
        partes = caminho_atual.split(".")
        secao = ""
        for i, p in enumerate(partes):
            if p in ("LISTA_INFORM", "DOC_ARQ"):
                continue
            secao = p
            break

        dados.campos.append((secao, caminho_atual, texto))
        return

    # Recursao nos filhos. Quando um elemento tem varios filhos DIRETOS
    # com o mesmo nome de tag (ex: duas ocorrencias de <CLASSE_SENIOR>
    # dentro do mesmo <NUM_COTISTAS>, uma por serie), um indice [n] e
    # acrescentado ao path de cada ocorrencia para nao misturar os
    # valores das diferentes series/ocorrencias sob o mesmo nome_campo.
    tags_filhos = [_tag_local(f.tag) for f in elem]
    repetidas = {t for t in tags_filhos if tags_filhos.count(t) > 1}
    contadores: dict[str, int] = {}

    for filho in elem:
        tag_filho = _tag_local(filho.tag)
        if tag_filho in repetidas:
            contadores[tag_filho] = contadores.get(tag_filho, 0) + 1
            caminho_filho = f"{caminho_atual}.{tag_filho}[{contadores[tag_filho]}]"
            # Recursao direta nos NETOS, usando o path ja indexado
            # (evita que _extrair_recursivo monte o path sem indice).
            for neto in filho:
                _extrair_recursivo(neto, dados, caminho_filho, profundidade + 2)
            # Caso o proprio elemento repetido seja uma folha com texto
            # (sem filhos), registra o campo diretamente.
            texto_filho = (filho.text or "").strip()
            if texto_filho and not list(filho):
                partes = caminho_filho.split(".")
                secao = next((p for p in partes if p not in ("LISTA_INFORM", "DOC_ARQ")), "")
                dados.campos.append((secao, caminho_filho, texto_filho))
            # Tags de lista (cedentes/devedores) repetidas no mesmo nivel
            # tambem precisam do tratamento especializado:
            if tag_filho in TAGS_LISTA_CEDENTES or tag_filho in TAGS_LISTA_DEVEDORES:
                _extrair_recursivo(filho, dados, caminho_atual, profundidade + 1)
        else:
            _extrair_recursivo(filho, dados, caminho_atual, profundidade + 1)


def parse_xml(caminho: str) -> DadosExtraidos:
    """Faz o parse completo de um arquivo XML de Informe Mensal FIDC."""
    tree = ET.parse(caminho)
    raiz = tree.getroot()

    dados = DadosExtraidos()
    dados.cabecalho = _extrair_cabecalho(raiz)
    _extrair_recursivo(raiz, dados)

    return dados


# ============================================================================
# Banco de dados
# ============================================================================


def criar_banco(caminho_banco: str) -> sqlite3.Connection:
    """Cria (ou abre) o banco SQLite e garante que as tabelas existem."""
    pasta_banco = os.path.dirname(caminho_banco)
    if pasta_banco and not os.path.isdir(pasta_banco):
        os.makedirs(pasta_banco, exist_ok=True)

    conn = sqlite3.connect(caminho_banco)
    conn.execute("PRAGMA journal_mode=WAL")
    conn.execute("PRAGMA foreign_keys=ON")
    conn.executescript(SQL_CRIAR_TABELAS)
    conn.commit()
    return conn


def hash_ja_processado(conn: sqlite3.Connection, hash_arquivo: str) -> bool:
    """Retorna True se o hash ja esta registrado no banco."""
    cursor = conn.execute(
        "SELECT 1 FROM documentos WHERE hash_arquivo = ? LIMIT 1",
        (hash_arquivo,),
    )
    return cursor.fetchone() is not None


def carregar_nomes_ja_processados(conn: sqlite3.Connection) -> set[str]:
    """Carrega em memoria o conjunto de nomes de arquivo ja registrados.

    Usado como filtro rapido antes de calcular o hash SHA-256 (que exige
    ler o arquivo inteiro): como o nome do arquivo gerado pelo scraper e
    deterministico (<cnpj>_<tipo>_<competencia>.xml) e nunca se repete
    para conteudos diferentes, pular o calculo de hash para nomes ja
    vistos evita I/O redundante em pastas com milhares de XMLs.
    """
    cursor = conn.execute("SELECT nome_arquivo FROM documentos")
    return {row[0] for row in cursor.fetchall()}


def inserir_documento(
    conn: sqlite3.Connection,
    nome_arquivo: str,
    hash_arquivo: str,
    cnpj: str,
    dados: DadosExtraidos,
    caminho_xml: str,
) -> int:
    """Insere o documento e seus dados no banco. Retorna o id do documento."""
    cab = dados.cabecalho
    competencia = cab.get("DT_COMPT", "")
    if not competencia:
        competencia = extrair_competencia_do_nome(nome_arquivo)

    cnpj_fundo = cab.get("NR_CNPJ_FUNDO", "")
    # Nome da classe: campo essencial para exibicao (ex: no HTML), mas nao
    # presente no layout original ICVM 489 -- aparece em extensoes mais
    # recentes do informe (pos Resolucao 175). Fica vazio quando ausente.
    nome_classe = cab.get("NM_CLASSE", "")
    tipo_cond = cab.get("TP_CONDOMINIO", "")
    fdo_excl = cab.get("FDO_EXCL", "")
    versao_layout = cab.get("VERSAO", "")

    cursor = conn.execute(
        """INSERT INTO documentos
           (nome_arquivo, hash_arquivo, cnpj, cnpj_fundo, nome_classe,
            competencia, tipo_condominio, fdo_exclusivo, versao_layout,
            dt_processamento, caminho_xml)
           VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
        (
            nome_arquivo,
            hash_arquivo,
            cnpj,
            cnpj_fundo,
            nome_classe,
            competencia,
            tipo_cond,
            fdo_excl,
            versao_layout,
            datetime.now().isoformat(timespec="seconds"),
            caminho_xml,
        ),
    )
    doc_id = cursor.lastrowid
    assert doc_id is not None

    # Inserir campos em lote
    if dados.campos:
        conn.executemany(
            "INSERT INTO campos (documento_id, secao, nome_campo, valor) VALUES (?, ?, ?, ?)",
            [(doc_id, secao, nome, val) for secao, nome, val in dados.campos],
        )

    # Inserir cedentes em lote
    if dados.cedentes:
        conn.executemany(
            "INSERT INTO cedentes (documento_id, origem, nr_pf_pj, pr_participacao) VALUES (?, ?, ?, ?)",
            [
                (doc_id, c.get("origem", ""), c.get("nr_pf_pj", ""), c.get("pr_participacao", ""))
                for c in dados.cedentes
            ],
        )

    # Inserir devedores em lote
    if dados.devedores:
        conn.executemany(
            "INSERT INTO devedores (documento_id, nr_pf_pj, valor, pr_participacao) VALUES (?, ?, ?, ?)",
            [
                (
                    doc_id,
                    d.get("nr_pf_pj", ""),
                    d.get("valor", ""),
                    d.get("pr_participacao", ""),
                )
                for d in dados.devedores
            ],
        )

    return doc_id


# ============================================================================
# Backup rotativo (SQLite + CSV + Parquet)
# ============================================================================


def _pasta_backups(caminho_banco: str) -> str:
    """Pasta 'backups' ao lado do banco SQLite."""
    return os.path.join(os.path.dirname(os.path.abspath(caminho_banco)), "backups")


def _limpar_backups_antigos(pasta: str, manter: int) -> None:
    """Dentro de 'pasta', mantem so as 'manter' entradas mais recentes
    (arquivos ou subpastas), apagando as mais antigas pelo nome
    (os nomes usam timestamp AAAAMMDD_HHMMSS, entao ordenar por nome
    equivale a ordenar por data)."""
    if not os.path.isdir(pasta):
        return
    entradas = sorted(os.listdir(pasta))
    excedentes = entradas[: max(0, len(entradas) - manter)]
    for nome in excedentes:
        caminho = os.path.join(pasta, nome)
        try:
            if os.path.isdir(caminho):
                shutil.rmtree(caminho)
            else:
                os.remove(caminho)
        except OSError as e:
            print(f"[AVISO] Nao foi possivel remover backup antigo '{caminho}': {e}", flush=True)


def _backup_sqlite(caminho_banco: str, pasta_backups: str, timestamp: str) -> None:
    """Copia o arquivo .db atual para backups/sqlite/."""
    destino_pasta = os.path.join(pasta_backups, "sqlite")
    os.makedirs(destino_pasta, exist_ok=True)

    base = os.path.splitext(os.path.basename(caminho_banco))[0]
    destino = os.path.join(destino_pasta, f"{base}_{timestamp}.db")

    try:
        # Usa a API de backup do sqlite3 (seguro mesmo com WAL ativo)
        origem_conn = sqlite3.connect(caminho_banco)
        destino_conn = sqlite3.connect(destino)
        with destino_conn:
            origem_conn.backup(destino_conn)
        destino_conn.close()
        origem_conn.close()
        print(f"[BACKUP] SQLite -> {destino}", flush=True)
    except sqlite3.Error as e:
        print(f"[AVISO] Falha ao fazer backup SQLite: {e}", flush=True)
        return

    _limpar_backups_antigos(destino_pasta, MAX_BACKUPS_POR_FORMATO)


def _backup_csv(conn: sqlite3.Connection, pasta_backups: str, timestamp: str) -> None:
    """Exporta cada tabela para um CSV dentro de backups/csv/<timestamp>/."""
    destino_raiz = os.path.join(pasta_backups, "csv")
    destino_pasta = os.path.join(destino_raiz, timestamp)
    os.makedirs(destino_pasta, exist_ok=True)

    for tabela in TABELAS_BACKUP:
        caminho_csv = os.path.join(destino_pasta, f"{tabela}.csv")
        try:
            cursor = conn.execute(f"SELECT * FROM {tabela}")
            colunas = [d[0] for d in cursor.description]
            with open(caminho_csv, "w", encoding="utf-8-sig", newline="") as f:
                writer = csv_mod.writer(f)
                writer.writerow(colunas)
                writer.writerows(cursor)
        except (sqlite3.Error, OSError) as e:
            print(f"[AVISO] Falha ao exportar CSV de '{tabela}': {e}", flush=True)

    print(f"[BACKUP] CSV    -> {destino_pasta}", flush=True)
    _limpar_backups_antigos(destino_raiz, MAX_BACKUPS_POR_FORMATO)


def _backup_parquet(conn: sqlite3.Connection, pasta_backups: str, timestamp: str) -> None:
    """Exporta cada tabela para Parquet dentro de backups/parquet/<timestamp>/.
    Pulado silenciosamente (com aviso) se pyarrow nao estiver instalado."""
    if not _PYARROW_DISPONIVEL:
        print(
            "[AVISO] pacote 'pyarrow' nao instalado -- backup em Parquet "
            "pulado (pip install pyarrow para habilitar).",
            flush=True,
        )
        return

    destino_raiz = os.path.join(pasta_backups, "parquet")
    destino_pasta = os.path.join(destino_raiz, timestamp)
    os.makedirs(destino_pasta, exist_ok=True)

    for tabela in TABELAS_BACKUP:
        caminho_parquet = os.path.join(destino_pasta, f"{tabela}.parquet")
        try:
            cursor = conn.execute(f"SELECT * FROM {tabela}")
            colunas = [d[0] for d in cursor.description]
            linhas = cursor.fetchall()
            # Monta colunas no formato esperado pelo pyarrow (lista de listas por coluna)
            dados_colunas = {
                nome: [linha[i] for linha in linhas] for i, nome in enumerate(colunas)
            }
            tabela_arrow = pa.table(dados_colunas)
            pq.write_table(tabela_arrow, caminho_parquet)
        except Exception as e:
            print(f"[AVISO] Falha ao exportar Parquet de '{tabela}': {e}", flush=True)

    print(f"[BACKUP] Parquet -> {destino_pasta}", flush=True)
    _limpar_backups_antigos(destino_raiz, MAX_BACKUPS_POR_FORMATO)


def fazer_backup(caminho_banco: str) -> None:
    """Executa o backup rotativo (SQLite + CSV + Parquet) do estado atual
    do banco, antes de processar novos XMLs. Nao falha a execucao principal
    se algum formato de backup der erro."""
    if not os.path.isfile(caminho_banco):
        # Banco ainda nao existe (primeira execucao) -- nada para backupear.
        return

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    pasta_backups = _pasta_backups(caminho_banco)

    print(f"[INFO] Fazendo backup do estado atual ({timestamp})...", flush=True)

    _backup_sqlite(caminho_banco, pasta_backups, timestamp)

    try:
        conn = sqlite3.connect(f"file:{caminho_banco}?mode=ro", uri=True)
    except sqlite3.Error as e:
        print(f"[AVISO] Nao foi possivel abrir banco para backup CSV/Parquet: {e}", flush=True)
        return

    try:
        _backup_csv(conn, pasta_backups, timestamp)
        _backup_parquet(conn, pasta_backups, timestamp)
    finally:
        conn.close()

    print(flush=True)


# ============================================================================
# Processamento principal
# ============================================================================


def listar_xmls(pasta: str) -> list[str]:
    """Lista todos os .xml na pasta (nao recursivo)."""
    if not os.path.isdir(pasta):
        raise FileNotFoundError(f"Pasta de XMLs nao encontrada: {pasta}")
    return sorted(
        os.path.join(pasta, f)
        for f in os.listdir(pasta)
        if f.lower().endswith(".xml") and not f.startswith(".")
    )


def registrar_erro(pasta_log: str, nome_arquivo: str, erro: str) -> None:
    """Registra erro de processamento no log."""
    caminho_log = os.path.join(pasta_log, NOME_LOG_ERROS)
    try:
        with open(caminho_log, "a", encoding="utf-8") as f:
            f.write(f"[{datetime.now().isoformat(timespec='seconds')}] {nome_arquivo}: {erro}\n")
    except OSError:
        pass


def processar(pasta_xml: str, caminho_banco: str) -> None:
    """Fluxo principal: le todos os XMLs e insere no banco."""
    print("=" * 70, flush=True)
    print("  FIDC XML -> SQLite Parser", flush=True)
    print("=" * 70, flush=True)
    print(f"  Pasta XML : {pasta_xml}", flush=True)
    print(f"  Banco     : {caminho_banco}", flush=True)
    print(flush=True)

    # Listar XMLs
    try:
        arquivos = listar_xmls(pasta_xml)
    except FileNotFoundError as e:
        print(f"[ERRO] {e}", flush=True)
        sys.exit(1)

    total = len(arquivos)
    if total == 0:
        print("[INFO] Nenhum arquivo XML encontrado na pasta.", flush=True)
        return

    print(f"[INFO] {total} arquivo(s) XML encontrado(s).", flush=True)

    # Backup rotativo do estado atual (antes de qualquer alteracao no banco)
    try:
        fazer_backup(caminho_banco)
    except Exception as e:
        # Backup nunca deve impedir o processamento principal.
        print(f"[AVISO] Falha inesperada no backup, prosseguindo sem ele: {e}", flush=True)

    # Abrir/criar banco
    try:
        conn = criar_banco(caminho_banco)
    except sqlite3.Error as e:
        print(f"[ERRO] Falha ao criar/abrir banco: {e}", flush=True)
        sys.exit(1)

    novos = 0
    existentes = 0
    erros = 0

    # Filtro rapido: nomes ja registrados no banco. Evita calcular o
    # hash SHA-256 (leitura do arquivo inteiro) para arquivos que,
    # pelo nome deterministico do scraper, certamente ja foram
    # processados -- essencial em pastas com milhares de XMLs.
    nomes_conhecidos = carregar_nomes_ja_processados(conn)
    print(f"[INFO] {len(nomes_conhecidos)} arquivo(s) ja registrados no banco.", flush=True)

    for i, caminho in enumerate(arquivos, 1):
        nome = os.path.basename(caminho)
        progresso = f"[{i}/{total}]"

        if nome in nomes_conhecidos:
            existentes += 1
            continue

        # Calcular hash
        try:
            hash_arq = calcular_hash(caminho)
        except OSError as e:
            print(f"  {progresso} {nome} -> ERRO ao ler arquivo: {e}", flush=True)
            registrar_erro(pasta_xml, nome, f"Erro ao ler: {e}")
            erros += 1
            continue

        # Verificar duplicata por conteudo (protege contra arquivo
        # renomeado ou reprocessado com outro nome, mesmo conteudo)
        if hash_ja_processado(conn, hash_arq):
            existentes += 1
            continue

        # Parse XML
        try:
            dados = parse_xml(caminho)
        except ET.ParseError as e:
            print(f"  {progresso} {nome} -> ERRO de parse XML: {e}", flush=True)
            registrar_erro(pasta_xml, nome, f"XML invalido: {e}")
            erros += 1
            continue
        except Exception as e:
            print(f"  {progresso} {nome} -> ERRO inesperado: {e}", flush=True)
            registrar_erro(pasta_xml, nome, f"Erro inesperado: {traceback.format_exc()}")
            erros += 1
            continue

        # Extrair CNPJ do nome (fallback: do cabecalho)
        cnpj = extrair_cnpj_do_nome(nome)
        if not cnpj:
            cnpj = so_digitos(dados.cabecalho.get("NR_CNPJ_ADM", ""))

        # Inserir no banco
        try:
            doc_id = inserir_documento(conn, nome, hash_arq, cnpj, dados, caminho)
            conn.commit()
            qtd_campos = len(dados.campos)
            qtd_cedentes = len(dados.cedentes)
            qtd_devedores = len(dados.devedores)
            print(
                f"  {progresso} {nome} -> OK "
                f"(id={doc_id}, {qtd_campos} campos, "
                f"{qtd_cedentes} cedentes, {qtd_devedores} devedores)",
                flush=True,
            )
            novos += 1
        except sqlite3.IntegrityError:
            # Hash duplicado (race condition improvavel, mas seguro)
            conn.rollback()
            existentes += 1
        except sqlite3.Error as e:
            conn.rollback()
            print(f"  {progresso} {nome} -> ERRO ao inserir no banco: {e}", flush=True)
            registrar_erro(pasta_xml, nome, f"Erro SQLite: {e}")
            erros += 1

    conn.close()

    # Resumo
    print(flush=True)
    print("=" * 70, flush=True)
    print(f"  Concluido!", flush=True)
    print(f"  Novos     : {novos}", flush=True)
    print(f"  Existentes: {existentes} (ignorados)", flush=True)
    print(f"  Erros     : {erros}", flush=True)
    print("=" * 70, flush=True)


# ============================================================================
# CLI
# ============================================================================


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Parser de XMLs de Informe Mensal FIDC para SQLite.",
    )
    parser.add_argument(
        "--pasta-xml",
        default=PASTA_XML_PADRAO,
        help=f"Pasta contendo os XMLs (padrao: {PASTA_XML_PADRAO})",
    )
    parser.add_argument(
        "--banco",
        default=BANCO_PADRAO,
        help=f"Caminho do banco SQLite (padrao: {BANCO_PADRAO})",
    )
    args = parser.parse_args()
    processar(args.pasta_xml, args.banco)


if __name__ == "__main__":
    main()