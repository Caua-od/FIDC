"""
Baixa as carteiras mensais de fundos (CDA) da CVM referentes aos últimos N meses.

Fonte: https://dados.cvm.gov.br/dataset/fi-doc-cda
Padrão de arquivo: cda_fi_AAAAMM.zip

Para cada mês do período:
  1. Baixa o ZIP (se ainda não foi baixado antes).
  2. Extrai o conteúdo numa subpasta própria do mês.
  3. Apaga o ZIP após extrair com sucesso.

Idempotente: rodar de novo não duplica nem rebaixa meses já extraídos.
"""

import logging
import shutil
import sys
from datetime import date
from pathlib import Path

import requests

BASE_URL = "https://dados.cvm.gov.br/dados/FI/DOC/CDA/DADOS/cda_fi_{periodo}.zip"
DESTINO = Path(r"C:\Users\olive\Downloads\Franklin project\Data\Scrapping\CDA\dados")
MESES_PARA_BAIXAR = 7
TIMEOUT_SEGUNDOS = 60

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%H:%M:%S",
)
log = logging.getLogger(__name__)


def ultimos_meses(qtd: int) -> list[str]:
    """Retorna os últimos `qtd` períodos no formato AAAAMM, do mais antigo ao mais recente."""
    periodos = []
    ano, mes = date.today().year, date.today().month
    for _ in range(qtd):
        periodos.append(f"{ano}{mes:02d}")
        mes -= 1
        if mes == 0:
            mes, ano = 12, ano - 1
    return list(reversed(periodos))


def ja_extraido(pasta_periodo: Path) -> bool:
    """Considera extraído se a pasta existe e tem pelo menos um arquivo dentro."""
    return pasta_periodo.is_dir() and any(pasta_periodo.iterdir())


def baixar_zip(periodo: str, destino_zip: Path) -> bool:
    """Baixa o ZIP do período. Retorna False se o mês ainda não foi publicado (404)."""
    url = BASE_URL.format(periodo=periodo)
    try:
        with requests.get(url, stream=True, timeout=TIMEOUT_SEGUNDOS) as resp:
            if resp.status_code == 404:
                log.info("%s: ainda não publicado pela CVM — pulando.", periodo)
                return False
            resp.raise_for_status()
            with open(destino_zip, "wb") as f:
                shutil.copyfileobj(resp.raw, f)
        return True
    except requests.RequestException as e:
        log.error("%s: falha no download (%s)", periodo, e)
        return False


def processar_periodo(periodo: str) -> None:
    pasta_periodo = DESTINO / periodo
    zip_path = DESTINO / f"cda_fi_{periodo}.zip"

    if ja_extraido(pasta_periodo):
        log.info("%s: já extraído — pulando.", periodo)
        return

    log.info("%s: baixando...", periodo)
    if not baixar_zip(periodo, zip_path):
        return

    try:
        pasta_periodo.mkdir(parents=True, exist_ok=True)
        shutil.unpack_archive(str(zip_path), str(pasta_periodo))
        log.info("%s: extraído em %s", periodo, pasta_periodo)
    except (shutil.ReadError, OSError) as e:
        log.error("%s: falha ao extrair (%s) — arquivo zip mantido para inspeção.", periodo, e)
        return
    finally:
        if zip_path.exists() and ja_extraido(pasta_periodo):
            zip_path.unlink()


def main() -> None:
    DESTINO.mkdir(parents=True, exist_ok=True)
    log.info("Destino: %s", DESTINO)

    for periodo in ultimos_meses(MESES_PARA_BAIXAR):
        processar_periodo(periodo)

    log.info("Concluído.")


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        sys.exit(1)