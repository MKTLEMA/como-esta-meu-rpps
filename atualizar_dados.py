"""Atualiza os dados do app a partir das planilhas oficiais do ISP-RPPS (Ministério da Previdência).

Uso no projeto (gera app/dados/rpps.json):
    pip install openpyxl
    python scripts/atualizar_dados.py

Uso no repositório publicado (troca os dados dentro do index.html, roda todo dia no GitHub Actions):
    python atualizar_dados.py --html index.html

O script acha na página oficial do ISP o resultado final mais recente e o do ano anterior (o Manual de
Certificação usa o ISP do penúltimo exercício para definir o porte), baixa para uma pasta temporária e só
regrava quando os dados mudam. Também vigia a página do Manual de Certificação: se sair versão maior que a
usada em app/regras.js, escreve um aviso (::warning:: no Actions e o arquivo ALERTA-NORMA.txt).
"""
import json
import re
import sys
import tempfile
import unicodedata
import urllib.request
from datetime import date
from pathlib import Path

import openpyxl

PAGINA_ISP = "https://www.gov.br/previdencia/pt-br/assuntos/rpps/indice-de-situacao-previdenciaria"
PAGINA_MANUAL = "https://www.gov.br/previdencia/pt-br/assuntos/rpps/Requisitos-para-Dirigentes-e-Conselheiros-de-RPPS/manual-da-certificacao-profissional"
MANUAL_NO_APP = (2, 0)  # versão do Manual usada em regras.js (REGRAS.manualVersao)
BASE = "https://www.gov.br"
DESTINO = Path(__file__).resolve().parent.parent / "app" / "dados" / "rpps.json"
UA = {"User-Agent": "Mozilla/5.0 (LEMA Edu - atualizador de dados RPPS)"}
BLOCO = re.compile(r'(<script id="dados" type="application/json">)(.*?)(</script>)', re.S)
FECHA = "<" + "\\" + "/"  # "<\/", forma segura de escrever "</" dentro de <script>


def baixar(url, destino=None):
    req = urllib.request.Request(url, headers=UA)
    with urllib.request.urlopen(req, timeout=120) as r:
        dados = r.read()
    if destino:
        Path(destino).write_bytes(dados)
    return dados


def sem_acento(s):
    return unicodedata.normalize("NFKD", str(s)).encode("ascii", "ignore").decode().upper().strip()


def achar_planilhas_finais():
    """Devolve {ano: url} das planilhas de resultado final publicadas na página do ISP."""
    html = baixar(PAGINA_ISP).decode("utf-8", "ignore")
    links = re.findall(r'href="([^"]+\.xlsx)"', html, flags=re.I)
    achados = {}
    for link in links:
        nome = link.lower()
        if "preliminar" in nome or "previo" in nome:
            continue
        if "final" not in nome and "consolidado" not in nome:
            continue
        ano = re.search(r"(20\d\d)", nome)
        if not ano:
            continue
        url = link if link.startswith("http") else BASE + link
        achados.setdefault(int(ano.group(1)), url)
    return achados


def ler_planilha(caminho):
    wb = openpyxl.load_workbook(caminho, read_only=True, data_only=True)
    ws = wb.worksheets[0]
    linhas = list(ws.iter_rows(values_only=True))
    ini = next(i for i, r in enumerate(linhas) if r and sem_acento(r[0]) == "ENTE")
    cab = [sem_acento(c or "") for c in linhas[ini]]

    def col(*partes):
        for i, c in enumerate(cab):
            if all(p in c for p in partes):
                return i
        raise KeyError(partes)

    i_ente, i_uf = col("ENTE"), col("UF")
    i_grupo, i_sub = col("GRUPO"), col("SUBGRUPO")
    try:  # a planilha do ano anterior pode ter outro formato; só o ano atual precisa do ISP
        i_isp = col("INDICADOR DE SITUACAO PREVIDENCIARIA")
    except KeyError:
        i_isp = None
    i_perfil = col("PERFIL ATUARIAL")
    out = {}
    for r in linhas[ini + 1:]:
        if not r[i_ente] or not r[i_uf]:
            continue
        nome = re.sub(r"\s+-\s+[A-Z]{2}$", "", str(r[i_ente]).strip())
        out[(r[i_uf], nome)] = {
            "grupo": sem_acento(r[i_grupo] or ""),
            "subgrupo": sem_acento(r[i_sub] or ""),
            "isp": str(r[i_isp]).strip() if i_isp is not None and r[i_isp] else "",
            "perfil": str(r[i_perfil]).strip() if r[i_perfil] else "",
        }
    return out


def versao_do_manual():
    """Maior 'versão X.Y' listada na página oficial do Manual de Certificação."""
    try:
        html = baixar(PAGINA_MANUAL).decode("utf-8", "ignore")
    except Exception as e:  # página fora do ar não deve derrubar a atualização dos dados
        print("::warning::Não consegui abrir a página do Manual:", e)
        return None
    achadas = [(int(a), int(b)) for a, b in re.findall(r"vers[ãa]o\s*(\d+)\.(\d+)", html, flags=re.I)]
    return max(achadas) if achadas else None


def montar(planilhas, dados, ano_atual, ano_anterior):
    entes = []
    for (uf, nome), d in sorted(dados[ano_atual].items()):
        ant = dados[ano_anterior].get((uf, nome), {})
        entes.append({
            "n": nome, "uf": uf,
            "g": d["grupo"], "sg": d["subgrupo"], "isp": d["isp"], "pa": d["perfil"],
            "gAnt": ant.get("grupo", ""),
        })
    return {
        "meta": {
            "fonte": "Ministério da Previdência Social, SPREV: ISP-RPPS, resultado final",
            "paginaOficial": PAGINA_ISP,
            "anoISP": ano_atual, "anoISPAnterior": ano_anterior,
            "urlPlanilha": planilhas[ano_atual],
            "geradoEm": date.today().isoformat(),
        },
        "entes": entes,
    }


def main():
    html_alvo = Path(sys.argv[sys.argv.index("--html") + 1]) if "--html" in sys.argv else None
    planilhas = achar_planilhas_finais()
    anos = sorted(planilhas)
    if len(anos) < 2:
        sys.exit(f"Não achei duas planilhas finais na página do ISP: {planilhas}")
    ano_atual, ano_anterior = anos[-1], anos[-2]
    tmp = Path(tempfile.mkdtemp(prefix="isp_"))
    dados = {}
    for ano in (ano_atual, ano_anterior):
        arq = tmp / f"isp{ano}.xlsx"
        print("Baixando", ano, planilhas[ano])
        baixar(planilhas[ano], arq)
        dados[ano] = ler_planilha(arq)
    novo = montar(planilhas, dados, ano_atual, ano_anterior)

    # Só regrava quando os dados mudam de verdade (assim o repositório não recebe commit à toa).
    if html_alvo:
        achou = BLOCO.search(html_alvo.read_text(encoding="utf-8"))
        antigo = json.loads(achou.group(2).replace(FECHA, "</")) if achou else {}
    else:
        antigo = json.loads(DESTINO.read_text(encoding="utf-8")) if DESTINO.exists() else {}
    meta_antiga = antigo.get("meta", {})
    if (antigo.get("entes") == novo["entes"] and meta_antiga.get("anoISP") == ano_atual
            and meta_antiga.get("urlPlanilha") == planilhas[ano_atual]):
        print("Sem mudanças nos dados do ISP.")
    else:
        texto = json.dumps(novo, ensure_ascii=False, separators=(",", ":"))
        if html_alvo:
            html = html_alvo.read_text(encoding="utf-8")
            seguro = texto.replace("</", FECHA)
            html = BLOCO.sub(lambda m: m.group(1) + seguro + m.group(3), html, count=1)
            html_alvo.write_text(html, encoding="utf-8")
            print(f"OK: {len(novo['entes'])} entes gravados em {html_alvo}")
        else:
            DESTINO.parent.mkdir(parents=True, exist_ok=True)
            DESTINO.write_text(texto, encoding="utf-8")
            print(f"OK: {len(novo['entes'])} entes gravados em {DESTINO}")

    v = versao_do_manual()
    if v and v > MANUAL_NO_APP:
        msg = f"Saiu o Manual de Certificação versão {v[0]}.{v[1]}. Revise regras.js (hoje usa {MANUAL_NO_APP[0]}.{MANUAL_NO_APP[1]})."
        print("::warning::" + msg)
        Path("ALERTA-NORMA.txt").write_text(msg + "\n", encoding="utf-8")
        sys.exit(3)  # faz o GitHub avisar por e-mail que o job falhou


if __name__ == "__main__":
    main()
