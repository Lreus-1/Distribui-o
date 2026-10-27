# -*- coding: utf-8 -*-
"""
Servidor do painel de Distribuição de Cerveja (LUVI BI · Nuernberg).

Entrega o index.html e, para cada fonte, sempre o arquivo MAIS RECENTE da pasta:
  - Vendas   : arquivo com "03.02.37" no nome   (pasta da Crítica)
  - Clientes : arquivo com "01.20.11" no nome   (pasta da Crítica)
  - Cesta    : arquivo que começa com "Cerveja" (subpasta Cestas)

O cruzamento é feito aqui no servidor (uma vez a cada mudança nos arquivos) e o
navegador recebe só o resultado, compactado. Só usa a biblioteca padrão do
Python (nada para instalar).
"""
import json
import os
import re
import socket
import sys
import time
import unicodedata
from datetime import datetime
from http.server import ThreadingHTTPServer, BaseHTTPRequestHandler
from pathlib import Path
from urllib.parse import urlparse, unquote

# ============================ CONFIGURAÇÃO ============================
PORTA = 8098   # 8090, 8096 e 8097 já usadas

# Testadas em ordem; vale a primeira que existir (a letra F: pode não existir
# quando o servidor roda por tarefa agendada, por isso o caminho de rede também).
PASTAS_RELATORIOS = [
    r"F:\PUBLICO\33 - FECHAMENTO DIÁRIO\Crítica",
    r"\\192.168.4.4\Fileserver\PUBLICO\33 - FECHAMENTO DIÁRIO\Crítica",
]
if os.environ.get("PASTA_CRITICA"):          # opcional: sobrescreve a pasta sem editar o arquivo
    PASTAS_RELATORIOS.insert(0, os.environ["PASTA_CRITICA"])
SUBPASTA_CESTAS = "Cestas"

CODIGO_VENDAS = "030237"     # 03.02.37 (só os dígitos)
CODIGO_CLIENTES = "012011"   # 01.20.11
NOME_CESTA = "cerveja"       # início do nome do arquivo da cesta
SUBPASTA_SCORE5 = "Base Score 5"   # pasta da Base Score 5 (opcional): vale o arquivo mais recente dela, qualquer nome
NOME_SCORE5 = "basescore"    # arquivo da Base Score 5 (o nome começa com "Base Score 5")
NOME_BONI = "boni"           # planilha de bonificação (ex.: "00. Planificador Boni IMC MDLZ"): o nome contém "boni"

EXTENSOES = {".csv", ".txt", ".xlsx", ".xls", ".xlsm", ".xlsb"}
# =====================================================================

BASE = Path(__file__).resolve().parent
TIPOS = {
    ".html": "text/html; charset=utf-8",
    ".js": "application/javascript; charset=utf-8",
    ".csv": "text/csv",
    ".txt": "text/plain",
    ".xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    ".xls": "application/vnd.ms-excel",
    ".xlsm": "application/vnd.ms-excel.sheet.macroEnabled.12",
    ".xlsb": "application/vnd.ms-excel.sheet.binary.macroEnabled.12",
    ".png": "image/png",
    ".ico": "image/x-icon",
}


def sem_acento(txt):
    return "".join(c for c in unicodedata.normalize("NFD", txt) if unicodedata.category(c) != "Mn")


def pasta_relatorios():
    for p in PASTAS_RELATORIOS:
        if os.path.isdir(p):
            return Path(p)
    return None


def mais_recente(pasta, teste):
    """Arquivo mais recente (data de modificação) da pasta que passa no teste do nome."""
    melhor = None
    try:
        with os.scandir(pasta) as it:
            for e in it:
                nome = e.name
                if nome.startswith("~$") or not e.is_file():
                    continue
                if os.path.splitext(nome)[1].lower() not in EXTENSOES:
                    continue
                if not teste(nome):
                    continue
                st = e.stat()
                if melhor is None or st.st_mtime > melhor["mtime"]:
                    melhor = {"caminho": e.path, "nome": nome, "mtime": st.st_mtime, "tamanho": st.st_size}
    except OSError:
        return None
    return melhor


_cache = {"t": 0, "dados": None}


def localizar():
    """Procura os três arquivos. Guarda o resultado por 10 s para não varrer a rede a cada pedido."""
    if _cache["dados"] and time.time() - _cache["t"] < 10:
        return _cache["dados"]
    base = pasta_relatorios()
    if base is None:
        dados = {"ok": False, "erro": "Pasta da Crítica não encontrada: " + " | ".join(PASTAS_RELATORIOS), "arquivos": {}}
    else:
        digitos = lambda n: re.sub(r"\D", "", n)
        achados = {
            "vendas": mais_recente(base, lambda n: CODIGO_VENDAS in digitos(n)),
            "clientes": mais_recente(base, lambda n: CODIGO_CLIENTES in digitos(n)),
            "cesta": mais_recente(base / SUBPASTA_CESTAS,
                                  lambda n: sem_acento(os.path.splitext(n)[0]).lower().startswith(NOME_CESTA)),
        }
        # Opcionais, na pasta "Base Score 5" (ou solta na Crítica / em Cestas):
        #   - Base Score 5  -> filtro Score 5     - Planificador Boni -> aviso de bonificação disponível
        # Se não existirem, o recurso correspondente simplesmente não aparece.
        norm = lambda n: re.sub(r"[^a-z0-9]", "", sem_acento(os.path.splitext(n)[0]).lower())
        eh_boni = lambda n: NOME_BONI in norm(n)
        eh_score = lambda n: norm(n).startswith(NOME_SCORE5) and not eh_boni(n)
        pastas = [base / SUBPASTA_SCORE5, base, base / SUBPASTA_CESTAS]
        melhor = lambda cand: max(cand, key=lambda c: c["mtime"]) if cand else None
        score5 = melhor([c for c in (mais_recente(pt, eh_score) for pt in pastas) if c])
        boni = melhor([c for c in (mais_recente(pt, eh_boni) for pt in pastas) if c])
        faltando = [k for k, v in achados.items() if v is None]
        dados = {
            "ok": not faltando,
            "pasta": str(base),
            "erro": ("Não encontrei: " + ", ".join(faltando)) if faltando else "",
            "arquivos": {k: v for k, v in achados.items() if v},
        }
        pasta5 = base / SUBPASTA_SCORE5
        try:
            vistos = sorted(n for n in os.listdir(pasta5) if not n.startswith("~$"))
            existe5 = True
        except OSError:
            vistos, existe5 = [], False
        dados["base_score5"] = {"pasta": str(pasta5), "existe": existe5, "arquivos": vistos}
        if score5:
            dados["arquivos"]["score5"] = score5
        if boni:
            dados["arquivos"]["boni"] = boni
    _cache.update(t=time.time(), dados=dados)
    return dados


def ler_arquivo(caminho, tentativas=5):
    """Lê o arquivo; se estiver sendo gravado/travado, tenta de novo por alguns segundos."""
    for i in range(tentativas):
        try:
            with open(caminho, "rb") as f:
                return f.read()
        except PermissionError:
            time.sleep(1 + i)
    with open(caminho, "rb") as f:
        return f.read()


def iso(mtime):
    return datetime.fromtimestamp(mtime).astimezone().isoformat(timespec="seconds")



# =====================================================================
#   CRUZAMENTO (mesma lógica do painel: DISTINCT cliente + SKU da cesta)
# =====================================================================
import csv
import gzip
import io
import threading
import zipfile
import xml.etree.ElementTree as ET

STATUS_IGNORAR = {"C", "D"}   # linhas de venda canceladas (C) ou devolvidas (D)
DIAS = ["DOM", "SEG", "TER", "QUA", "QUI", "SEX", "SAB"]


def norm_h(v):
    return re.sub(r"[^a-z0-9]", "", sem_acento(str(v if v is not None else "")).lower())


def clean(v):
    return re.sub(r"\s+", " ", str(v if v is not None else "")).strip()


def norm_code(v):
    # "22.983" é milhar (CSV), não decimal: só tira tudo que não for dígito.
    return re.sub(r"\D", "", str(v if v is not None else "")).lstrip("0")


def parse_num(v):
    t = str(v if v is not None else "").strip()
    if not t:
        return 0.0
    if "," in t:
        t = t.replace(".", "").replace(",", ".")
    t = re.sub(r"[^0-9.\-]", "", t)
    try:
        return float(t)
    except ValueError:
        m = re.match(r"-?\d+(\.\d+)?", t)
        return float(m.group(0)) if m else 0.0


def parse_date(v):
    t = str(v if v is not None else "").strip()
    if not t:
        return None
    m = re.match(r"^(\d{1,2})[/.\-](\d{1,2})[/.\-](\d{2,4})", t)
    if m:
        y = int(m.group(3))
        y = y + 2000 if y < 100 else y
        return "%04d-%02d-%02d" % (y, int(m.group(2)), int(m.group(1)))
    m = re.match(r"^(\d{4})-(\d{1,2})-(\d{1,2})", t)
    if m:
        return "%s-%02d-%02d" % (m.group(1), int(m.group(2)), int(m.group(3)))
    if re.fullmatch(r"\d{5}(\.\d+)?", t):
        n = float(t)
        if 20000 < n < 80000:
            return datetime.utcfromtimestamp(round((n - 25569) * 86400)).strftime("%Y-%m-%d")
    return None


def parse_dias(v):
    toks = [x for x in re.split(r"[^A-Z]+", sem_acento(str(v or "")).upper()) if x]
    mask = 0
    for t in toks:
        if t[:3] in DIAS:
            mask |= 1 << DIAS.index(t[:3])
    return mask


MARCAS = [
    ("BEATS", "BEATS"), ("SKOL BEATS", "BEATS"), ("SKOL", "SKOL"), ("MALZBIER BRAHMA", "BRAHMA"), ("BRAHMA", "BRAHMA"),
    ("ANTARCTICA", "ANTARCTICA"), ("ORIGINAL", "ORIGINAL"), ("BUDWEISER", "BUDWEISER"), ("STELLA ARTOIS", "STELLA ARTOIS"),
    ("CORONITA", "CORONA"), ("CORONA", "CORONA"), ("PATAGONIA", "PATAGONIA"), ("COLORADO", "COLORADO"), ("MICHELOB", "MICHELOB"),
    ("SPATEN", "SPATEN"), ("BOHEMIA", "BOHEMIA"), ("CARACU", "CARACU"), ("SERRAMALTE", "SERRAMALTE"), ("MIKES", "MIKES"),
    ("XEQUE MATE", "XEQUE MATE"), ("BRUTAL FRUIT", "BRUTAL FRUIT"), ("FLYING FISH", "FLYING FISH"), ("GOOSE ISLAND", "GOOSE ISLAND"),
    ("BECKS", "BECKS"), ("HOEGAARDEN", "HOEGAARDEN"), ("LEFFE", "LEFFE"), ("WALS", "WALS"),
]


def infer_marca(desc):
    d = " " + re.sub(r"^CHOPP\s+(CLARO\s+)?", "", desc.upper()) + " "
    best, pos = None, 10 ** 9
    for k, label in MARCAS:
        i = d.find(" " + k + " ")
        if 0 <= i < pos:
            pos, best = i, label
    return best or ((desc.split() or ["OUTROS"])[0]).upper()


def infer_emb(desc):
    d = " " + desc.upper() + " "
    if re.search(r" (KEG|BARRIL) ", d): return "Barril"
    if re.search(r" (LONG NECK|LN) ", d): return "Long Neck"
    if re.search(r" (ONE WAY|OW) ", d): return "One Way"
    if re.search(r" (LATA|LT|LAT) ", d): return "Lata"
    if re.search(r" (GFA|GARRAFA|VD) ", d) or re.search(r"(600 ?ML|1L|990ML)", d): return "Retornável"
    return "Outros"


def infer_vol(desc):
    m = re.search(r"(\d+(?:,\d+)?)\s?(ML|L)\b", desc.upper())
    return (m.group(1) + ("ml" if m.group(2) == "ML" else "L")) if m else ""


COLS = {
    "cesta": {
        "sku": (["cod", "codigo", "codsku", "sku", "codproduto", "codigoproduto", "codigosku", "material", "codmaterial"], True, "Código SKU"),
        "desc": (["descricaodoproduto", "descricaoproduto", "descricao", "descproduto", "produto", "nomeproduto", "nome"], True, "Produto"),
        "cesta": (["cesta", "categoria", "grupo"], False, "Cesta"),
        "marca": (["marca"], False, "Marca"),
        "emb": (["embalagem", "tipoembalagem"], False, "Embalagem"),
        "vol": (["volume"], False, "Volume"),
    },
    "vendas": {
        "cli": (["cliente", "codcliente", "codigocliente", "codcli", "codpdv", "pdv"], True, "Cliente"),
        "sku": (["codproduto", "codigoproduto", "codsku", "sku", "codprod", "produto"], True, "SKU"),
        "qtd": (["quantidade", "qtd", "qtde", "qtdvendida", "quantidadevendida"], False, "Quantidade"),
        "valor": (["totalvenda", "valortotal", "vlrtotal", "total", "valor"], False, "Valor"),
        "status": (["status", "statusvenda", "statusnota", "situacao"], False, "Status (C/D ignorados)"),
        "data": (["dtoperacao", "dataoperacao", "datavenda", "data", "dtvenda", "dataemissao", "dtemissao", "emissao", "datafaturamento",
                  "dtfaturamento", "datanf", "datapedido", "dtpedido", "dtentrega", "dataentrega"], False, "Data da venda"),
        "desc": (["descproduto", "descricaoproduto", "descricao"], False, "Descrição"),
    },
    "boni": {
        # a ordem importa: "flag" reserva a coluna "Bonificação" antes de "boni" tentar casar por semelhança
        "pdv": (["unbpdv", "pdvunb", "codpdv", "pdv"], True, "UNB_PDV (cliente)"),
        "flag": (["bonificacao", "bonificar", "statusbonificacao"], False, "Bonificação (Sim/Não)"),
        "boni": (["boni", "produtoboni", "produtobonificado"], False, "BONI (produto a bonificar)"),
        "prod": (["nomeproduto", "nomedoproduto", "descricaoproduto", "descproduto", "descricao"], False, "Nome do produto"),
    },
    "score5": {
        "cli": (["cliente", "codcliente", "codigocliente", "codcli", "codpdv", "pdv"], True, "Cliente"),
        "base": (["base", "basescore", "score5", "score", "classificacao", "tipo"], True, "Base (High End / Core)"),
    },
    "clientes": {
        "cod": (["codigocliente", "codcliente", "cliente", "codigo", "codpdv"], True, "Código do cliente"),
        "fantasia": (["nomefantasia", "fantasia", "nomecliente", "nome"], False, "Nome fantasia"),
        "razao": (["razaosocial", "razao"], False, "Razão social"),
        "setor": (["codsetor", "setor", "codigosetor", "rota", "codrota"], True, "Setor / rota"),
        "setorNome": (["descricaosetor", "nomesetor", "descsetor", "nomerota"], False, "Nome do setor"),
        "dia": (["frequenciavisita", "diavisita", "diadevisita", "diasvisita", "diasdevisita", "dia"], False, "Dia de visita"),
        "cidade": (["cidade", "municipio"], False, "Cidade"),
        "bairro": (["bairro"], False, "Bairro"),
        "endereco": (["endereco", "logradouro"], False, "Endereço"),
        "status": (["statusdocliente", "statuscliente", "status", "situacao"], False, "Status"),
        "segmento": (["segmento"], False, "Segmento"),
        "prox": (["proximavisita"], False, "Próxima visita"),
    },
}


def find_col(H, cands, used):
    """Devolve (índice da coluna, posição do nome na lista de candidatos) ou (-1, 0)."""
    for r, c in enumerate(cands):
        for i, h in enumerate(H):
            if h == c and i not in used:
                return i, r
    for r, c in enumerate(cands):
        if len(c) < 4:
            continue
        for i, h in enumerate(H):
            if c in h and i not in used:
                return i, r
    return -1, 0


def map_columns(sheets, spec):
    best = None
    for si, rows in enumerate(sheets):
        for r in range(min(20, len(rows))):
            H = [norm_h(x) for x in (rows[r] or [])]
            used, idx, score, missing = set(), {}, 0, []
            for key, (cands, req, label) in spec.items():
                i, rank = find_col(H, cands, used)
                idx[key] = i
                if i >= 0:
                    used.add(i)
                    score += (10 if req else 1) + (len(cands) - rank) / 100.0
                elif req:
                    missing.append(label)
            if best is None or score > best["score"]:
                best = {"score": score, "sheet": si, "hr": r, "idx": idx, "missing": missing, "headers": rows[r] or []}
    return best


def describe_cols(m, spec):
    out = []
    for k, (_, _, label) in spec.items():
        i = m["idx"][k]
        out.append("  " + label + ": " + (("“" + clean(m["headers"][i]) + "”") if i >= 0 else "— não encontrada"))
    return "\n".join(out)


def cel(row, i):
    return row[i] if 0 <= i < len(row) else ""


def r2(x):
    v = round(x + 0.0, 2)
    return int(v) if v == int(v) else v


def natural(s):
    return [int(t) if t.isdigit() else t.lower() for t in re.split(r"(\d+)", s)]


def build_payload(entrada, nomes):
    log, avisos, sets = [], [], {}
    for k in ("cesta", "vendas", "clientes"):
        m = map_columns(entrada[k], COLS[k])
        if not m or m["missing"]:
            raise ValueError("Arquivo de %s (%s): não encontrei a(s) coluna(s) %s." % (k, nomes.get(k, ""), ", ".join(m["missing"]) if m else "esperadas"))
        sets[k] = m
        log.append({"k": k, "txt": describe_cols(m, COLS[k])})

    # SCORE 5 (opcional)
    bases, score_map = [], {}
    if entrada.get("score5"):
        m5 = map_columns(entrada["score5"], COLS["score5"])
        if not m5 or m5["missing"]:
            avisos.append("Base Score 5 (%s): não encontrei a(s) coluna(s) %s. O filtro Score 5 ficou desativado." % (nomes.get("score5", ""), ", ".join(m5["missing"]) if m5 else "esperadas"))
        else:
            log.append({"k": "score5", "txt": describe_cols(m5, COLS["score5"])})
            tmp = {}
            for row in entrada["score5"][m5["sheet"]][m5["hr"] + 1:]:
                code = norm_code(cel(row, m5["idx"]["cli"])); base = clean(cel(row, m5["idx"]["base"])).upper()
                if code and base:
                    tmp[code] = base
            bases = sorted(set(tmp.values()))
            score_map = {c: bases.index(b) + 1 for c, b in tmp.items()}

    # BONIFICAÇÃO (opcional): UNB_PDV = "636002_12345" -> vale só o número depois do "_".
    # Tem bonificação disponível quando a coluna "Bonificação" = Sim (se ela não existir, quando BONI está preenchida).
    # O produto mostrado vem de "Nome Produto" (ou, na falta dela, de BONI).
    boni_map = {}
    if entrada.get("boni"):
        mb = map_columns(entrada["boni"], COLS["boni"])
        ib = mb["idx"] if mb else {}
        if not mb or mb["missing"] or (ib["flag"] < 0 and ib["boni"] < 0):
            avisos.append("Planificador de bonificação (%s): não encontrei as colunas UNB_PDV e Bonificação (ou BONI). A bonificação ficou desativada." % nomes.get("boni", ""))
        else:
            SIM = {"sim", "s", "yes", "y", "1", "true", "x", "ok", "disponivel"}
            VAZIOS = ("", "0", "-", "nan", "none", "null")
            valores, ignoradas = {}, 0
            for row in entrada["boni"][mb["sheet"]][mb["hr"] + 1:]:
                code = norm_code(clean(cel(row, ib["pdv"])).split("_")[-1])
                if not code:
                    continue
                boni_txt = clean(cel(row, ib["boni"])) if ib["boni"] >= 0 else ""
                if ib["flag"] >= 0:
                    bruto = clean(cel(row, ib["flag"]))
                    valores[bruto or "(vazio)"] = valores.get(bruto or "(vazio)", 0) + 1
                    n = norm_h(bruto)
                    ok = n in SIM or n.startswith("sim")
                else:
                    ok = boni_txt.lower() not in VAZIOS
                if not ok:
                    ignoradas += 1
                    continue
                prod = clean(cel(row, ib["prod"])) if ib["prod"] >= 0 else ""
                if prod.lower() in VAZIOS:
                    prod = boni_txt if boni_txt.lower() not in VAZIOS else ""
                lst = boni_map.setdefault(code, [])
                if prod and prod not in lst:
                    lst.append(prod)
            txt = describe_cols(mb, COLS["boni"])
            if valores:
                txt += "\n  Valores encontrados em Bonificação: " + ", ".join("%s = %d" % (k, v) for k, v in sorted(valores.items(), key=lambda e: -e[1])[:6])
            log.append({"k": "boni", "txt": txt})
            if ignoradas:
                avisos.append("%d linha(s) do planificador de bonificação sem bonificação disponível foram ignoradas." % ignoradas)
            if not boni_map:
                avisos.append("Planificador de bonificação (%s): nenhuma linha com bonificação disponível (Bonificação = Sim)." % nomes.get("boni", ""))

    # CESTA
    mc = sets["cesta"]; rc = entrada["cesta"][mc["sheet"]]; ic = mc["idx"]
    cesta, sku_idx, cesta_nomes = [], {}, {}
    for row in rc[mc["hr"] + 1:]:
        code = norm_code(cel(row, ic["sku"]))
        if not code or code in sku_idx:
            continue
        desc = clean(cel(row, ic["desc"]))
        if not desc:
            continue
        if ic["cesta"] >= 0:
            cn = clean(cel(row, ic["cesta"])); cesta_nomes[cn] = cesta_nomes.get(cn, 0) + 1
        marca = (clean(cel(row, ic["marca"])).upper() if ic["marca"] >= 0 else "") or infer_marca(desc)
        emb = (clean(cel(row, ic["emb"])) if ic["emb"] >= 0 else "") or infer_emb(desc)
        vol = (clean(cel(row, ic["vol"])) if ic["vol"] >= 0 else "") or infer_vol(desc)
        sku_idx[code] = len(cesta)
        cesta.append([code, desc, marca, emb, vol])
    if not cesta:
        raise ValueError("A cesta não tem nenhum SKU válido.")

    # CLIENTES
    ml = sets["clientes"]; rl = entrada["clientes"][ml["sheet"]]; il = ml["idx"]
    clientes, cli_idx, setores, setor_idx = [], {}, [], {}
    g = lambda row, i: clean(cel(row, i)) if i >= 0 else ""
    for row in rl[ml["hr"] + 1:]:
        code = norm_code(cel(row, il["cod"]))
        if not code or code in cli_idx:
            continue
        sc = g(row, il["setor"]) or "SEM SETOR"
        if sc not in setor_idx:
            setor_idx[sc] = len(setores); setores.append([sc, g(row, il["setorNome"])])
        cli_idx[code] = len(clientes)
        clientes.append([code, g(row, il["fantasia"]), g(row, il["razao"]), setor_idx[sc],
                         parse_dias(cel(row, il["dia"]) if il["dia"] >= 0 else ""),
                         g(row, il["cidade"]), g(row, il["bairro"]), g(row, il["endereco"]), g(row, il["status"]),
                         g(row, il["segmento"]), parse_date(g(row, il["prox"])), score_map.get(code, 0),
                         1 if code in boni_map else 0, boni_map.get(code, [])])
    if boni_map:
        faltam_b = sorted(c for c in boni_map if c not in cli_idx)
        if faltam_b:
            avisos.append("%d PDV(s) do planificador de bonificação não estão no relatório de clientes (%s%s)." % (len(faltam_b), ", ".join(faltam_b[:8]), "…" if len(faltam_b) > 8 else ""))
    if score_map:
        faltam = sorted(c for c in score_map if c not in cli_idx)
        if faltam:
            avisos.append("%d cliente(s) da Base Score 5 não estão no relatório de clientes (%s%s)." % (len(faltam), ", ".join(faltam[:8]), "…" if len(faltam) > 8 else ""))
    ordem = sorted(range(len(setores)), key=lambda i: natural(setores[i][0]))
    remap = {old: new for new, old in enumerate(ordem)}
    setores_ord = [setores[i] for i in ordem]
    for c in clientes:
        c[3] = remap[c[3]]

    # VENDAS
    mv = sets["vendas"]; rv = entrada["vendas"][mv["sheet"]]; iv = mv["idx"]
    agg = {}
    linhas = na_cesta = fora = sem_cad = devol = cancel = 0
    dmin = dmax = None
    sem_cad_set = set()
    for row in rv[mv["hr"] + 1:]:
        cc = norm_code(cel(row, iv["cli"])); sc = norm_code(cel(row, iv["sku"]))
        if not cc or not sc:
            continue
        linhas += 1
        if iv["status"] >= 0 and clean(cel(row, iv["status"])).upper() in STATUS_IGNORAR:
            cancel += 1; continue
        si = sku_idx.get(sc)
        if si is None:
            fora += 1; continue
        ci = cli_idx.get(cc)
        if ci is None:
            sem_cad += 1; sem_cad_set.add(cc); continue
        q = parse_num(cel(row, iv["qtd"])) if iv["qtd"] >= 0 else 1.0
        if iv["qtd"] >= 0 and q <= 0:
            devol += 1; continue
        na_cesta += 1
        val = parse_num(cel(row, iv["valor"])) if iv["valor"] >= 0 else 0.0
        dt = parse_date(cel(row, iv["data"])) if iv["data"] >= 0 else None
        if dt:
            dmin = dt if dmin is None or dt < dmin else dmin
            dmax = dt if dmax is None or dt > dmax else dmax
        m = agg.setdefault(ci, {})
        cur = m.get(si)
        if cur:
            cur[1] += q; cur[2] += val
            if dt and (not cur[3] or dt > cur[3]):
                cur[3] = dt
        else:
            m[si] = [si, q, val, dt]
    dist = []
    for ci in range(len(clientes)):
        m = agg.get(ci)
        dist.append([[d[0], r2(d[1]), r2(d[2]), d[3]] for d in sorted(m.values())] if m else [])
    pares = sum(len(d) for d in dist)

    fmt = lambda n: "{:,}".format(n).replace(",", ".")
    if iv["data"] < 0:
        avisos.append("O relatório de vendas não tem coluna de data: o painel mostra “no período do relatório” no lugar da última venda.")
    if sem_cad_set:
        lst = sorted(sem_cad_set)
        avisos.append("%d cliente(s) com venda não estão no relatório de clientes e ficaram de fora (%s%s)." % (len(lst), ", ".join(lst[:8]), "…" if len(lst) > 8 else ""))
    if cancel:
        avisos.append("%s linha(s) com status %s (cancelada/devolvida) foram desconsideradas." % (fmt(cancel), " ou ".join(sorted(STATUS_IGNORAR))))
    if iv["status"] < 0:
        avisos.append("O relatório de vendas não tem coluna de status: nenhuma linha foi descartada como cancelada ou devolvida.")
    if devol:
        avisos.append("%d linha(s) de venda com quantidade zero ou negativa foram ignoradas." % devol)

    cesta_nome = next((k for k, _ in sorted(cesta_nomes.items(), key=lambda e: -e[1]) if k), "")
    payload = {
        "v": 1,
        "meta": {
            "geradoEm": iso(time.time()), "arquivos": nomes, "cestaNome": cesta_nome,
            "temData": iv["data"] >= 0, "periodo": [dmin, dmax] if dmin else None,
            "stats": {"skus": len(cesta), "clientes": len(clientes), "linhasVenda": linhas, "canceladas": cancel,
                      "linhasCesta": na_cesta, "foraCesta": fora, "semCadastro": sem_cad, "pares": pares,
                      "clientesComVenda": len(agg),
                      "score5": {"total": len(score_map), "casados": sum(1 for c in score_map if c in cli_idx)} if score_map else None,
                      "boni": {"total": len(boni_map), "casados": sum(1 for c in boni_map if c in cli_idx)} if boni_map else None},
        },
        "bases": bases, "cesta": cesta, "setores": setores_ord, "clientes": clientes, "dist": dist,
    }
    return payload, log, avisos


# ----------------------------- leitura de arquivos -----------------------------
def ler_csv(dados):
    if dados.startswith(b"\xef\xbb\xbf"):
        dados = dados[3:]
    try:
        texto = dados.decode("utf-8")
    except UnicodeDecodeError:
        texto = dados.decode("cp1252", errors="replace")
    primeira = texto[:texto.find("\n")] if "\n" in texto else texto
    sep = max([";", ",", "\t", "|"], key=primeira.count)
    linhas = [r for r in csv.reader(io.StringIO(texto, newline=""), delimiter=sep) if any(c.strip() for c in r)]
    return [linhas]


def ler_xlsx(dados):
    NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
    REL = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
    z = zipfile.ZipFile(io.BytesIO(dados))
    shared = []
    if "xl/sharedStrings.xml" in z.namelist():
        for si in ET.fromstring(z.read("xl/sharedStrings.xml")).iter(NS + "si"):
            shared.append("".join(t.text or "" for t in si.iter(NS + "t")))
    rels = {}
    if "xl/_rels/workbook.xml.rels" in z.namelist():
        for r in ET.fromstring(z.read("xl/_rels/workbook.xml.rels")):
            rels[r.get("Id")] = r.get("Target")
    planilhas = []
    for sh in ET.fromstring(z.read("xl/workbook.xml")).iter(NS + "sheet"):
        alvo = rels.get(sh.get(REL + "id"), "")
        alvo = alvo.lstrip("/")
        caminho = alvo if alvo.startswith("xl/") else "xl/" + alvo
        if caminho not in z.namelist():
            continue
        linhas = []
        for row in ET.fromstring(z.read(caminho)).iter(NS + "row"):
            vals = {}
            for c in row.iter(NS + "c"):
                ref = re.match(r"([A-Z]+)", c.get("r", "A"))
                col = 0
                for ch in ref.group(1):
                    col = col * 26 + (ord(ch) - 64)
                t = c.get("t")
                v = c.find(NS + "v")
                if t == "s" and v is not None:
                    txt = shared[int(v.text)]
                elif t == "inlineStr":
                    txt = "".join(x.text or "" for x in c.iter(NS + "t"))
                else:
                    txt = v.text if v is not None and v.text is not None else ""
                    if re.fullmatch(r"-?\d+\.0+", txt):
                        txt = txt.split(".")[0]
                vals[col - 1] = txt
            if vals:
                linha = [""] * (max(vals) + 1)
                for i, x in vals.items():
                    linha[i] = x
                linhas.append(linha)
        planilhas.append(linhas)
    return planilhas


def ler_planilhas(caminho, nome):
    if nome.lower().endswith((".xls", ".xlsb")):
        raise ValueError("formato .xls/.xlsb não é suportado: salve o arquivo como .xlsx")
    dados = ler_arquivo(caminho)
    if nome.lower().endswith((".xlsx", ".xlsm")):
        return ler_xlsx(dados)
    return ler_csv(dados)


# ----------------------------- cache do resultado -----------------------------
_resultado = {"sig": None, "gz": None, "json": None, "erro": "", "trava": threading.Lock()}


def assinatura(d):
    return "|".join("%s@%s" % (d["arquivos"][k]["nome"], d["arquivos"][k]["mtime"]) for k in ("cesta", "vendas", "clientes", "score5", "boni") if k in d["arquivos"])


def processar(forcar=False):
    """Refaz o cruzamento só quando algum dos três arquivos mudou."""
    d = localizar()
    if not d["ok"]:
        _resultado["erro"] = d.get("erro", "")
        return
    sig = assinatura(d)
    if not forcar and sig == _resultado["sig"]:
        return
    with _resultado["trava"]:
        if not forcar and sig == _resultado["sig"]:
            return
        t0 = time.time()
        try:
            chaves = [k for k in ("cesta", "vendas", "clientes", "score5", "boni") if k in d["arquivos"]]
            entrada, extra = {}, []
            for k in chaves:
                try:
                    entrada[k] = ler_planilhas(d["arquivos"][k]["caminho"], d["arquivos"][k]["nome"])
                except Exception as e:
                    if k in ("score5", "boni"):   # opcionais: não derrubam o painel
                        extra.append("Não consegui ler %s (%s): %s" % (d["arquivos"][k]["nome"], k, e))
                    else:
                        raise
            nomes = {k: d["arquivos"][k]["nome"] for k in entrada}
            payload, log, avisos = build_payload(entrada, nomes)
            avisos = extra + avisos
            payload["meta"]["geradoEm"] = iso(d["arquivos"]["vendas"]["mtime"])
            corpo = json.dumps({"sig": sig, "payload": payload, "log": log, "avisos": avisos},
                               ensure_ascii=False, separators=(",", ":")).encode("utf-8")
            _resultado.update(sig=sig, json=corpo, gz=gzip.compress(corpo, 6), erro="")
            st = payload["meta"]["stats"]
            print("%s  dados processados em %.1fs: %d clientes, %d distribuições" % (
                datetime.now().strftime("%d/%m %H:%M:%S"), time.time() - t0, st["clientes"], st["pares"]))
        except Exception as e:
            _resultado["erro"] = "Erro ao processar: %s" % e
            print("ERRO ao processar:", e)


def vigia():
    while True:
        time.sleep(60)
        try:
            processar()
        except Exception as e:
            print("ERRO no vigia:", e)


class Handler(BaseHTTPRequestHandler):
    server_version = "LuviDistribuicao/1.0"

    def _envia(self, codigo, corpo, tipo, extra=None):
        self.send_response(codigo)
        self.send_header("Content-Type", tipo)
        self.send_header("Content-Length", str(len(corpo)))
        self.send_header("Cache-Control", "no-store")
        for k, v in (extra or {}).items():
            self.send_header(k, v)
        self.end_headers()
        if self.command != "HEAD":
            self.wfile.write(corpo)

    def _json(self, obj, codigo=200):
        self._envia(codigo, json.dumps(obj, ensure_ascii=False).encode("utf-8"), "application/json; charset=utf-8")

    def _estatico(self, caminho):
        try:
            caminho = caminho.resolve()
            if BASE not in caminho.parents and caminho != BASE:
                raise FileNotFoundError
            corpo = caminho.read_bytes()
        except (OSError, FileNotFoundError):
            return self._envia(404, b"Nao encontrado", "text/plain; charset=utf-8")
        if caminho.name == "index.html":
            # O mesmo index.html serve para o Vercel (dados embutidos) e para este servidor.
            # Aqui os dados vêm sempre da pasta, então o payload embutido é descartado.
            corpo = re.sub(rb'(<script id="payload" type="application/json">)[\s\S]*?(</script>)',
                           lambda m: m.group(1) + b"null" + m.group(2), corpo, count=1)
        self._envia(200, corpo, TIPOS.get(caminho.suffix.lower(), "application/octet-stream"))

    def do_HEAD(self):
        self.do_GET()

    def do_GET(self):
        rota = unquote(urlparse(self.path).path)
        try:
            if rota in ("/", "/index.html"):
                return self._estatico(BASE / "index.html")
            if rota.startswith("/lib/"):
                return self._estatico(BASE / "lib" / Path(rota).name)
            if rota == "/api/status":
                if "f=1" in (urlparse(self.path).query or ""):
                    _cache["t"] = 0          # "Verificar agora": varre a pasta de novo
                d = localizar()
                arquivos = {k: {"nome": v["nome"], "modificado": iso(v["mtime"]), "tamanho": v["tamanho"]}
                            for k, v in d["arquivos"].items()}
                return self._json({"ok": d["ok"], "erro": d.get("erro") or _resultado["erro"], "pasta": d.get("pasta", ""),
                                   "base_score5": d.get("base_score5"),
                                   "arquivos": arquivos, "sig": assinatura(d) if d["ok"] else "", "agora": iso(time.time())})
            if rota == "/api/dados":
                processar()
                if not _resultado["json"]:
                    return self._json({"erro": _resultado["erro"] or localizar().get("erro") or "Sem dados."}, 503)
                if "gzip" in (self.headers.get("Accept-Encoding") or ""):
                    return self._envia(200, _resultado["gz"], "application/json; charset=utf-8", {"Content-Encoding": "gzip"})
                return self._envia(200, _resultado["json"], "application/json; charset=utf-8")
            return self._envia(404, b"Nao encontrado", "text/plain; charset=utf-8")
        except (ConnectionError, BrokenPipeError):
            pass
        except Exception as e:  # nunca derruba o servidor
            try:
                self._json({"erro": str(e)}, 500)
            except Exception:
                pass

    def log_message(self, fmt, *args):
        if "/api/dados" in (self.path or "") or self.path in ("/", "/index.html"):
            sys.stdout.write("%s  %s  %s\n" % (datetime.now().strftime("%d/%m %H:%M:%S"),
                                               self.client_address[0], self.path.split("?")[0]))
            sys.stdout.flush()


def ip_local():
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(("8.8.8.8", 80))
        ip = s.getsockname()[0]
        s.close()
        return ip
    except OSError:
        return "127.0.0.1"


if __name__ == "__main__":
    print("=" * 60)
    print("  Painel de Distribuição de Cerveja - LUVI BI")
    print("=" * 60)
    d = localizar()
    print("Pasta:", d.get("pasta") or "NÃO ENCONTRADA")
    for k in ("cesta", "vendas", "clientes", "score5", "boni"):
        a = d["arquivos"].get(k)
        if a is None and k in ("score5", "boni"):
            print("  %-9s %s" % (k, "(opcional, não encontrada)"))
            continue
        print("  %-9s %s" % (k, ("%s  (%s)" % (a["nome"], iso(a["mtime"])[:16].replace("T", " "))) if a else "-- não encontrado --"))
    if not (BASE / "index.html").exists():
        print("\nATENÇÃO: index.html não está ao lado do servidor.py")
    print("\nAbra no navegador:")
    print("  Nesta máquina : http://localhost:%d" % PORTA)
    print("  Na rede       : http://%s:%d" % (ip_local(), PORTA))
    print("\n(Deixe esta janela aberta. Ctrl+C para parar.)\n")
    print("Processando os arquivos pela primeira vez...")
    processar(forcar=True)
    threading.Thread(target=vigia, daemon=True).start()
    # No Windows, vários processos conseguem abrir a MESMA porta (e cada acesso cai num deles, às vezes
    # num servidor antigo). Aqui o segundo servidor se recusa a subir e avisa.
    ThreadingHTTPServer.allow_reuse_address = False
    try:
        srv = ThreadingHTTPServer(("0.0.0.0", PORTA), Handler)
    except OSError:
        print("\n*** A PORTA %d JÁ ESTÁ EM USO: há outro servidor rodando. ***" % PORTA)
        print("Descubra o PID na coluna LISTENING e derrube só ele:")
        print("    netstat -ano | findstr :%d" % PORTA)
        print("    taskkill /F /PID <número da última coluna>\n")
        sys.exit(1)
    srv.daemon_threads = True
    try:
        srv.serve_forever()
    except KeyboardInterrupt:
        print("\nServidor encerrado.")
