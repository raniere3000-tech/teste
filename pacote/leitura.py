"""Lê uma planilha, reconhece se é OS executada ou base de ligações e devolve registros enxutos."""
import json
import os
import re
import unicodedata

import pandas as pd

_AQUI = os.path.dirname(os.path.abspath(__file__))
with open(os.path.join(_AQUI, "regras.json"), encoding="utf-8") as _f:
    REGRAS = json.load(_f)


def norm(v):
    s = "" if v is None else str(v)
    s = unicodedata.normalize("NFD", s)
    s = "".join(c for c in s if unicodedata.category(c) != "Mn")
    return re.sub(r"\s+", " ", s).strip().upper()


def chave_mat(v):
    s = "" if v is None else str(v).strip()
    if re.fullmatch(r"\d+\.0", s):
        s = s[:-2]
    s = re.sub(r"\D", "", s).lstrip("0")
    return s


def _brutas(caminho):
    """Devolve [(nome_aba, DataFrame sem cabeçalho)]."""
    ext = caminho.lower().rsplit(".", 1)[-1]
    if ext == "csv":
        for enc in ("utf-8-sig", "latin-1"):
            try:
                df = pd.read_csv(caminho, header=None, dtype=str, sep=None, engine="python", encoding=enc)
                return [("csv", df)]
            except UnicodeDecodeError:
                continue
        raise ValueError("não foi possível decodificar o CSV")
    motor = "xlrd" if ext == "xls" else "openpyxl"
    abas = pd.read_excel(caminho, sheet_name=None, header=None, dtype=str, engine=motor)
    return list(abas.items())


def _acha_cabecalho(df, obrigatorias):
    obr = [norm(c) for c in obrigatorias]
    for i in range(min(20, len(df))):
        linha = {norm(x) for x in df.iloc[i].tolist()}
        if all(o in linha for o in obr):
            return i
    return None


def _monta(df, i):
    nomes, vistos = [], {}
    for x in df.iloc[i].tolist():
        n = "" if x is None or (isinstance(x, float) and pd.isna(x)) else str(x).strip()
        k = norm(n)
        if k in vistos:
            vistos[k] += 1
            n = f"{n}__{vistos[k]}"
        else:
            vistos[k] = 0
        nomes.append(n)
    out = df.iloc[i + 1:].copy()
    out.columns = nomes
    return out.fillna("")


def _col(df, nome):
    alvo = norm(nome)
    for c in df.columns:
        if norm(c) == alvo:
            return c
    return None


def _data_iso(v):
    if v is None or str(v).strip() == "":
        return ""
    s = str(v).strip()
    try:
        if re.match(r"^\d{4}-\d{2}-\d{2}", s):
            d = pd.to_datetime(s[:19], errors="coerce")
        elif re.fullmatch(r"\d+(\.\d+)?", s):  # número de série do Excel
            d = pd.Timestamp("1899-12-30") + pd.to_timedelta(float(s), unit="D")
        else:
            d = pd.to_datetime(s, dayfirst=True, errors="coerce")
    except Exception:
        return ""
    return "" if pd.isna(d) else d.strftime("%Y-%m-%d")


def _os(df):
    R = REGRAS["os"]
    C = {k: _col(df, v) for k, v in R["colunas"].items()}
    g = lambda row, k: str(row[C[k]]).strip() if C[k] else ""
    alvo = norm(R["codigo_alvo"])
    execs = {norm(s) for s in R["status_executada"]}
    impl = norm(R["implantacao"])
    regs, fora = [], 0
    for _, row in df.iterrows():
        if alvo not in norm(g(row, "codigo")):
            fora += 1
            continue
        st = g(row, "status")
        e = norm(st) in execs
        regs.append({
            "id": g(row, "id"), "r": g(row, "recurso"), "m": chave_mat(g(row, "matricula")),
            "d": _data_iso(g(row, "data")), "c": g(row, "cidade"), "b": g(row, "bairro"),
            "en": g(row, "endereco"), "cl": g(row, "cliente"), "s": st,
            "e": e, "i": e and impl in norm(g(row, "adicionais")),
            "cod": g(row, "codigo"), "ad": g(row, "adicionais"), "fim": g(row, "fim"),
        })
    faltando = [R["colunas"][k] for k, v in C.items() if v is None]
    return regs, {"fora_do_codigo": fora, "colunas_ausentes": faltando}


def _base(df):
    R = REGRAS["base"]
    C = {k: _col(df, v) for k, v in R["colunas"].items()}
    g = lambda row, k: str(row[C[k]]).strip() if C[k] else ""
    pad = norm(R["faturamento_padronizacao_contem"])
    nao = {norm(x) for x in R["engenharia_valores_nao"]}
    regs = []
    for _, row in df.iterrows():
        m = chave_mat(g(row, "matricula"))
        if not m:
            continue
        fr = g(row, "faturamento")
        cl = "SEM" if not norm(fr) else ("PAD" if pad in norm(fr) else "IND")
        regs.append({
            "m": m, "cl": g(row, "cliente"), "c": g(row, "cidade"), "b": g(row, "bairro"),
            "rua": g(row, "logradouro") or "SEM LOGRADOURO", "en": g(row, "endereco"),
            "f": cl, "fr": fr, "eng": norm(g(row, "engenharia")) not in nao, "sit": g(row, "situacao"),
            "sub": g(row, "subcategoria"), "eco": g(row, "economias"), "engv": g(row, "engenharia"),
        })
    faltando = [R["colunas"][k] for k, v in C.items() if v is None]
    return regs, {"colunas_ausentes": faltando}


def le_arquivo(caminho, nome):
    """Retorna dict: tipo ('os' | 'base' | 'nao_usado'), registros, linhas, info, motivo."""
    base_nome = norm(os.path.basename(nome))
    if base_nome.startswith(norm(REGRAS["prefixo_bases_geradas"])):
        return {"arquivo": nome, "tipo": "nao_usado", "motivo": "base gerada pelo próprio painel", "registros": [], "linhas": 0}
    try:
        abas = _brutas(caminho)
    except Exception as e:  # noqa: BLE001
        return {"arquivo": nome, "tipo": "nao_usado", "motivo": f"erro ao abrir: {e}", "registros": [], "linhas": 0}
    # tenta pelas colunas; o nome do arquivo decide a ordem de tentativa
    ordem = ["os", "base"]
    if any(norm(p) in base_nome for p in REGRAS["base"]["reconhece_nome"]) and not any(
            norm(p) in base_nome for p in REGRAS["os"]["reconhece_nome"]):
        ordem = ["base", "os"]
    for aba, bruto in abas:
        for tipo in ordem:
            i = _acha_cabecalho(bruto, REGRAS[tipo]["colunas_obrigatorias"])
            if i is None:
                continue
            df = _monta(bruto, i)
            regs, info = (_os if tipo == "os" else _base)(df)
            info["aba"] = aba
            return {"arquivo": nome, "tipo": tipo, "registros": regs, "linhas": len(df), "info": info, "motivo": ""}
    obr = " / ".join(f"{REGRAS[t]['rotulo']}: {', '.join(REGRAS[t]['colunas_obrigatorias'])}" for t in ("os", "base"))
    return {"arquivo": nome, "tipo": "nao_usado", "registros": [], "linhas": 0,
            "motivo": f"cabeçalho não reconhecido (procuro {obr})"}


def le_json(caminho, nome):
    return json.dumps(le_arquivo(caminho, nome), ensure_ascii=False)
