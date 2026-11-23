"""
Gera os arquivos que alimentam a pagina Computadores para Inclusao, a partir do BI
(data/doacoes.csv e data/formacao.csv, baixados por scripts/extrair.py).

SAIDAS
  data/resumo.json   -> totais, por ano e por UF (cartoes, metas e barras da pagina)
  data/muns.json     -> por municipio: [chave, UF, nome, computadores, IMPACTO SOCIAL] e alunos (mapa)
  data/meta.json     -> data de referencia e conferencias
  Auditoria (para conferir, nao vao para a pagina):
  data/impacto_pontos.csv          -> como CADA ponto foi contado no impacto social
  data/escolas_para_consulta.csv   -> escolas (codigo INEP) que o robo do QEdu precisa consultar
  data/escolas_pendentes.csv       -> escolas sem matricula confirmada na fonte
  data/escolas_lista_sem_doacao.csv-> escolas da lista da area que nao aparecem no BI
  data/municipios_corrigidos.csv   -> nomes de municipio corrigidos por regra automatica
  data/municipios_sugestoes.csv    -> municipios sem codigo IBGE, com sugestao (NAO aplicada)
  data/sem_ibge.csv                -> municipios sem codigo IBGE

=================== IMPACTO SOCIAL (regra da area gestora - CGID/DEPIN) ===================
Texto da planilha da area (aba "Dados dos Projetos"): "Considera-se que 1 computador atende 5 pessoas
e para instituicoes de ensino consideramos o numero de matriculas pelo site do QEdu.br".
  - Ponto que NAO e instituicao de ensino -> 5 x computadores doados ao ponto
  - Instituicao de ensino                  -> numero de MATRICULAS no QEdu, pelo codigo INEP
Quem e instituicao de ensino: data/fixos/LISTA_ESCOLAS.xlsx (lista da area; TIPO = INST. DE ENSINO).
Matriculas: data/fixos/matriculas_qedu.csv (preenchido pelo robo scripts/buscar_matriculas_qedu.py).
            Excecao com fonte documentada: data/fixos/matriculas_manual.csv (exige fonte e link).
Cada escola conta UMA vez (pelo codigo INEP), mesmo que tenha recebido doacoes em varios anos.
Escola sem matricula confirmada -> o municipio fica PENDENTE (impacto = null) e a pagina mantem o
valor que ja tinha. Nenhum numero de matricula e estimado.
==========================================================================================
"""
import csv, datetime as dt, difflib, json, os, re, sys, unicodedata
import pandas as pd

HOJE = os.environ.get("HOJE") or dt.date.today().isoformat()
MIN_DOACOES, MIN_ALUNOS = 1000, 1000          # protecao: abaixo disso nao grava nada
PESSOAS_POR_COMPUTADOR = 5                    # regra da area para pontos que nao sao escola
FIXOS = "data/fixos"
LISTA_ESCOLAS = f"{FIXOS}/LISTA_ESCOLAS.xlsx"
MATRICULAS_QEDU = f"{FIXOS}/matriculas_qedu.csv"
MATRICULAS_MANUAL = f"{FIXOS}/matriculas_manual.csv"
IBGE_OFICIAL = f"{FIXOS}/ibge_municipios.csv"  # lista oficial completa (baixada da API do IBGE)
IBGE_RESERVA = f"{FIXOS}/ibge.csv"             # reserva: municipios que ja estavam na pagina
ALIAS = f"{FIXOS}/municipios_alias.csv"

UF_POR_CODIGO = {"11": "RO", "12": "AC", "13": "AM", "14": "RR", "15": "PA", "16": "AP", "17": "TO", "21": "MA",
                 "22": "PI", "23": "CE", "24": "RN", "25": "PB", "26": "PE", "27": "AL", "28": "SE", "29": "BA",
                 "31": "MG", "32": "ES", "33": "RJ", "35": "SP", "41": "PR", "42": "SC", "43": "RS", "50": "MS",
                 "51": "MT", "52": "GO", "53": "DF"}
UF_NOME = {"AC": "Acre", "AL": "Alagoas", "AM": "Amazonas", "AP": "Amapa", "BA": "Bahia", "CE": "Ceara",
           "DF": "Distrito Federal", "ES": "Espirito Santo", "GO": "Goias", "MA": "Maranhao", "MG": "Minas Gerais",
           "MS": "Mato Grosso do Sul", "MT": "Mato Grosso", "PA": "Para", "PB": "Paraiba", "PE": "Pernambuco",
           "PI": "Piaui", "PR": "Parana", "RJ": "Rio de Janeiro", "RN": "Rio Grande do Norte", "RO": "Rondonia",
           "RR": "Roraima", "RS": "Rio Grande do Sul", "SC": "Santa Catarina", "SE": "Sergipe", "SP": "Sao Paulo",
           "TO": "Tocantins"}
# Regioes administrativas do DF: o DF tem um unico municipio (Brasilia, 5300108)
DF_RAS = """Plano Piloto|Brasilia|Asa Norte|Asa Sul|Gama|Taguatinga|Brazlandia|Sobradinho|Sobradinho II|Planaltina|
Paranoa|Nucleo Bandeirante|Ceilandia|Guara|Cruzeiro|Samambaia|Santa Maria|Sao Sebastiao|Recanto das Emas|Lago Sul|
Riacho Fundo|Riacho Fundo II|Lago Norte|Candangolandia|Aguas Claras|Sudoeste|Octogonal|Sudoeste/Octogonal|Varjao|
Park Way|SCIA|Estrutural|Jardim Botanico|Itapoa|SIA|Vicente Pires|Fercal|Sol Nascente|Por do Sol|
Sol Nascente/Por do Sol|Arniqueira|Arapoanga|Agua Quente"""


def norm(s):
    s = unicodedata.normalize("NFD", str(s if s is not None else ""))
    s = "".join(c for c in s if unicodedata.category(c) != "Mn").upper()
    return re.sub(r"[^A-Z0-9]+", " ", s).strip()


def compacto(s):
    return norm(s).replace(" ", "")


PEQUENAS = {"DE", "DA", "DO", "DAS", "DOS", "E", "D"}


def titulo(s):
    partes = []
    for i, p in enumerate(re.split(r"(\s+|'|-)", str(s).strip())):
        if p in ("", None) or re.fullmatch(r"\s+|'|-", p):
            partes.append(p)
        elif i > 0 and p.upper() in PEQUENAS:
            partes.append(p.lower())
        else:
            partes.append(p[:1].upper() + p[1:].lower())
    return "".join(partes)


def ler(caminho, colunas):
    d = pd.read_csv(caminho, dtype=str, keep_default_na=False)
    faltam = [c for c in colunas if c not in d.columns]
    if faltam:
        sys.exit(f"{caminho}: faltam as colunas {faltam}. Colunas encontradas: {list(d.columns)}")
    return d


def gravar_csv(caminho, cabecalho, linhas):
    with open(caminho, "w", newline="", encoding="utf-8") as fh:
        w = csv.writer(fh)
        w.writerow(cabecalho)
        w.writerows(linhas)


# =============================== MUNICIPIOS ===============================
def baixar_ibge_oficial():
    """Baixa a lista oficial de municipios (API do IBGE) uma vez e guarda em data/fixos/ibge_municipios.csv.
    Para atualizar, basta apagar o arquivo: o robo baixa de novo."""
    try:
        import requests
        r = requests.get("https://servicodados.ibge.gov.br/api/v1/localidades/municipios", timeout=60)
        r.raise_for_status()
        linhas = []
        for m in r.json():
            cod = str(m.get("id") or "").strip()
            uf = UF_POR_CODIGO.get(cod[:2])
            if len(cod) == 7 and uf and m.get("nome"):
                linhas.append([cod, uf, m["nome"]])
        if len(linhas) < 5500:
            raise ValueError(f"lista incompleta ({len(linhas)} municipios)")
        gravar_csv(IBGE_OFICIAL, ["codigo", "uf", "nome"], sorted(linhas))
        print(f"Lista oficial do IBGE baixada: {len(linhas)} municipios")
        return True
    except Exception as e:
        print(f"Aviso: nao consegui baixar a lista oficial do IBGE ({e}). Usando {IBGE_RESERVA}.")
        return False


def carregar_municipios():
    if not os.path.exists(IBGE_OFICIAL):
        baixar_ibge_oficial()
    caminho = IBGE_OFICIAL if os.path.exists(IBGE_OFICIAL) else IBGE_RESERVA
    idx, por_uf, por_nome = {}, {}, {}
    with open(caminho, encoding="utf-8") as fh:
        for r in csv.DictReader(fh):
            k = (r["uf"], compacto(r["nome"]))
            if k in idx:
                continue
            idx[k] = (r["codigo"], r["nome"])
            por_uf.setdefault(r["uf"], {})[compacto(r["nome"])] = r["nome"]
            por_nome.setdefault(compacto(r["nome"]), []).append((r["uf"], r["nome"]))
    print(f"Municipios de referencia: {len(idx)} ({caminho})")
    return idx, por_uf, por_nome


def carregar_alias():
    """data/fixos/municipios_alias.csv: grafias do BI -> nome oficial. So vale a linha com aplicar = sim."""
    out = {}
    if os.path.exists(ALIAS):
        with open(ALIAS, encoding="utf-8-sig") as fh:
            for r in csv.DictReader(fh):
                if norm(r.get("aplicar")) in ("SIM", "S", "X"):
                    out[(r["uf_bi"].strip().upper(), compacto(r["nome_bi"]))] = (r["uf"].strip().upper(), r["nome_oficial"], r.get("motivo", ""))
    return out


OF, OF_UF, OF_NOME = carregar_municipios()
ALIASES = carregar_alias()
DF_RAS_C = {compacto(x) for x in re.split(r"[|\n]", DF_RAS) if x.strip()}
sem_ibge, corrigidos, _cache_mun = {}, {}, {}


def resolver(uf, nome_bi):
    """Devolve ((codigo, nome_oficial), regra_aplicada) ou (None, None)."""
    uf = (uf or "").strip().upper()
    c = compacto(nome_bi)
    a = ALIASES.get((uf, c))
    if a and (a[0], compacto(a[1])) in OF:
        return OF[(a[0], compacto(a[1]))], f"tabela de correcoes: {a[2]}".strip()
    if (uf, c) in OF:
        return OF[(uf, c)], None
    if uf == "DF" and c in DF_RAS_C and ("DF", "BRASILIA") in OF:
        return OF[("DF", "BRASILIA")], "regiao administrativa do DF -> Brasilia"
    estado = compacto(UF_NOME.get(uf, ""))
    cand = []
    if uf and c.endswith(uf):                      # "Cascavel-PR", "Barra de Sao Miguel-AL"
        cand += [c[:-2], c[:-2] + estado]          # "Sao Domingos do MA" -> "...do Maranhao"
    if estado and c.endswith(estado):              # "Orizona Goias"
        cand.append(c[:-len(estado)])
    if estado:                                     # "Alto Paraiso" -> "Alto Paraiso de Goias"
        cand += [c + p + estado for p in ("DO", "DE", "DA")]
    achados = {OF[(uf, x)] for x in cand if x and (uf, x) in OF}
    if len(achados) == 1:
        return achados.pop(), "nome com/sem o nome do estado"
    return None, None


def chave(uf, mun):
    k0 = ((uf or "").strip().upper(), str(mun))
    if k0 in _cache_mun:
        return _cache_mun[k0]
    hit, regra = resolver(uf, mun)
    if hit:
        res = (hit[0], hit[1])
        if regra:
            corrigidos[(k0[0], str(mun))] = (hit[0], hit[1], regra)
    else:
        k = f"B|{k0[0]}|{norm(mun)}"
        sem_ibge[k] = (k0[0], titulo(mun))
        res = (k, titulo(mun))
    _cache_mun[k0] = res
    return res


def sugestoes():
    out = []
    for k, (uf, nome) in sorted(sem_ibge.items()):
        c = compacto(nome)
        for uf2, n2 in OF_NOME.get(c, []):
            if uf2 != uf:
                out.append([uf, nome, uf2, n2, "mesmo nome em outra UF (UF do BI pode estar errada)"])
        prox = difflib.get_close_matches(c, list(OF_UF.get(uf, {}).keys()), n=2, cutoff=0.82)
        for p in prox:
            out.append([uf, nome, uf, OF_UF[uf][p], "grafia parecida na mesma UF"])
        if not OF_NOME.get(c) and not prox:
            out.append([uf, nome, "", "", "sem sugestao (pode ser distrito/localidade: conferir)"])
    return out


# =============================== ESCOLAS ===============================
def so_digitos(v):
    s = str(v if v is not None else "").strip()
    if re.fullmatch(r"\d+\.0+", s):
        s = s.split(".")[0]
    return re.sub(r"\D", "", s)


def carregar_escolas():
    """Le a lista da area. Devolve dict (UF, municipio, ponto) -> info, ou None se o arquivo nao existir."""
    if not os.path.exists(LISTA_ESCOLAS):
        print(f"Aviso: {LISTA_ESCOLAS} nao encontrado. Sem a lista da area o impacto social NAO e calculado.")
        return None
    abas = pd.read_excel(LISTA_ESCOLAS, sheet_name=None, dtype=str)
    escolhida = None
    for nome_aba, df in abas.items():
        cols = {compacto(c): c for c in df.columns}
        if any("INEP" in c for c in cols) and any("NOME" in c for c in cols):
            escolhida = (nome_aba, df, cols)
            break
    if not escolhida:
        sys.exit(f"{LISTA_ESCOLAS}: nenhuma aba com colunas de NOME e INEP. Abas: {list(abas)}")
    nome_aba, df, cols = escolhida

    def col(*preferidas, contem=None):
        for p in preferidas:
            if p in cols:
                return cols[p]
        if contem:
            for c, orig in cols.items():
                if contem in c:
                    return orig
        return None

    c_nome, c_uf = col("NOMEPID", contem="NOME"), col("UFPID", "UF")
    c_mun, c_inep = col("MUNICIPIOPID", contem="MUNICIPIO"), col(contem="INEP")
    c_tipo, c_ibge = col("TIPO"), col("CDIBGE", contem="IBGE")
    if not all([c_nome, c_uf, c_mun, c_inep]):
        sys.exit(f"{LISTA_ESCOLAS} (aba {nome_aba}): faltam colunas. Encontradas: {list(df.columns)}")
    if c_tipo:
        df = df[df[c_tipo].fillna("").map(compacto).str.contains("ENSINO")]
    escolas = {}
    for _, r in df.iterrows():
        uf = str(r[c_uf] or "").strip().upper()
        k = (uf, compacto(r[c_mun]), compacto(r[c_nome]))
        inep = so_digitos(r[c_inep])
        e = escolas.setdefault(k, {"uf": uf, "municipio": str(r[c_mun]).strip(), "ponto": str(r[c_nome]).strip(),
                                   "ineps": set(), "ibge": so_digitos(r[c_ibge]) if c_ibge else ""})
        if inep:
            e["ineps"].add(inep)
    print(f"Lista de escolas: {len(escolas)} escolas distintas (aba '{nome_aba}')")
    return escolas


def carregar_matriculas():
    """Codigo INEP -> {matriculas, ano, fonte, url}. So entra o que tem status ok (QEdu) ou fonte+link (manual)."""
    out = {}
    if os.path.exists(MATRICULAS_QEDU):
        with open(MATRICULAS_QEDU, encoding="utf-8-sig") as fh:
            for r in csv.DictReader(fh):
                m = so_digitos(r.get("matriculas"))
                if r.get("status") == "ok" and m:
                    out[so_digitos(r["inep"])] = {"matriculas": int(m), "ano": r.get("ano_censo", ""),
                                                  "fonte": "QEdu", "url": r.get("url", "")}
    manual = {}
    if os.path.exists(MATRICULAS_MANUAL):
        with open(MATRICULAS_MANUAL, encoding="utf-8-sig") as fh:
            for r in csv.DictReader(fh):
                m = so_digitos(r.get("matriculas"))
                if not (m and (r.get("fonte") or "").strip() and (r.get("url") or "").strip()):
                    continue                     # sem fonte e link, nao vale
                info = {"matriculas": int(m), "ano": r.get("ano", ""), "fonte": r["fonte"].strip(), "url": r["url"].strip()}
                if so_digitos(r.get("inep")):
                    out[so_digitos(r["inep"])] = info
                else:
                    manual[(str(r.get("uf", "")).strip().upper(), compacto(r.get("municipio")), compacto(r.get("ponto")))] = info
    return out, manual


# =============================== DADOS DO BI ===============================
d = ler("data/doacoes.csv", ["ANO_PID", "UF_PID", "MUNICIPIO_PID", "NOME_PID", "Soma_QTD_DOAÇÕES"])
f = ler("data/formacao.csv", ["ANO", "UF_CRC", "MUNICIPIO", "NOME_CRC", "CURSO_JOV", "Soma_CONTADOR", "GENERO_MASC", "GENERO_FEM"])
d["q"] = pd.to_numeric(d["Soma_QTD_DOAÇÕES"], errors="coerce").fillna(0).astype(int)
f["q"] = pd.to_numeric(f["Soma_CONTADOR"], errors="coerce").fillna(0).astype(int)
for c in ("GENERO_MASC", "GENERO_FEM"):
    f[c] = pd.to_numeric(f[c], errors="coerce").fillna(0).astype(int)
if d.q.sum() < MIN_DOACOES or f.q.sum() < MIN_ALUNOS:
    sys.exit(f"Totais baixos demais (doacoes={d.q.sum()}, alunos={f.q.sum()}). Nada foi gravado.")

kd = [chave(u, m) for u, m in zip(d.UF_PID, d.MUNICIPIO_PID)]
d["key"], d["nome"] = [k[0] for k in kd], [k[1] for k in kd]
kf = [chave(u, m) for u, m in zip(f.UF_CRC, f.MUNICIPIO)]
f["key"], f["nome"] = [k[0] for k in kf], [k[1] for k in kf]
d["pid"] = d["key"] + "|" + d["NOME_PID"]

# =============================== IMPACTO SOCIAL ===============================
ESCOLAS = carregar_escolas()
MATR, MATR_MANUAL = carregar_matriculas()

pontos = (d.assign(uf=d.UF_PID.str.strip().str.upper(), mun_c=d.MUNICIPIO_PID.map(compacto), pt_c=d.NOME_PID.map(compacto))
           .groupby(["key", "uf", "mun_c", "pt_c"], as_index=False)
           .agg(computadores=("q", "sum"), ponto=("NOME_PID", "first"), municipio=("nome", "first")))

# indice auxiliar: escola por (UF, nome) quando o municipio vier escrito diferente (so se for unico na UF)
por_uf_nome = {}
if ESCOLAS:
    for (uf, mc, pc), e in ESCOLAS.items():
        por_uf_nome.setdefault((uf, pc), []).append((uf, mc, pc))

linhas_imp, contados, usadas = [], set(), set()
for r in pontos.sort_values(["uf", "municipio", "ponto"]).itertuples(index=False):
    comp = int(r.computadores)
    base = [r.uf, r.municipio, r.key, r.ponto, comp]
    if ESCOLAS is None:
        linhas_imp.append(base + ["", "", "sem lista de escolas", "", "", "", None, "pendente"])
        continue
    k = (r.uf, r.mun_c, r.pt_c)
    if k not in ESCOLAS and len(por_uf_nome.get((r.uf, r.pt_c), [])) == 1:
        k = por_uf_nome[(r.uf, r.pt_c)][0]
    if k not in ESCOLAS:                                   # nao e escola -> 5 x computadores
        linhas_imp.append(base + ["nao", "", f"{PESSOAS_POR_COMPUTADOR} x computadores", "", "", "",
                                  PESSOAS_POR_COMPUTADOR * comp, "ok"])
        continue
    usadas.add(k)
    e = ESCOLAS[k]
    ineps = sorted(e["ineps"])
    info = MATR_MANUAL.get(k)
    inep = ineps[0] if len(ineps) == 1 else ""
    if not info and len(ineps) == 1:
        info = MATR.get(inep)
    if len(ineps) > 1:
        st, regra, imp = "pendente", "escola com mais de um codigo INEP na lista (conferir)", None
    elif not info:
        st = "pendente"
        regra = "escola sem codigo INEP na lista" if not ineps else (
            "codigo nao e INEP de escola basica (QEdu nao cobre): usar matriculas_manual.csv" if len(inep) != 8
            else "matricula ainda nao confirmada no QEdu")
        imp = None
    elif (inep or k) in contados:
        st, regra, imp = "ok", "escola ja contada em outro ponto (matricula conta uma vez)", 0
    else:
        contados.add(inep or k)
        st, regra, imp = "ok", f"matriculas ({info['fonte']})", info["matriculas"]
    linhas_imp.append(base + ["sim", ";".join(ineps), regra, info["matriculas"] if info else "",
                              info["ano"] if info else "", info["url"] if info else "", imp, st])

COLS_IMP = ["uf", "municipio", "codigo_ibge", "ponto", "computadores", "escola", "inep", "regra",
            "matriculas", "ano_censo", "url_fonte", "impacto", "status"]
imp = pd.DataFrame(linhas_imp, columns=COLS_IMP)
pend_mun = set(imp.loc[imp.status == "pendente", "codigo_ibge"])
imp_mun = {k: (None if k in pend_mun else int(pd.to_numeric(g.impacto).sum())) for k, g in imp.groupby("codigo_ibge")}

# =============================== RESUMO ===============================
def por_uf(df, col, uf):
    return {str(k): int(v) for k, v in df.groupby(uf)[col].sum().sort_index().items()}


def b_doacoes(x):
    return {"computadores": int(x.q.sum()), "municipios": int(x.key.nunique()), "pontos": int(x.pid.nunique())}


def b_formacao(x):
    return {"alunos": int(x.q.sum()), "masculino": int(x.GENERO_MASC.sum()), "feminino": int(x.GENERO_FEM.sum()),
            "cursos_distintos": int(x.CURSO_JOV.nunique()), "municipios": int(x.key.nunique())}


anos = sorted(set(d.ANO_PID) | set(f.ANO))
resumo = {
    "gerado_em": HOJE,
    "doacoes": {**b_doacoes(d), "por_ano": {a: b_doacoes(d[d.ANO_PID == a]) for a in anos if (d.ANO_PID == a).any()},
                "por_uf": por_uf(d, "q", "UF_PID")},
    "formacao": {**b_formacao(f), "por_ano": {a: b_formacao(f[f.ANO == a]) for a in anos if (f.ANO == a).any()},
                 "por_uf": por_uf(f, "q", "UF_CRC")},
}


def lista_mun(df, ufcol, com_impacto=False):
    ufs = df.drop_duplicates("key").set_index("key")[ufcol].str.strip().str.upper().to_dict()
    g = df.groupby(["key", "nome"], as_index=False)["q"].sum()
    g = g[g["q"] > 0]
    out = []
    for r in g.itertuples(index=False):
        uf = r.key and (UF_POR_CODIGO.get(str(r.key)[:2]) if str(r.key).isdigit() else ufs[r.key])
        linha = [r.key, uf, r.nome, int(r.q)]
        if com_impacto:
            linha.append(imp_mun.get(r.key))   # impacto social (null = pendente)
        out.append(linha)
    return sorted(out, key=lambda x: (x[1], x[2]))


muns = {"gerado_em": HOJE,
        "regra_impacto": f"instituicao de ensino = matriculas no QEdu (codigo INEP, conta uma vez); demais = {PESSOAS_POR_COMPUTADOR} x computadores",
        "doados": lista_mun(d, "UF_PID", com_impacto=True), "formacao": lista_mun(f, "UF_CRC"),
        "sem_ibge": sorted([[k, v[0], v[1]] for k, v in sem_ibge.items()])}
assert sum(x[3] for x in muns["doados"]) == resumo["doacoes"]["computadores"]
assert sum(x[3] for x in muns["formacao"]) == resumo["formacao"]["alunos"]
resumo["doacoes"]["municipios"] = len(muns["doados"])
resumo["formacao"]["municipios"] = len(muns["formacao"])

esc = imp[imp.escola == "sim"]
calc = [x[4] for x in muns["doados"] if x[4] is not None]
resumo["doacoes"]["impacto_social"] = {
    "regra": muns["regra_impacto"],
    "pessoas_total": int(sum(calc)) if len(calc) == len(muns["doados"]) else None,   # so quando nao ha pendencia
    "pessoas_municipios_calculados": int(sum(calc)),
    "municipios_calculados": len(calc),
    "municipios_pendentes": len(muns["doados"]) - len(calc),
    "pessoas_escolas": int(pd.to_numeric(esc.impacto).fillna(0).sum()),
    "pessoas_demais_5x": int(pd.to_numeric(imp[imp.escola == "nao"].impacto).fillna(0).sum()),
    "escolas_pontos": int(len(esc)),
    "escolas_com_matricula": int((esc.status == "ok").sum()),
    "escolas_pendentes": int((esc.status == "pendente").sum()),
    "anos_censo": sorted({str(a) for a in esc.ano_censo if str(a).strip()}),
}

# =============================== GRAVAR ===============================
os.makedirs("data", exist_ok=True)
with open("data/resumo.json", "w", encoding="utf-8") as fh:
    json.dump(resumo, fh, ensure_ascii=False, indent=1)
with open("data/muns.json", "w", encoding="utf-8") as fh:
    json.dump(muns, fh, ensure_ascii=False, separators=(",", ":"))
imp.to_csv("data/impacto_pontos.csv", index=False, encoding="utf-8")
gravar_csv("data/escolas_pendentes.csv", ["uf", "municipio", "ponto", "inep", "motivo"],
           [[r.uf, r.municipio, r.ponto, r.inep, r.regra] for r in esc[esc.status == "pendente"].itertuples(index=False)])
consulta = sorted({(so_digitos(i), r.ponto, r.uf, r.municipio) for r in esc.itertuples(index=False)
                   for i in str(r.inep).split(";") if len(so_digitos(i)) == 8})
gravar_csv("data/escolas_para_consulta.csv", ["inep", "nome", "uf", "municipio"], consulta)
if ESCOLAS:
    gravar_csv("data/escolas_lista_sem_doacao.csv", ["uf", "municipio", "ponto", "inep"],
               sorted([[e["uf"], e["municipio"], e["ponto"], ";".join(sorted(e["ineps"]))]
                       for k, e in ESCOLAS.items() if k not in usadas]))
gravar_csv("data/municipios_corrigidos.csv", ["uf_bi", "nome_bi", "codigo_ibge", "nome_oficial", "regra"],
           sorted([[k[0], k[1], v[0], v[1], v[2]] for k, v in corrigidos.items()]))
gravar_csv("data/municipios_sugestoes.csv", ["uf_bi", "nome_bi", "sugestao_uf", "sugestao_nome", "motivo"], sugestoes())
gravar_csv("data/sem_ibge.csv", ["chave", "uf", "nome"], muns["sem_ibge"])
meta = {"data_referencia": HOJE, "gerado_em": HOJE, "linhas_doacoes": int(len(d)), "linhas_formacao": int(len(f)),
        "municipios_doacoes": len(muns["doados"]), "municipios_formacao": len(muns["formacao"]),
        "sem_ibge": len(sem_ibge), "municipios_corrigidos": len(corrigidos),
        "escolas_com_matricula": resumo["doacoes"]["impacto_social"]["escolas_com_matricula"],
        "escolas_pendentes": resumo["doacoes"]["impacto_social"]["escolas_pendentes"],
        "municipios_impacto_pendente": resumo["doacoes"]["impacto_social"]["municipios_pendentes"]}
with open("data/meta.json", "w", encoding="utf-8") as fh:
    json.dump(meta, fh, ensure_ascii=False)
print("OK", {"doados": resumo["doacoes"]["computadores"], "alunos": resumo["formacao"]["alunos"]})
print("Impacto social:", json.dumps(resumo["doacoes"]["impacto_social"], ensure_ascii=False))
print("Meta:", meta)
if sem_ibge:
    print(f"Aviso: {len(sem_ibge)} municipios sem codigo IBGE (data/sem_ibge.csv e data/municipios_sugestoes.csv)")
