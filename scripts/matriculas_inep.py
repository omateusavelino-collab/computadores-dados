"""
Matriculas por escola a partir dos MICRODADOS OFICIAIS do Censo Escolar (INEP) -- fonte da qual o QEdu tambem se alimenta.

Por que existe: o QEdu passou a bloquear (HTTP 403) as consultas feitas pelo GitHub Actions, ja na primeira escola.
Em vez de abrir milhares de paginas, este script baixa UMA vez o arquivo oficial do Censo e le a matricula de cada escola
pelo codigo INEP (CO_ENTIDADE).

Entrada : data/escolas_para_consulta.csv      (gerada por scripts/gerar_json.py)
          data/fixos/matriculas_qedu.csv      (matriculas ja conferidas no QEdu: NUNCA sao alteradas)
Saida   : data/fixos/matriculas_inep.csv      (matriculas vindas do INEP, so para escolas SEM matricula do QEdu)
          data/relatorio_equivalencia_inep.csv (teste de equivalencia INEP x QEdu, para auditoria)

SEGURANCA (nao grava numero que nao seja comprovado):
  1. Descobre qual coluna/soma do INEP reproduz o numero do QEdu comparando com as escolas JA conferidas no QEdu.
  2. So usa essa coluna se bater exatamente em pelo menos EQUIV_MIN (95%) das escolas comparadas (minimo de 50 escolas).
     Se nao bater, NAO grava nada e termina com erro, mostrando o ranking de candidatos.
  3. Codigo INEP precisa existir no arquivo oficial; matricula zero ou ausente nao e gravada (a escola continua pendente).
  4. Guarda ano do Censo, coluna usada e link da fonte em cada linha.

DOWNLOAD: o servidor do INEP (download.inep.gov.br) usa um certificado que o GitHub Actions nao reconhece
("unable to get local issuer certificate"). O script tenta primeiro com verificacao normal; se falhar por SSL, repete
sem verificar o certificado, SOMENTE para este download de dado publico. A integridade e garantida pelo ZIP valido +
teste de equivalencia com o QEdu.

Variaveis: ANO_CENSO (2025), INEP_ZIP_URL (opcional: link direto do .zip), INEP_ZIP_ARQUIVO (opcional: caminho de um .zip local),
           EQUIV_MIN (0.95)
"""
import csv, datetime as dt, itertools, os, re, sys, time, zipfile
import requests
import urllib3
import numpy as np
import pandas as pd

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

ENTRADA = "data/escolas_para_consulta.csv"
CACHE_QEDU = "data/fixos/matriculas_qedu.csv"
SAIDA = "data/fixos/matriculas_inep.csv"
RELATORIO = "data/relatorio_equivalencia_inep.csv"
DIAGNOSTICO = "data/diagnostico_equivalencia.csv"
CAMPOS = ["inep", "status", "matriculas", "ano_censo", "fonte", "coluna_usada", "url", "consultado_em", "obs"]
ANO = int(os.environ.get("ANO_CENSO") or 2025)
EQUIV_MIN = float(os.environ.get("EQUIV_MIN") or 0.95)
MIN_COMPARADAS = 50
HOJE = dt.date.today()
PAGINA_INEP = "https://www.gov.br/inep/pt-br/acesso-a-informacao/dados-abertos/microdados/censo-escolar"
UA = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36"}


def ler_csv(caminho):
    if not os.path.exists(caminho):
        return []
    with open(caminho, encoding="utf-8-sig") as fh:
        return list(csv.DictReader(fh))


# ------------------------------------------------------------------ download
def urls_candidatas():
    urls = []
    if os.environ.get("INEP_ZIP_URL", "").strip():
        urls.append(os.environ["INEP_ZIP_URL"].strip())
    try:  # tenta achar o link oficial na pagina do INEP
        html = requests.get(PAGINA_INEP, headers=UA, timeout=60).text
        for u in re.findall(r'https?://download\.inep\.gov\.br/[^\s"\'<>]+\.zip', html, flags=re.I):
            if str(ANO) in u and "censo_escolar" in u.lower():
                urls.append(u)
    except Exception as ex:
        print(f"Aviso: nao consegui ler a pagina do INEP ({str(ex)[:100]}); uso os enderecos padrao.")
    urls += [f"https://download.inep.gov.br/dados_abertos/microdados_censo_escolar_{ANO}.zip",
             f"https://download.inep.gov.br/dados_abertos/microdados_censo_escolar_{ANO}_v2.zip"]
    vistos, out = set(), []
    for u in urls:
        if u not in vistos:
            vistos.add(u)
            out.append(u)
    return out


def _baixar_uma(url, destino, verificar_ssl):
    """Devolve None se baixou um ZIP valido; senao, texto do problema."""
    with requests.get(url, headers=UA, stream=True, timeout=(30, 180), verify=verificar_ssl) as r:
        if r.status_code != 200:
            return f"HTTP {r.status_code}"
        total = 0
        with open(destino, "wb") as fh:
            for bloco in r.iter_content(1 << 20):
                fh.write(bloco)
                total += len(bloco)
    if total < 1_000_000 or not zipfile.is_zipfile(destino):
        return "arquivo recebido nao e um ZIP valido"
    print(f"  OK: {total/1e6:.0f} MB", flush=True)
    return None


def baixar():
    local = os.environ.get("INEP_ZIP_ARQUIVO")
    if local:
        print(f"Usando arquivo local: {local}")
        return local, "arquivo local"
    os.makedirs("work", exist_ok=True)
    destino = f"work/censo_escolar_{ANO}.zip"
    for url in urls_candidatas():
        verificar = True
        tentativa = 0
        while tentativa < 4:
            tentativa += 1
            modo = "" if verificar else " [SSL sem verificacao]"
            print(f"Baixando {url}{modo} (tentativa {tentativa}) ...", flush=True)
            try:
                erro = _baixar_uma(url, destino, verificar)
                if erro is None:
                    return destino, url
                print(f"  {erro}")
                break                                  # servidor respondeu algo ruim (404...): proximo endereco
            except requests.exceptions.SSLError as ex:
                print(f"  certificado SSL nao reconhecido: {str(ex)[:90]}")
                if verificar:
                    verificar = False                   # repete sem verificar, so para este dado publico
                    continue
                break
            except Exception as ex:
                print(f"  conexao falhou: {str(ex)[:110]}")
                time.sleep(5 * tentativa)               # conexao cortada: espera e tenta de novo
    sys.exit("ERRO: nao consegui baixar os microdados do INEP pelo GitHub Actions. Baixe o ZIP no seu computador, "
             "publique-o em uma Release do repositorio e informe o link no campo 'url_zip'.")


# ------------------------------------------------------------------ leitura
def detectar_formato(zf, nome):
    """Descobre codificacao (utf-8 ou latin-1) e separador (; ou ,) olhando o comeco do arquivo."""
    with zf.open(nome) as fh:
        topo = fh.read(4_000_000)
    enc = "utf-8-sig"
    try:
        topo.decode("utf-8")
    except UnicodeDecodeError as ex:
        if ex.start < len(topo) - 4:                 # erro no meio do texto (nao e so o corte do bloco)
            enc = "latin-1"
    linha = topo.split(b"\n", 1)[0].decode("latin-1")
    sep = ";" if linha.count(";") >= linha.count(",") else ","
    cols = [c.strip().strip('"').upper() for c in linha.split(sep)]
    return enc, sep, cols


def abrir_tabela(zf, nome, colunas_pred):
    """Le so as colunas pedidas (arquivo inteiro)."""
    enc, sep, cols = detectar_formato(zf, nome)
    print(f"  formato: codificacao={enc}, separador='{sep}', {len(cols)} colunas")
    with zf.open(nome) as fh:
        cab = pd.read_csv(fh, sep=sep, encoding=enc, encoding_errors="replace", nrows=0, dtype=str)
    usar = [c for c in cab.columns if colunas_pred(c.strip().upper())]
    with zf.open(nome) as fh:
        df = pd.read_csv(fh, sep=sep, encoding=enc, encoding_errors="replace", usecols=usar, dtype=str, low_memory=False)
    df.columns = [c.strip().upper() for c in df.columns]
    return df, cols


def tabela_escola(zf):
    nomes = [n for n in zf.namelist() if n.lower().endswith(".csv")]
    alvo = [n for n in nomes if re.search(r"tabela_escola|microdados_ed_basica", n, re.I)
            and not re.search(r"gestor|docente|turma|matricula|curso", n, re.I)]
    if not alvo:
        sys.exit("ERRO: nao achei a tabela de escolas no ZIP. Arquivos: " + ", ".join(nomes[:20]))
    nome = sorted(alvo, key=len)[0]
    print(f"Tabela de escolas: {nome}")
    df, cols = abrir_tabela(zf, nome, lambda c: c == "CO_ENTIDADE" or c.startswith("QT_MAT"))
    return df, cols, nome, nomes


def candidatas_de_escola(df):
    num = {c: pd.to_numeric(df[c], errors="coerce") for c in df.columns if c.startswith("QT_MAT")}
    cand = dict(num)

    def soma(nome, partes, menos=()):
        if all(p in num for p in list(partes) + list(menos)):
            cand[nome] = sum(num[p].fillna(0) for p in partes) - sum(num[m].fillna(0) for m in menos)

    soma("QT_MAT_INF+QT_MAT_FUND+QT_MAT_MED", ["QT_MAT_INF", "QT_MAT_FUND", "QT_MAT_MED"])
    soma("QT_MAT_INF+QT_MAT_FUND+QT_MAT_MED+QT_MAT_EJA", ["QT_MAT_INF", "QT_MAT_FUND", "QT_MAT_MED", "QT_MAT_EJA"])
    soma("QT_MAT_BAS-QT_MAT_EJA", ["QT_MAT_BAS"], ["QT_MAT_EJA"])
    soma("QT_MAT_BAS-QT_MAT_ESP", ["QT_MAT_BAS"], ["QT_MAT_ESP"])
    soma("QT_MAT_BAS-QT_MAT_PROF", ["QT_MAT_BAS"], ["QT_MAT_PROF"])
    soma("QT_MAT_INF+QT_MAT_FUND+QT_MAT_MED+QT_MAT_EJA+QT_MAT_PROF", ["QT_MAT_INF", "QT_MAT_FUND", "QT_MAT_MED", "QT_MAT_EJA", "QT_MAT_PROF"])
    return cand


def candidata_tabela_matricula(zf, nomes, ineps_lista):
    """Plano B (layout novo com tabela de matriculas separada): le em BLOCOS (arquivo grande) e conta linhas / soma colunas QT_."""
    alvo = [n for n in nomes if re.search(r"tabela_matricula", n, re.I)]
    if not alvo:
        return {}
    nome = alvo[0]
    print(f"Tabela de matriculas (plano B): {nome}")
    enc, sep, cols = detectar_formato(zf, nome)
    print(f"  formato: codificacao={enc}, separador='{sep}', {len(cols)} colunas: {', '.join(cols[:40])}")
    if "CO_ENTIDADE" not in cols:
        return {}
    qt = [c for c in cols if c.startswith("QT_")]
    usar = ["CO_ENTIDADE"] + qt
    linhas, somas = None, {c: None for c in qt}
    with zf.open(nome) as fh:
        for bloco in pd.read_csv(fh, sep=sep, encoding=enc, encoding_errors="replace", dtype=str,
                                 usecols=lambda c: c.strip().upper() in usar, chunksize=500_000):
            bloco.columns = [c.strip().upper() for c in bloco.columns]
            chave = bloco["CO_ENTIDADE"].str.strip().str.zfill(8)
            filtro = chave.isin(ineps_lista)          # so as escolas da nossa lista
            bloco, chave = bloco[filtro], chave[filtro]
            if bloco.empty:
                continue
            n = chave.value_counts()
            linhas = n if linhas is None else linhas.add(n, fill_value=0)
            for c in qt:
                v = pd.to_numeric(bloco[c], errors="coerce").fillna(0).groupby(chave).sum()
                somas[c] = v if somas[c] is None else somas[c].add(v, fill_value=0)
    out = {"LINHAS(Tabela_Matricula)": linhas}
    for c in qt:
        out[f"SOMA({c})"] = somas[c]

    def comb(nome, partes):
        if all(f"SOMA(QT_MAT_{x})" in out for x in partes):
            out[nome] = sum(out[f"SOMA(QT_MAT_{x})"].fillna(0) for x in partes)
    # Regra que reproduziu o QEdu no diagnostico: Ensino Fundamental + Ensino Medio (sem infantil, EJA, etc.)
    comb("QT_MAT_FUND+QT_MAT_MED", ["FUND", "MED"])
    comb("QT_MAT_INF+QT_MAT_FUND+QT_MAT_MED", ["INF", "FUND", "MED"])
    comb("QT_MAT_FUND+QT_MAT_MED+QT_MAT_EJA", ["FUND", "MED", "EJA"])
    return out


def _fmt_formula(nomes, sinais):
    return " ".join(("+" if sg > 0 else "-") + n.replace("SOMA(", "").rstrip(")") for n, sg in zip(nomes, sinais)).lstrip("+")


def buscar_formula(cand, base, topo=28):
    """Procura soma/subtracao de ate 4 colunas QT_MAT* que reproduza o numero do QEdu nas escolas ja conferidas."""
    nomes_cols = [n for n in cand if n.startswith("SOMA(QT_MAT")]
    if not nomes_cols:
        return {}
    W = pd.DataFrame({n: cand[n] for n in nomes_cols}).fillna(0)
    W.index = W.index.astype(str).str.strip().str.zfill(8)
    ks = [i for i in base if i in W.index]
    if len(ks) < MIN_COMPARADAS:
        return {}
    y = np.array([base[i] for i in ks], dtype=float)
    X = W.loc[ks].to_numpy(dtype=float)
    corr = np.array([abs(np.corrcoef(X[:, j], y)[0, 1]) if X[:, j].std() > 0 else 0.0 for j in range(X.shape[1])])
    idx = list(np.argsort(-corr)[:topo])
    for c in ("BAS", "BAS_D", "INF", "FUND", "MED", "EJA", "PROF", "ESP"):   # colunas principais sempre entram
        n = f"SOMA(QT_MAT_{c})"
        if n in nomes_cols and nomes_cols.index(n) not in idx:
            idx.append(nomes_cols.index(n))
    achados = []
    for k in (1, 2, 3, 4):
        pool = idx if k <= 3 else idx[:14] + [j for j in idx[14:] if nomes_cols[j] in ('SOMA(QT_MAT_FUND)','SOMA(QT_MAT_MED)','SOMA(QT_MAT_INF)','SOMA(QT_MAT_EJA)','SOMA(QT_MAT_PROF)')]
        for comb in itertools.combinations(pool, k):
            sub = X[:, comb]
            for sinais in itertools.product((1, -1), repeat=k):
                if sinais[0] == -1:
                    continue
                v = sub @ np.array(sinais, dtype=float)
                ig = int((v == y).sum())
                achados.append((ig, comb, sinais))
    achados.sort(key=lambda t: -t[0])
    out = {}
    for ig, comb, sinais in achados[:10]:
        nome = "FORMULA: " + _fmt_formula([nomes_cols[j] for j in comb], sinais)
        out[nome] = sum(sg * W[nomes_cols[j]] for j, sg in zip(comb, sinais))
    return out


def gravar_diagnostico(cand, base):
    """Tabela lado a lado (QEdu x colunas principais do INEP) para as escolas ja conferidas, para conferir a olho."""
    chaves = ["BAS", "BAS_D", "BAS_N", "BAS_EAD", "INF", "FUND", "MED", "EJA", "PROF", "ESP"]
    cols = [(c, cand.get(f"SOMA(QT_MAT_{c})")) for c in chaves]
    cols = [(c, sr) for c, sr in cols if sr is not None]
    if not cols:
        return
    os.makedirs(os.path.dirname(DIAGNOSTICO), exist_ok=True)
    with open(DIAGNOSTICO, "w", newline="", encoding="utf-8") as fh:
        w = csv.writer(fh)
        w.writerow(["inep", "matriculas_qedu"] + [f"INEP_{c}" for c, _ in cols] + ["qedu_menos_BAS_D"])
        for i in sorted(base)[:150]:
            linha = []
            for _, sr in cols:
                sr.index = sr.index.astype(str).str.strip().str.zfill(8)
                linha.append(int(sr[i]) if i in sr.index and pd.notna(sr[i]) else "")
            bd = dict((c, v) for (c, _), v in zip(cols, linha)).get("BAS_D", "")
            w.writerow([i, base[i]] + linha + [base[i] - bd if bd != "" else ""])
    print(f"  diagnostico gravado em {DIAGNOSTICO}")


# ------------------------------------------------------------------ principal
def main():
    escolas = [e for e in ler_csv(ENTRADA) if re.fullmatch(r"\d{8}", e.get("inep", ""))]
    ineps_lista = sorted({e["inep"] for e in escolas})
    qedu = {r["inep"]: r for r in ler_csv(CACHE_QEDU) if r.get("status") == "ok" and (r.get("matriculas") or "").isdigit()}
    ja_inep = {r["inep"]: r for r in ler_csv(SAIDA) if r.get("status") == "ok"}
    pend = [i for i in ineps_lista if i not in qedu and i not in ja_inep]
    print(f"Escolas na lista: {len(ineps_lista)} | com matricula do QEdu: {len(qedu)} | com matricula do INEP: {len(ja_inep)} | pendentes: {len(pend)}")
    if not pend:
        print("Nada a fazer: todas as escolas ja tem matricula.")
        return

    caminho, origem = baixar()
    base = {i: int(r["matriculas"]) for i, r in qedu.items() if str(r.get("ano_censo")) == str(ANO)}

    def avaliar(cand):
        ranking = []
        for nome, serie in cand.items():
            serie = serie.copy()
            serie.index = serie.index.astype(str).str.strip().str.zfill(8)
            comparadas = iguais = 0
            for i, v in base.items():
                if i in serie.index and pd.notna(serie[i]):
                    comparadas += 1
                    iguais += int(round(float(serie[i]))) == v
            ranking.append((iguais / comparadas if comparadas else 0.0, iguais, comparadas, nome, serie))
        ranking.sort(key=lambda x: (-x[0], -x[2]))
        return ranking

    def aprovado(r):
        return bool(r) and r[0][2] >= MIN_COMPARADAS and r[0][0] >= EQUIV_MIN

    ranking = []
    with zipfile.ZipFile(caminho) as zf:
        df, cols, nome_tab, nomes = tabela_escola(zf)
        if "CO_ENTIDADE" not in df.columns:
            sys.exit(f"ERRO: a tabela {nome_tab} nao tem CO_ENTIDADE. Colunas: {cols[:40]}")
        print(f"  colunas de matricula na tabela de escolas: {[c for c in cols if c.startswith('QT_MAT')] or 'nenhuma'}")
        df["CO_ENTIDADE"] = df["CO_ENTIDADE"].str.strip().str.zfill(8)
        df = df.drop_duplicates("CO_ENTIDADE").set_index("CO_ENTIDADE")
        cand = {k: v for k, v in candidatas_de_escola(df).items() if len(v)}
        ranking = avaliar(cand) if cand else []
        if not aprovado(ranking):
            print("Tabela de escolas nao reproduz o QEdu (ou nao tem colunas QT_MAT*); tentando a tabela de matriculas...", flush=True)
            extra = candidata_tabela_matricula(zf, nomes, set(ineps_lista))
            if extra:
                ranking = sorted(ranking + avaliar(extra), key=lambda x: (-x[0], -x[2]))
                gravar_diagnostico(extra, base)
                if not aprovado(ranking):
                    print("Nenhuma coluna isolada reproduz o QEdu; procurando combinacoes (soma/subtracao de ate 4 colunas)...", flush=True)
                    formulas = buscar_formula(extra, base)
                    if formulas:
                        ranking = sorted(ranking + avaliar(formulas), key=lambda x: (-x[0], -x[2]))
    if not ranking:
        sys.exit("ERRO: nao encontrei nenhuma coluna de matriculas no ZIP do INEP. Colunas da tabela de escolas: " + ", ".join(cols[:80]))
    with open(RELATORIO, "w", newline="", encoding="utf-8") as fh:
        w = csv.writer(fh)
        w.writerow(["candidato", "escolas_comparadas", "valores_iguais_ao_qedu", "percentual_igual", "gerado_em", "ano_censo"])
        for pct, ig, comp, nome, _ in ranking[:60]:
            w.writerow([nome, comp, ig, f"{pct:.4f}", HOJE.isoformat(), ANO])
    print("\nTeste de equivalencia INEP x QEdu (melhores candidatos):")
    for pct, ig, comp, nome, _ in ranking[:10]:
        print(f"  {pct*100:6.2f}%  ({ig}/{comp})  {nome}")

    pct, ig, comp, melhor, serie = ranking[0]
    if comp < MIN_COMPARADAS or pct < EQUIV_MIN:
        sys.exit(f"\nERRO: nenhum candidato reproduz o numero do QEdu (melhor: {melhor} = {pct*100:.1f}% em {comp} escolas; "
                 f"minimo exigido {EQUIV_MIN*100:.0f}% em >= {MIN_COMPARADAS}). NADA FOI GRAVADO. Veja {RELATORIO} e {DIAGNOSTICO}.")
    print(f"\nCandidato escolhido: {melhor} ({pct*100:.2f}% igual ao QEdu em {comp} escolas)")

    # ---------- grava so escolas sem matricula do QEdu
    novos, sem_dado = [], 0
    for i in pend:
        v = serie[i] if i in serie.index else None
        if v is None or pd.isna(v) or int(round(float(v))) <= 0:
            sem_dado += 1
            continue
        novos.append({"inep": i, "status": "ok", "matriculas": int(round(float(v))), "ano_censo": ANO,
                      "fonte": "INEP - Censo Escolar (microdados)", "coluna_usada": melhor,
                      "url": origem, "consultado_em": HOJE.isoformat(),
                      "obs": f"equivalencia com QEdu: {pct*100:.1f}% ({ig}/{comp})"})
    todos = dict(ja_inep)
    for n in novos:
        todos[n["inep"]] = n
    os.makedirs(os.path.dirname(SAIDA), exist_ok=True)
    with open(SAIDA, "w", newline="", encoding="utf-8") as fh:
        w = csv.DictWriter(fh, fieldnames=CAMPOS)
        w.writeheader()
        for k in sorted(todos):
            w.writerow({c: todos[k].get(c, "") for c in CAMPOS})
    print(f"Gravadas {len(novos)} escolas novas | sem matricula no INEP (continuam pendentes): {sem_dado}")


if __name__ == "__main__":
    main()
