"""
Busca no QEdu (qedu.org.br) o numero de MATRICULAS de cada escola, pelo codigo INEP.
Fonte definida pela area gestora: "para instituicoes de ensino consideramos o numero de matriculas pelo site do QEdu.br".

Entrada : data/escolas_para_consulta.csv  (gerado por scripts/gerar_json.py a partir da LISTA_ESCOLAS.xlsx)
Saida   : data/fixos/matriculas_qedu.csv  (cache: cada escola e consultada uma vez e revisada 1x por ano)

REGRAS DE SEGURANCA (para nunca gravar matricula errada):
  1. A pagina aberta precisa mostrar "Codigo INEP: <codigo>" igual ao da lista.
  2. O numero e o que aparece no bloco "Situacao da escola" imediatamente antes de "Matriculas".
  3. O municipio mostrado no QEdu precisa ser o mesmo da lista (quando o QEdu mostra o municipio).
  4. Guarda o ano do Censo ("Fonte: Censo Escolar AAAA") e o link consultado.
  Se qualquer regra falhar, grava o status do problema e NAO grava numero (a escola fica pendente).

Uso:  python scripts/buscar_matriculas_qedu.py            -> consulta ate LIMITE escolas (padrao 900)
      python scripts/buscar_matriculas_qedu.py --contar   -> so conta quantas estao na fila
"""
import csv, datetime as dt, os, random, re, sys, time, unicodedata

ENTRADA = "data/escolas_para_consulta.csv"
CACHE = "data/fixos/matriculas_qedu.csv"
CAMPOS = ["inep", "status", "matriculas", "ano_censo", "nome_qedu", "municipio_qedu", "url", "consultado_em", "obs"]
LIMITE = int(os.environ.get("LIMITE") or 900)
TEMPO_MAX_MIN = float(os.environ.get("TEMPO_MAX_MIN") or 140)   # para antes do limite do GitHub e salva
REVISAR_OK_DIAS = 330          # escola com matricula: revisa 1x por ano (novo Censo)
REVISAR_ERRO_DIAS = 7          # escola com problema: tenta de novo depois de 7 dias
HOJE = dt.date.today()


def norm(s):
    s = unicodedata.normalize("NFD", str(s or ""))
    s = "".join(c for c in s if unicodedata.category(c) != "Mn").upper()
    return re.sub(r"[^A-Z0-9]+", "", s)


def slug(s):
    s = unicodedata.normalize("NFD", str(s or ""))
    s = "".join(c for c in s if unicodedata.category(c) != "Mn").lower()
    return re.sub(r"[^a-z0-9]+", "-", s).strip("-")[:80] or "escola"


def ler_csv(caminho):
    if not os.path.exists(caminho):
        return []
    with open(caminho, encoding="utf-8-sig") as fh:
        return list(csv.DictReader(fh))


def salvar_cache(cache):
    tmp = CACHE + ".tmp"
    with open(tmp, "w", newline="", encoding="utf-8") as fh:
        w = csv.DictWriter(fh, fieldnames=CAMPOS)
        w.writeheader()
        for k in sorted(cache):
            w.writerow({c: cache[k].get(c, "") for c in CAMPOS})
    os.replace(tmp, CACHE)


def fila(escolas, cache):
    out = []
    for e in escolas:
        c = cache.get(e["inep"])
        if not c:
            out.append(e)
            continue
        try:
            dias = (HOJE - dt.date.fromisoformat(c.get("consultado_em", "")[:10])).days
        except ValueError:
            dias = 9999
        if (c.get("status") == "ok" and dias >= REVISAR_OK_DIAS) or (c.get("status") != "ok" and dias >= REVISAR_ERRO_DIAS):
            out.append(e)
    return out


def analisar(inep, texto, municipio_lista):
    """Le o texto da pagina do QEdu. Devolve dict com status e, so se tudo conferir, matriculas."""
    t = re.sub(r"[ \t\xa0]+", " ", texto)
    if not re.search(r"INEP:?\s*" + re.escape(inep) + r"\b", t):
        return {"status": "inep_nao_confere", "obs": "a pagina nao mostra o codigo INEP da lista"}
    # (acentos tolerados: \S{1,4} aceita "ç", "ã", "í" mesmo se a pagina vier com codificacao diferente)
    m = (re.search(r"Situa\S{1,4}o da escola\s*([\d\.]+)\s*Matr\S{1,4}culas", t, re.I)
         or re.search(r"([\d\.]+)\s*Matr\S{1,4}culas\s*[\d\.]+\s*Reprova", t, re.I))
    if not m:
        return {"status": "sem_matricula_no_qedu", "obs": "codigo confere, mas a pagina nao mostra Matriculas"}
    matr = m.group(1).replace(".", "")
    if not matr.isdigit():
        return {"status": "sem_matricula_no_qedu", "obs": f"valor nao numerico: {m.group(1)}"}
    # ano do Censo: o "Fonte: Censo Escolar AAAA" mais proximo depois do numero (ou antes, se nao houver)
    anos_dep = re.findall(r"Fonte:\s*Censo Escolar\s*(\d{4})", t[m.end():m.end() + 400])
    anos_ant = re.findall(r"Fonte:\s*Censo Escolar\s*(\d{4})", t[max(0, m.start() - 600):m.start()])
    ano = (anos_dep[:1] or anos_ant[-1:] or [""])[0]
    # municipio mostrado pelo QEdu: linha "Escola Municipal/Estadual/... <Estado> / <Municipio>"
    mm = re.search(r"Escola (?:Municipal|Estadual|Federal|Privada)[^\n]*?\n?\s*([^\n/]{2,40}?)\s*/\s*([^\n]{2,60})", t)
    mun_qedu = mm.group(2).strip() if mm else ""
    if mun_qedu and municipio_lista and norm(municipio_lista) not in norm(mun_qedu) and norm(mun_qedu) not in norm(municipio_lista):
        # o municipio do QEdu pode vir seguido de outras palavras na mesma linha; por isso o "in"
        return {"status": "municipio_diferente", "municipio_qedu": mun_qedu[:60], "ano_censo": ano,
                "obs": f"lista diz {municipio_lista}; QEdu diz {mun_qedu[:60]}"}
    return {"status": "ok", "matriculas": matr, "ano_censo": ano, "municipio_qedu": mun_qedu[:60]}


def consultar(page, e):
    ultimo = None
    for url in (f"https://qedu.org.br/escola/{e['inep']}-{slug(e['nome'])}", f"https://qedu.org.br/escola/{e['inep']}"):
        for tentativa in range(3):
            try:
                resp = page.goto(url, wait_until="domcontentloaded", timeout=45000)
                if resp is not None and resp.status == 404:
                    ultimo = {"status": "nao_encontrado", "obs": "QEdu respondeu 404"}
                    break
                try:
                    page.wait_for_selector("text=/Matr.{1,4}culas/", timeout=15000)
                except Exception:
                    pass
                res = analisar(e["inep"], page.inner_text("body"), e.get("municipio", ""))
                res["url"] = page.url
                try:
                    res["nome_qedu"] = page.title().split(" | ")[0].strip()[:120]
                except Exception:
                    pass
                if res["status"] == "ok" or res["status"] == "municipio_diferente":
                    return res
                ultimo = res
                break
            except Exception as ex:
                ultimo = {"status": "erro_acesso", "obs": repr(ex)[:200]}
                time.sleep(5 * (tentativa + 1))
        if ultimo and ultimo.get("status") not in ("nao_encontrado", "inep_nao_confere"):
            return ultimo
    return ultimo or {"status": "erro_acesso", "obs": "sem resposta"}


def main():
    escolas = [e for e in ler_csv(ENTRADA) if re.fullmatch(r"\d{8}", e.get("inep", ""))]
    vistos, unicas = set(), []
    for e in escolas:
        if e["inep"] not in vistos:
            vistos.add(e["inep"])
            unicas.append(e)
    cache = {r["inep"]: r for r in ler_csv(CACHE)}
    pend = fila(unicas, cache)
    if "--contar" in sys.argv:
        print(f"Escolas na lista: {len(unicas)} | ja com matricula: {sum(1 for r in cache.values() if r.get('status') == 'ok')} | na fila: {len(pend)}")
        if os.environ.get("GITHUB_OUTPUT"):
            with open(os.environ["GITHUB_OUTPUT"], "a") as fh:
                fh.write(f"pendentes={len(pend)}\n")
        return
    if not pend:
        print("Nada a consultar: todas as escolas ja tem matricula conferida no QEdu.")
        return
    from playwright.sync_api import sync_playwright
    inicio, feitos, erros_seguidos = time.time(), 0, 0
    print(f"Consultando ate {min(LIMITE, len(pend))} de {len(pend)} escolas na fila...", flush=True)
    with sync_playwright() as p:
        nav = p.chromium.launch()
        ctx = nav.new_context(locale="pt-BR", user_agent=(
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36"))
        page = ctx.new_page()
        for e in pend[:LIMITE]:
            if (time.time() - inicio) / 60 > TEMPO_MAX_MIN:
                print("Tempo maximo atingido: salvando e parando (o resto fica para a proxima execucao).")
                break
            res = consultar(page, e)
            res.update({"inep": e["inep"], "consultado_em": HOJE.isoformat()})
            if res.get("status") != "ok":
                res["matriculas"] = ""            # nunca grava numero sem conferencia completa
            cache[e["inep"]] = res
            feitos += 1
            erros_seguidos = erros_seguidos + 1 if res["status"] == "erro_acesso" else 0
            print(f"  {feitos:>4} {e['inep']} {res['status']:<22} {res.get('matriculas', ''):>6} {res.get('ano_censo', '')} {e['nome'][:50]}", flush=True)
            if feitos % 25 == 0:
                salvar_cache(cache)
            if erros_seguidos >= 10:
                print("10 erros de acesso seguidos: o QEdu pode estar bloqueando. Salvando e parando.")
                break
            time.sleep(random.uniform(1.5, 3.0))   # pausa educada entre consultas
        nav.close()
    salvar_cache(cache)
    ok = sum(1 for r in cache.values() if r.get("status") == "ok")
    print(f"Fim: {feitos} consultadas nesta execucao | com matricula conferida: {ok} | ainda na fila: {len(fila(unicas, cache))}")


if __name__ == "__main__":
    main()
