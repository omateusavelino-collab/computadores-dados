"""
Busca no QEdu (qedu.org.br) as MATRICULAS de cada escola, pelo codigo INEP.  -- VERSAO RAPIDA --
Fonte definida pela area gestora: "para instituicoes de ensino consideramos o numero de matriculas pelo site do QEdu.br".

Entrada : data/escolas_para_consulta.csv  (gerado por scripts/gerar_json.py a partir da LISTA_ESCOLAS.xlsx)
Saida   : data/fixos/matriculas_qedu.csv  (cache: cada escola e consultada uma vez e revisada 1x por ano)

O QUE DEIXA MAIS RAPIDO (em relacao a versao anterior):
  - consulta PARALELO escolas ao mesmo tempo (padrao 3 abas);
  - nao baixa imagens, fontes, videos, estilos nem scripts de propaganda/analytics;
  - em vez de esperar 15 s fixos, le a pagina a cada 0,5 s e para assim que o numero aparece (maximo ESPERA_MAX s);
  - abre direto o endereco /escola/<INEP> (o QEdu completa o endereco sozinho).

SEGURANCA (igual a versao anterior: nunca grava matricula errada):
  1. A pagina precisa mostrar "Codigo INEP: <codigo>" igual ao da lista.
  2. O numero e o do bloco "Situacao da escola" imediatamente antes de "Matriculas".
  3. O municipio mostrado no QEdu precisa ser o da lista (quando o QEdu mostra o municipio).
  4. Guarda o ano do Censo e o link. Se algo falhar, grava o status do problema e NAO grava numero.

Uso:  python scripts/buscar_matriculas_qedu.py            -> consulta ate LIMITE escolas
      python scripts/buscar_matriculas_qedu.py --contar   -> so conta a fila (e escreve pendentes= em GITHUB_OUTPUT)
Variaveis: LIMITE (2000), TEMPO_MAX_MIN (140), PARALELO (3), ESPERA_MAX (10)
"""
import asyncio, csv, datetime as dt, os, random, re, sys, time, unicodedata

ENTRADA = "data/escolas_para_consulta.csv"
CACHE = "data/fixos/matriculas_qedu.csv"
CAMPOS = ["inep", "status", "matriculas", "ano_censo", "nome_qedu", "municipio_qedu", "url", "consultado_em", "obs"]
LIMITE = int(os.environ.get("LIMITE") or 2000)
TEMPO_MAX_MIN = float(os.environ.get("TEMPO_MAX_MIN") or 140)
PARALELO = max(1, min(3, int(os.environ.get("PARALELO") or 1)))
ESPERA_MAX = float(os.environ.get("ESPERA_MAX") or 10)
REVISAR_OK_DIAS = 330         # escola com matricula: revisa 1x por ano (novo Censo)
REVISAR_ERRO_DIAS = 0         # erro de acesso nao vira cache: volta para a fila na proxima execucao
REVISAR_DUVIDA_DIAS = 7       # QEdu nao achou / INEP nao confere / municipio diferente: tenta de novo em 7 dias
HOJE = dt.date.today()
BLOQUEAR_TIPOS = {"image", "media", "font", "stylesheet"}
BLOQUEAR_HOSTS = ("google", "doubleclick", "facebook", "hotjar", "clarity", "tiktok", "analytics", "gtm", "youtube")


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
        st = c.get("status")
        limite = REVISAR_OK_DIAS if st == "ok" else REVISAR_ERRO_DIAS if st == "erro_acesso" else REVISAR_DUVIDA_DIAS
        if dias >= limite:
            out.append(e)
    return out


def analisar(inep, texto, municipio_lista):
    t = re.sub(r"[ \t\xa0]+", " ", texto)
    if not re.search(r"INEP:?\s*" + re.escape(inep) + r"\b", t):
        outro = re.search(r"INEP:?\s*(\d{8})", t)
        if outro:   # a pagina e de OUTRA escola
            return {"status": "inep_nao_confere", "obs": f"a pagina mostra o INEP {outro.group(1)}"}
        # pagina sem nenhum INEP: bloqueio, verificacao anti-robo ou pagina que nao carregou
        return {"status": "erro_acesso", "obs": "pagina sem codigo INEP (bloqueio ou nao carregou): " + t[:150].replace("\n", " ")}
    m = (re.search(r"Situa\S{1,4}o da escola\s*([\d\.]+)\s*Matr\S{1,4}culas", t, re.I)
         or re.search(r"([\d\.]+)\s*Matr\S{1,4}culas\s*[\d\.]+\s*Reprova", t, re.I))
    if not m:
        return {"status": "sem_matricula_no_qedu", "obs": "codigo confere, mas a pagina nao mostra Matriculas"}
    matr = m.group(1).replace(".", "")
    if not matr.isdigit():
        return {"status": "sem_matricula_no_qedu", "obs": f"valor nao numerico: {m.group(1)}"}
    anos_dep = re.findall(r"Fonte:\s*Censo Escolar\s*(\d{4})", t[m.end():m.end() + 400])
    anos_ant = re.findall(r"Fonte:\s*Censo Escolar\s*(\d{4})", t[max(0, m.start() - 600):m.start()])
    ano = (anos_dep[:1] or anos_ant[-1:] or [""])[0]
    mm = re.search(r"Escola (?:Municipal|Estadual|Federal|Privada)[^\n]*?\n?\s*([^\n/]{2,40}?)\s*/\s*([^\n]{2,60})", t)
    mun_qedu = mm.group(2).strip() if mm else ""
    if mun_qedu and municipio_lista and norm(municipio_lista) not in norm(mun_qedu) and norm(mun_qedu) not in norm(municipio_lista):
        return {"status": "municipio_diferente", "municipio_qedu": mun_qedu[:60], "ano_censo": ano,
                "obs": f"lista diz {municipio_lista}; QEdu diz {mun_qedu[:60]}"}
    return {"status": "ok", "matriculas": matr, "ano_censo": ano, "municipio_qedu": mun_qedu[:60]}


_diag = {"n": 0}


def diagnostico(inep, url, status_http, texto):
    """Guarda o que o QEdu devolveu nas 5 primeiras falhas (data/diagnostico_qedu.txt) e mostra no log."""
    if _diag["n"] >= 5:
        return
    _diag["n"] += 1
    trecho = re.sub(r"\s+", " ", texto or "")[:600]
    print(f"    [diagnostico] INEP {inep} | HTTP {status_http} | {url}\n    [diagnostico] texto: {trecho[:300]}", flush=True)
    with open("data/diagnostico_qedu.txt", "a", encoding="utf-8") as fh:
        fh.write(f"{dt.datetime.now():%Y-%m-%d %H:%M} INEP {inep} HTTP {status_http} {url}\n{trecho}\n\n")


async def bloquear(route):
    req = route.request
    if req.resource_type in BLOQUEAR_TIPOS or any(h in req.url for h in BLOQUEAR_HOSTS):
        await route.abort()
    else:
        await route.fallback()


async def ler_quando_pronto(page, inep):
    """Le o texto da pagina a cada 0,5 s ate aparecerem o INEP e 'Matriculas' (ou estourar ESPERA_MAX)."""
    fim = time.monotonic() + ESPERA_MAX
    texto = ""
    while True:
        try:
            texto = await page.inner_text("body", timeout=3000)
        except Exception:
            texto = texto or ""
        if inep in texto and re.search(r"\d\s*Matr\S{1,4}culas", texto):
            await asyncio.sleep(0.3)                     # deixa o numero terminar de renderizar
            try:
                texto = await page.inner_text("body", timeout=3000)
            except Exception:
                pass
            return texto
        if time.monotonic() > fim:
            return texto
        await asyncio.sleep(0.5)


async def consultar(page, e):
    """Consulta uma escola sem tentar contornar bloqueios do site.

    Em 403/429/503 devolve erro_acesso para a execucao aplicar backoff global.
    O resultado nunca recebe matricula se pagina/INEP/municipio nao forem validados.
    """
    ultimo = None
    urls = (f"https://qedu.org.br/escola/{e['inep']}",
            f"https://qedu.org.br/escola/{e['inep']}-{slug(e['nome'])}")
    for url in urls:
        try:
            resp = await page.goto(url, wait_until="domcontentloaded", timeout=30000)
            status_http = resp.status if resp is not None else None
            if status_http == 404:
                return {"status": "nao_encontrado", "obs": "QEdu respondeu 404", "url": page.url}
            if status_http in (403, 429, 503):
                return {"status": "erro_acesso", "obs": f"HTTP {status_http} (bloqueio/limite temporario)",
                        "url": page.url, "http_status": str(status_http)}

            texto = await ler_quando_pronto(page, e["inep"])
            res = analisar(e["inep"], texto, e.get("municipio", ""))
            if res["status"] != "ok":
                diagnostico(e["inep"], page.url, status_http, texto)
            res["url"] = page.url
            try:
                res["nome_qedu"] = (await page.title()).split(" | ")[0].strip()[:120]
            except Exception:
                pass
            if "/escola/" not in page.url and res["status"] != "ok":
                return {"status": "nao_encontrado", "obs": f"QEdu redirecionou para {page.url[:80]}", "url": page.url}
            if res["status"] == "inep_nao_confere":
                return res
            if res["status"] == "erro_acesso":
                ultimo = res
                break
            return res
        except Exception as ex:
            ultimo = {"status": "erro_acesso", "obs": repr(ex)[:200], "url": url}
            break
    return ultimo or {"status": "erro_acesso", "obs": "sem resposta"}


async def rodar(pend, cache):
    from playwright.async_api import async_playwright
    inicio = time.time()
    estado = {"feitos": 0, "bloqueios_seguidos": 0, "parar": False}
    fila_q = asyncio.Queue()
    for e in pend[:LIMITE]:
        fila_q.put_nowait(e)
    trava = asyncio.Lock()

    async def trabalhador(ctx, n):
        page = await ctx.new_page()
        await page.route("**/*", bloquear)
        while not estado["parar"]:
            try:
                e = fila_q.get_nowait()
            except asyncio.QueueEmpty:
                break
            if (time.time() - inicio) / 60 > TEMPO_MAX_MIN:
                estado["parar"] = True
                print("Tempo maximo atingido: salvando e parando (o resto fica para a proxima execucao).", flush=True)
                break

            res = await consultar(page, e)
            bloqueado = res.get("status") == "erro_acesso" and any(x in res.get("obs", "") for x in ("HTTP 403", "HTTP 429", "HTTP 503"))
            res.update({"inep": e["inep"], "consultado_em": HOJE.isoformat()})
            if res.get("status") != "ok":
                res["matriculas"] = ""

            async with trava:
                estado["feitos"] += 1
                if bloqueado:
                    estado["bloqueios_seguidos"] += 1
                    # Nao grava bloqueio temporario no cache. Assim a escola continua na fila.
                    cache.pop(e["inep"], None)
                else:
                    cache[e["inep"]] = {k: v for k, v in res.items() if k in CAMPOS}
                    estado["bloqueios_seguidos"] = 0

                f = estado["feitos"]
                ritmo = f / max(1e-6, (time.time() - inicio) / 60)
                print(f"  {f:>4} {e['inep']} {res['status']:<22} {res.get('matriculas', ''):>6} {res.get('ano_censo', '')} "
                      f"{e['nome'][:45]}  [{ritmo:.1f}/min]", flush=True)
                if f % 25 == 0:
                    salvar_cache(cache)
                if estado["bloqueios_seguidos"] >= 3:
                    estado["parar"] = True
                    print("3 respostas consecutivas de bloqueio/limite do QEdu. Parando cedo para nao insistir; "
                          "as escolas permanecem na fila para a proxima execucao.", flush=True)

            # Uma unica aba por padrao e pausa maior reduzem rajadas no servidor.
            await asyncio.sleep(random.uniform(3.0, 6.0) if not bloqueado else random.uniform(20.0, 35.0))
        await page.close()

    async with async_playwright() as p:
        nav = await p.chromium.launch()
        ctx = await nav.new_context(locale="pt-BR", user_agent=(
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36"))
        await asyncio.gather(*(trabalhador(ctx, i) for i in range(PARALELO)))
        await nav.close()
    return estado["feitos"], time.time() - inicio


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
    print(f"Consultando ate {min(LIMITE, len(pend))} de {len(pend)} escolas na fila, {PARALELO} por vez...", flush=True)
    try:
        feitos, seg = asyncio.run(rodar(pend, cache))
    finally:
        salvar_cache(cache)                                # salva mesmo se der erro no meio
    ok = sum(1 for r in cache.values() if r.get("status") == "ok")
    print(f"Fim: {feitos} consultadas em {seg/60:.0f} min ({feitos/max(1,seg/60):.1f}/min) | "
          f"com matricula conferida: {ok} | ainda na fila: {len(fila(unicas, cache))}")


if __name__ == "__main__":
    main()
