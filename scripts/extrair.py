"""
Baixa as tabelas do Power BI publico "Painel da Area Gestora" (Computadores para Inclusao).
Cada arquivo em consultas/*.json vira um CSV em data/<nome>.csv.
Se vierem poucas linhas, ou o total nao bater com o total do proprio BI, NAO grava nada.
"""
import csv, datetime as dt, glob, json, os, sys, time
import requests

CFG = json.load(open("config.json", encoding="utf-8"))
URL, PAGINA = CFG["url"], int(CFG.get("page_size", 500))
# Protecao contra resposta quebrada. Valores baixos de proposito: so barram "quase nada".
MIN_LINHAS = {"doacoes": 500, "formacao": 100}
HEADERS = {"Content-Type": "application/json;charset=UTF-8", "Accept": "application/json, text/plain, */*",
           "X-PowerBI-ResourceKey": CFG["resource_key"], "User-Agent": "Mozilla/5.0 (compatible; atualizador-dados)"}
EPOCA, TIPO_DATA = dt.datetime(1970, 1, 1), 7


def nome_coluna(s, i):
    for k in ("Column", "Measure"):
        if k in s:
            return s[k]["Property"]
    if "Aggregation" in s:
        return "Soma_" + s["Aggregation"]["Expression"]["Column"]["Property"]
    return s.get("Name") or f"coluna_{i}"


def preparar(corpo, restart):
    corpo = json.loads(json.dumps(corpo))
    cmd = corpo["queries"][0]["Query"]["Commands"][0]["SemanticQueryDataShapeCommand"]
    j = cmd["Binding"].setdefault("DataReduction", {}).setdefault("Primary", {}).setdefault("Window", {})
    j["Count"] = PAGINA
    j.pop("RestartTokens", None)
    if restart:
        j["RestartTokens"] = restart
    return corpo


def pedir(corpo, tentativas=4):
    ultimo = None
    for t in range(tentativas):
        try:
            r = requests.post(URL, data=json.dumps(corpo), headers=HEADERS, timeout=120)
            if r.status_code == 200:
                return r.json()
            ultimo = f"HTTP {r.status_code}: {r.text[:400]}"
            if r.status_code in (400, 401, 403, 404):
                break
        except Exception as e:
            ultimo = repr(e)
        time.sleep(3 * (t + 1))
    raise RuntimeError(ultimo)


def ler_resposta(j):
    res = j["results"][0]["result"]["data"]["dsr"]
    if "DataShapes" in res:
        raise RuntimeError("Erro do BI: " + json.dumps(res["DataShapes"])[:500])
    ds = res["DS"][0]
    if "odata.error" in ds:
        raise RuntimeError("Erro do BI: " + json.dumps(ds["odata.error"])[:500])
    return ds


def decodificar(ds, n):
    """Formato compacto do Power BI -> (linhas, tipos, total). A linha de total (so agregados) vai para 'total'."""
    dicts = ds.get("ValueDicts", {})
    linhas, tipos, total = [], [None] * n, None
    for ph in ds.get("PH", []):
        for dm in ph.values():
            esq, ant = None, [None] * n
            for row in dm:
                if "S" in row:
                    esq = row["S"]
                    if not str(esq[0].get("N", "")).startswith(("A", "M")):
                        tipos = [(esq[i].get("T") if i < len(esq) else None) for i in range(n)]
                if "C" not in row and esq and str(esq[0].get("N", "")).startswith(("A", "M")):
                    total = [row.get(s["N"]) for s in esq]      # linha de total do BI
                    continue
                c, rep, nulo, atual, k = list(row.get("C", [])), row.get("R", 0), row.get("\u00d8", 0), [], 0
                for i in range(n):
                    if rep & (1 << i):
                        v = ant[i]
                    elif nulo & (1 << i):
                        v = None
                    else:
                        v = c[k] if k < len(c) else None
                        k += 1
                        col = esq[i] if esq and i < len(esq) else {}
                        if "DN" in col and isinstance(v, int) and not isinstance(v, bool):
                            d = dicts.get(col["DN"], [])
                            if 0 <= v < len(d):
                                v = d[v]
                    atual.append(v)
                ant = atual
                linhas.append(atual)
    return linhas, tipos, total


def texto(v, tipo):
    if v is None:
        return ""
    if tipo == TIPO_DATA and isinstance(v, (int, float)):
        return (EPOCA + dt.timedelta(milliseconds=v)).strftime("%Y-%m-%d")
    if tipo == TIPO_DATA and isinstance(v, str):
        return v[:10]
    return str(v)


def validar(linhas):
    if not linhas:
        raise RuntimeError("O BI devolveu 0 linhas na primeira pagina. Confira a consulta.")
    if sum(1 for l in linhas if all(v is None for v in l)) > len(linhas) * 0.5:
        raise RuntimeError("Mais da metade das linhas veio vazia: decodificacao desalinhada.")


def baixar(corpo, n):
    linhas, restart, vistos, tipos, total = [], None, set(), [None] * n, None
    for pg in range(1, 1000):
        ds = ler_resposta(pedir(preparar(corpo, restart)))
        novas, t, tot = decodificar(ds, n)
        if any(t):
            tipos = t
        if tot is not None and total is None:
            total = tot
        if pg == 1:
            print("  amostra:", novas[:2], "| total do BI:", total, flush=True)
            validar(novas)
        linhas.extend(novas)
        print(f"  pagina {pg}: +{len(novas)} (total {len(linhas)})", flush=True)
        rt = ds.get("RT")
        if not rt or not novas or (ds.get("IC") and len(novas) < PAGINA):
            break
        k = json.dumps(rt)
        if k in vistos:
            print("  aviso: paginacao repetiu; encerrando")
            break
        vistos.add(k)
        restart = rt
    return linhas, tipos, total


def main():
    os.makedirs("data", exist_ok=True)
    falhas = []
    for arq in sorted(glob.glob("consultas/*.json")):
        nome = os.path.splitext(os.path.basename(arq))[0]
        try:
            corpo = json.load(open(arq, encoding="utf-8"))
            sel = corpo["queries"][0]["Query"]["Commands"][0]["SemanticQueryDataShapeCommand"]["Query"]["Select"]
            cols = [nome_coluna(s, i) for i, s in enumerate(sel)]
            print(f"== {nome}: {cols}", flush=True)
            linhas, tipos, total = baixar(corpo, len(cols))
            print(f"  linhas lidas: {len(linhas)}", flush=True)
            if len(linhas) < MIN_LINHAS.get(nome, 1):
                raise RuntimeError(f"Poucas linhas ({len(linhas)} < {MIN_LINHAS.get(nome, 1)}). Nada foi gravado.")
            # conferencia com o total do proprio BI (primeira coluna de soma)
            if total and isinstance(total[0], (int, float)):
                idx = next((i for i, c in enumerate(cols) if c.startswith("Soma_")), None)
                if idx is not None:
                    soma = sum(l[idx] or 0 for l in linhas)
                    print(f"  conferencia: soma das linhas = {soma} | total do BI = {total[0]}", flush=True)
                    if abs(soma - total[0]) > max(1, 0.002 * total[0]):
                        raise RuntimeError(f"A soma das linhas ({soma}) nao bate com o total do BI ({total[0]}). Nada foi gravado.")
            tmp = f"data/{nome}.csv.tmp"
            with open(tmp, "w", newline="", encoding="utf-8") as f:
                w = csv.writer(f)
                w.writerow(cols)
                for l in linhas:
                    w.writerow([texto(v, tp) for v, tp in zip(l, tipos)])
            os.replace(tmp, f"data/{nome}.csv")
            print(f"  OK: {len(linhas)} linhas em data/{nome}.csv", flush=True)
        except Exception as e:
            print(f"  FALHOU [{nome}]: {e}", flush=True)
            falhas.append(nome)
    if falhas:
        sys.exit("Tabelas com falha: " + ", ".join(falhas) + ". Veja as linhas FALHOU acima.")


if __name__ == "__main__":
    main()
