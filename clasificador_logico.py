"""
Clasificador lógico transparente - ABP Lógica (ISPC)
Aplica las reglas P, Q, R y F = (P y Q) o (P y R) o (Q y R) sobre los datos de estimaciones agrícolas.

Entrada: ABP_Logica_Clasificador.xlsx (hoja "MiniDataset")
Uso:     python clasificador_logico.py [ABP_Logica_Clasificador.xlsx] [resultado.csv]
Requiere: pip install pandas openpyxl
"""
import sys
import pandas as pd

UMBRAL_P = 20.0   # % de superficie no cosechada (criterio 1.1)
RACHA_R = 3       # caídas consecutivas mínimas (criterio 1.3)
HOJA_DATOS = "MiniDataset"
HOJA_EXCEL = "Clasificador"   # solo para comparar resultados

COLUMNAS = {  # encabezado del Excel -> nombre interno
    "Región": "region", "Cultivo": "cultivo", "Campaña": "campania",
    "Sup. sembrada (ha)": "sembrada", "Sup. cosechada (ha)": "cosechada",
    "Producción (tm)": "produccion",
    "Rendimiento regional (kg/ha)": "rend_regional",
    "Rendimiento nacional (kg/ha)": "rend_nacional",
}


def tasa_no_cosechada(sembrada, cosechada):
    """Tasa = (sembrada - cosechada) / sembrada * 100"""
    return (sembrada - cosechada) / sembrada * 100 if sembrada > 0 else 0.0


def regla_p(tasa):
    """P: tasa de superficie no cosechada > 20 %"""
    return tasa > UMBRAL_P


def regla_q(rend_regional, rend_nacional):
    """Q: rendimiento regional < rendimiento nacional (mismo cultivo y campaña)"""
    return rend_regional < rend_nacional


def regla_r(racha):
    """R: 3 o más caídas consecutivas de producción"""
    return racha >= RACHA_R


def foco_critico(p, q, r):
    """F = (P y Q) o (P y R) o (Q y R)"""
    return (p and q) or (p and r) or (q and r)


def incidencia(sup_act, sup_ant, rend_act, rend_ant):
    """Criterio 1.4: qué variable explica la caída de producción."""
    baja_sup, baja_rend = sup_act < sup_ant, rend_act < rend_ant
    if baja_sup and baja_rend:
        return "Ambas (superficie cosechada y rendimiento)"
    if baja_sup:
        return "Superficie cosechada"
    if baja_rend:
        return "Rendimiento"
    return "Indeterminada"


def explicar(p, q, r, racha, f):
    nombres = []
    if p: nombres.append(f"P (superficie no cosechada > {UMBRAL_P:.0f} %)")
    if q: nombres.append("Q (rendimiento regional < nacional)")
    if r: nombres.append(f"R ({racha} caídas consecutivas de producción)")
    vv = lambda b: "V" if b else "F"
    base = f"P={vv(p)}, Q={vv(q)}, R={vv(r)} -> "
    if f:
        return base + "FOCO CRÍTICO: se cumplen " + " y ".join(nombres) + "."
    if not nombres:
        return base + "No es foco crítico: no se cumple ninguna condición desfavorable."
    return base + "No es foco crítico: solo se cumple " + nombres[0] + "."


def cargar_datos(ruta):
    try:
        df = pd.read_excel(ruta, sheet_name=HOJA_DATOS)
    except ValueError:
        raise SystemExit(f'No se encontró la hoja "{HOJA_DATOS}" en {ruta}.')
    df.columns = [str(c).strip() for c in df.columns]
    faltan = [c for c in COLUMNAS if c not in df.columns]
    if faltan:
        raise SystemExit(f"Faltan columnas en la hoja {HOJA_DATOS}: {faltan}\nColumnas encontradas: {list(df.columns)}")
    df = df.rename(columns=COLUMNAS)[list(COLUMNAS.values())].dropna(subset=["region", "cultivo", "campania"])
    for c in ("sembrada", "cosechada", "produccion", "rend_regional", "rend_nacional"):
        df[c] = pd.to_numeric(df[c], errors="coerce")
    return df


def clasificar(df):
    filas = []
    for (cultivo, region), g in df.groupby(["cultivo", "region"], sort=False):
        g = g.sort_values("campania", kind="stable")   # "2015/2016" ordena cronológicamente
        racha, ant = 0, None
        for _, x in g.iterrows():
            # Caída: producción actual < producción de la campaña disponible anterior
            caida = ant is not None and x.produccion < ant.produccion
            racha = racha + 1 if caida else 0
            tasa = tasa_no_cosechada(x.sembrada, x.cosechada)
            p = regla_p(tasa)
            q = regla_q(x.rend_regional, x.rend_nacional)
            r = regla_r(racha)
            f = foco_critico(p, q, r)
            filas.append({
                "region": region, "cultivo": cultivo, "campania": x.campania,
                "tasa_no_cosechada_%": round(tasa, 2),
                "P": p, "Q": q, "caida": caida, "racha_caidas": racha, "R": r,
                "foco_critico": f,
                "explicacion": explicar(p, q, r, racha, f),
                "incidencia_caida": incidencia(x.cosechada, ant.cosechada,
                                               x.rend_regional, ant.rend_regional) if caida else "-",
            })
            ant = x
    return pd.DataFrame(filas)


def comparar_con_excel(res, ruta):
    """Compara P, Q, R y foco crítico del script con los calculados por las fórmulas de Excel."""
    try:
        ex = pd.read_excel(ruta, sheet_name=HOJA_EXCEL)
        ex = ex.rename(columns={"Región": "region", "Cultivo": "cultivo", "Campaña": "campania",
                                "Foco crítico": "F_excel", "P": "P_excel", "Q": "Q_excel", "R": "R_excel"})
        ex = ex[["region", "cultivo", "campania", "P_excel", "Q_excel", "R_excel", "F_excel"]].dropna()
        if ex.empty:
            return
    except Exception:
        return
    m = res.merge(ex, on=["region", "cultivo", "campania"])
    ok = ((m.P == m.P_excel.astype(bool)) & (m.Q == m.Q_excel.astype(bool)) &
          (m.R == m.R_excel.astype(bool)) & (m.foco_critico == m.F_excel.astype(bool)))
    print(f"\nComparación con las fórmulas de Excel (hoja {HOJA_EXCEL}): {int(ok.sum())} de {len(m)} casos coinciden.")


def tabla(filas, encabezados, alinear=None):
    """Imprime una tabla alineada con bordes. filas: lista de listas de texto."""
    alinear = alinear or ["<"] * len(encabezados)
    ancho = [max(len(str(h)), *(len(str(f[i])) for f in filas)) for i, h in enumerate(encabezados)]
    linea = lambda l, m, r: l + m.join("─" * (w + 2) for w in ancho) + r
    fila = lambda vals: "│" + "│".join(f" {str(v):{al}{w}} " for v, w, al in zip(vals, ancho, alinear)) + "│"
    out = [linea("┌", "┬", "┐"), fila(encabezados), linea("├", "┼", "┤")]
    out += [fila(f) for f in filas]
    out.append(linea("└", "┴", "┘"))
    print("\n".join(out))


if __name__ == "__main__":
    entrada = sys.argv[1] if len(sys.argv) > 1 else "ABP_Logica_Clasificador.xlsx"
    salida = sys.argv[2] if len(sys.argv) > 2 else "resultado_clasificador.csv"
    res = clasificar(cargar_datos(entrada))
    vf = lambda b: "V" if b else "F"
    inc_corta = {"Ambas (superficie cosechada y rendimiento)": "Ambas", "Superficie cosechada": "Sup. cosechada"}

    print("\nCLASIFICADOR LÓGICO  |  F = (P y Q) o (P y R) o (Q y R)")
    print("P: no cosechada > 20 %   Q: rend. regional < nacional   R: 3 o más caídas consecutivas\n")

    print("1) RESULTADOS POR CASO")
    tabla([[x.cultivo, x.region, x.campania, f"{x['tasa_no_cosechada_%']:.2f} %", vf(x.P), vf(x.Q),
            x.racha_caidas, vf(x.R), "SI" if x.foco_critico else "NO",
            inc_corta.get(x.incidencia_caida, x.incidencia_caida)] for _, x in res.iterrows()],
          ["Cultivo", "Región", "Campaña", "No cosechada", "P", "Q", "Racha", "R", "Foco crítico", "Incidencia de la caída"],
          ["<", "<", "<", ">", "^", "^", "^", "^", "^", "<"])

    print("\n2) EXPLICACIÓN LÓGICA")
    for _, x in res.iterrows():
        print(f"  {x.cultivo} | {x.region} | {x.campania}")
        print(f"      {x.explicacion}")

    print("\n3) COMBINACIONES P/Q/R")
    res["comb"] = res.apply(lambda x: vf(x.P) + vf(x.Q) + vf(x.R), axis=1)
    cuenta = res.groupby("comb").agg(casos=("comb", "size"), foco=("foco_critico", "max"))
    tabla([[c, r.casos, "SI" if r.foco else "NO"] for c, r in cuenta.iterrows()],
          ["P Q R", "Casos", "Foco crítico"], ["^", "^", "^"])
    print(f"\n  Total: {len(res)} casos | Focos críticos: {int(res.foco_critico.sum())}")

    comparar_con_excel(res, entrada)
    res.drop(columns="comb").to_csv(salida, index=False, encoding="utf-8-sig")
    print(f"\nResultados guardados en {salida}")
