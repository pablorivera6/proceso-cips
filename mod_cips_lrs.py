# mod_cips_lrs.py — sin geopandas ni matplotlib
import os
import numpy as np
import pandas as pd
import shapefile  # pyshp
from pyproj import Transformer
from shapely.geometry import Point, LineString, MultiLineString
from shapely.ops import linemerge
from sklearn.linear_model import LinearRegression


# ── Leer línea desde shapefile (pyshp + pyproj) ───────────────────────────────

def _leer_linea_proyectada(ruta_shp):
    """Lee el shapefile y devuelve una LineString en EPSG:3857 (metros)."""
    sf    = shapefile.Reader(ruta_shp)
    prj   = ruta_shp.replace(".shp", ".prj")
    crs_src = "EPSG:4326"  # default WGS84

    # Intentar leer CRS del .prj
    if os.path.exists(prj):
        txt = open(prj).read().upper()
        if "3857" in txt or "WEB_MERCATOR" in txt or "PSEUDO" in txt:
            crs_src = "EPSG:3857"
        elif "32618" in txt or "UTM" in txt:
            crs_src = "EPSG:32618"
        # Si no reconocemos, asumimos WGS84

    t = Transformer.from_crs(crs_src, "EPSG:3857", always_xy=True)

    lineas = []
    for shape in sf.shapes():
        pts = shape.points
        if not pts:
            continue
        xs, ys = zip(*pts)
        xs_m, ys_m = t.transform(xs, ys)
        lineas.append(LineString(zip(xs_m, ys_m)))

    if not lineas:
        raise ValueError("No se encontraron geometrías en el shapefile.")

    merged = linemerge(lineas)
    if isinstance(merged, MultiLineString):
        coords = []
        for ls in merged.geoms:
            coords.extend(list(ls.coords))
        return LineString(coords)
    return merged


# ── Procesamiento principal ───────────────────────────────────────────────────

def ejecutar_cips_lrs(carpeta, archivo_unificado, shp_path):
    print("=== MOD_CIPS_LRS INICIADO ===")

    CRITERIO_OK       = -850
    CRITERIO_MARGINAL = -1200

    # 1. Carga
    xls    = pd.ExcelFile(archivo_unificado)
    df     = pd.read_excel(xls, sheet_name="Survey Data")
    df_dcp = pd.read_excel(xls, sheet_name="DCP Data")

    df = df.rename(columns={
        "Dist From Start":          "PK_equipo",
        "On Voltage":               "On_V",
        "Off Voltage":              "Off_V",
        "Latitude":                 "Lat",
        "Longitude":                "Long",
        "Comment":                  "Comentario",
        "DCP/Feature/DCVG Anomaly": "Anomalia",
    })

    # 2. Limpieza de coordenadas
    def _clean(v):
        try:
            v = str(v).strip()
            return np.nan if v in ("", ".", "-", "None", "nan") else float(v)
        except Exception:
            return np.nan

    df["Lat"]  = df["Lat"].apply(_clean)
    df["Long"] = df["Long"].apply(_clean)

    # 3. Interpolación GPS faltante
    for coord in ("Lat", "Long"):
        mask = df[coord].isna()
        if mask.any() and (~mask).sum() >= 2:
            m = LinearRegression()
            m.fit(df.loc[~mask, ["PK_equipo"]], df.loc[~mask, coord])
            df.loc[mask, coord] = m.predict(df.loc[mask, ["PK_equipo"]])

    # 4. WGS84 → EPSG:3857
    t_fwd = Transformer.from_crs("EPSG:4326", "EPSG:3857", always_xy=True)
    df["X"], df["Y"] = t_fwd.transform(df["Long"].values, df["Lat"].values)

    # Crear columna de geometría como objetos shapely Point
    df["geometry"] = [Point(x, y) for x, y in zip(df["X"], df["Y"])]

    # 5. Leer y proyectar ducto
    linea = _leer_linea_proyectada(shp_path)
    print(f"✔ Ducto: {linea.geom_type}, longitud = {round(linea.length, 1)} m")

    # 6. Snap de puntos a la traza
    snapped = [linea.interpolate(linea.project(p)) for p in df["geometry"]]
    df["geom_snap"]    = snapped
    df["Dist_traza_m"] = [p.distance(s) for p, s in zip(df["geometry"], snapped)]
    df["geometry"]     = snapped

    # 7. PK geométrico
    df["PK_geom_m"] = [linea.project(p) for p in df["geometry"]]

    # 8. Auto-detección sentido PK
    df_pk = df[["PK_equipo", "PK_geom_m"]].dropna()
    corr  = df_pk["PK_equipo"].corr(df_pk["PK_geom_m"])
    print(f"Correlación PK equipo vs geométrico: {round(corr, 3)}")
    if corr < 0:
        print("⚠ PK invertido — corrigiendo...")
        df["PK_geom_m"] = linea.length - df["PK_geom_m"]

    df["PK_real_m"]  = df["PK_geom_m"] - df["PK_geom_m"].min()
    df["PK_real_km"] = df["PK_real_m"] / 1000
    df = df.sort_values("PK_geom_m").reset_index(drop=True)

    # 9. Coordenadas corregidas (EPSG:3857 → WGS84)
    t_back = Transformer.from_crs("EPSG:3857", "EPSG:4326", always_xy=True)
    xs = [p.x for p in df["geometry"]]
    ys = [p.y for p in df["geometry"]]
    df["Long_corr"], df["Lat_corr"] = t_back.transform(xs, ys)

    # 10. mV
    df["On_mV"]      = df["On_V"]  * 1000
    df["Off_mV"]     = df["Off_V"] * 1000
    df["IR_Drop_mV"] = df["On_mV"] - df["Off_mV"]

    # 11. Limpieza de outliers
    WINDOW, UMBRAL = 15, 250
    for col, out_col in [("Off_mV", "Off_mV_limpio"), ("On_mV", "On_mV_limpio")]:
        med = df[col].rolling(WINDOW, center=True).median()
        mask_out = abs(df[col] - med) > UMBRAL
        df[out_col] = df[col].copy()
        df.loc[mask_out, out_col] = med[mask_out]

    df["IR_Drop_mV_limpio"] = df["On_mV_limpio"] - df["Off_mV_limpio"]

    # 12. Estado CP
    def validar(off):
        if pd.isna(off):     return "DESPROTEGIDO"
        if off <= CRITERIO_MARGINAL: return "SOBREPROTEGIDO"
        if off <= CRITERIO_OK:       return "PROTEGIDO"
        return "DESPROTEGIDO"

    df["Estado_CP"] = df["Off_mV_limpio"].apply(validar)

    # 13. Exportar
    DROP = ["X", "Y", "geometry", "geom_snap", "Off Time",
            "Fix Quality", "GPS Type", "Sats In Use",
            "PDOP", "HDOP", "VDOP", "Fix Time",
            "Dist_traza_m", "PK_real_km", "On_mV", "Off_mV", "IR_Drop_mV"]
    df_final = df.drop(columns=DROP, errors="ignore")

    salida = os.path.join(carpeta, "CIPS_VALIDADO_FINAL.xlsx")
    with pd.ExcelWriter(salida, engine="openpyxl") as writer:
        df_final.to_excel(writer, sheet_name="Survey Data", index=False)
        df_dcp.to_excel(writer,   sheet_name="DCP Data",    index=False)

    print(f"✔ Exportado: {salida} ({len(df_final)} filas)")
    return salida
