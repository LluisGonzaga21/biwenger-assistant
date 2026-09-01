"""
Funciones compartidas para cargar y analizar datos de Biwenger (mercado,
plantillas rivales, tu plantilla, ratios puntos/valor, forma reciente y
sugerencia de fichajes).

Las usan tanto 02_generar_excel.py como el notebook de analisis
(03_analisis.ipynb) para no duplicar la logica. Si tocas algo aqui, afecta
a los dos.

IMPORTANTE sobre 'score_id': Biwenger calcula los puntos de cada jugador en
paralelo con varios sistemas de puntuacion distintos (1, 2, 3...). Tanto
'seasons[].points' como 'reports[].points' son diccionarios {scoreID: pts},
NO un numero suelto. Hay que usar el scoreID de TU liga (league['scoreID'],
visible en output/debug/league.json) para leer el que realmente cuenta en
tu clasificacion. Todas las funciones de aqui que leen puntos piden
score_id explicitamente por eso.
"""

from __future__ import annotations

import json
import time
from datetime import datetime
from pathlib import Path

import pandas as pd

from biwenger_client import BiwengerAPIError, BiwengerClient

CACHE_DIR = Path(__file__).parent / "output" / "cache"
HISTORY_DIR = Path(__file__).parent / "output" / "history"
_CACHE_FILES = {
    "standings": CACHE_DIR / "standings.json",
    "score_id": CACHE_DIR / "score_id.json",
    "mi_team_id": CACHE_DIR / "mi_team_id.json",
    "balance": CACHE_DIR / "balance.json",
    "mercado_raw": CACHE_DIR / "mercado_raw.json",
    "rosters_raw": CACHE_DIR / "rosters_raw.json",
    "detalles": CACHE_DIR / "detalles.json",
    "ofertas_raw": CACHE_DIR / "ofertas_raw.json",
}

POSICIONES = {1: "Portero", 2: "Defensa", 3: "Centrocampista", 4: "Delantero"}

# Valores de 'status' del jugador que significan que no esta disponible para
# jugar. Confirmados con datos reales de la API: "ok", "doubt", "injured",
# "sanctioned", "discarded". "doubt" (duda de ultima hora) se deja FUERA a
# proposito -- suele significar que podria jugar, no que seguro no juega, asi
# que no se excluye automaticamente de los recomendadores (se ve igualmente
# en la columna 'Estado' para que decidas tu). Si ves otro valor nuevo en
# output/cache/detalles.json, anadelo aqui.
ESTADOS_NO_DISPONIBLE = {"injured", "sanctioned", "discarded"}

# ---------------------------------------------------------------------------
# Nombres de campo candidatos por concepto. Si algo sale vacio en el
# analisis, este es el primer sitio a mirar/editar tras revisar
# output/debug/*.json (ejecuta 01_explorar_api.py para generarlos).
# ---------------------------------------------------------------------------
CAMPOS_CANDIDATOS = {
    "market_list": ["sales", "players", "items", "data"],
    "market_player_id": ["player", "id"],
    "market_price": ["price", "value", "amount"],
    "roster_price": ["price", "marketValue", "value"],
    "roster_clause": ["clause", "buyoutClause", "clausePrice", "clauseValue"],
    "player_price": ["price", "value"],
}


def first_present(d: dict, keys: list, default=None):
    if not isinstance(d, dict):
        return default
    for k in keys:
        if k in d and d[k] is not None:
            return d[k]
    return default


def _dig(d, path):
    cur = d
    for p in path:
        if not isinstance(cur, dict) or p not in cur:
            return None
        cur = cur[p]
    return cur


def puntos_por_score(points_dict, score_id):
    """Extrae el numero de puntos para tu score_id de un dict {scoreID: pts}."""
    if not isinstance(points_dict, dict):
        return None
    return points_dict.get(str(score_id))


def puntos_por_partido(puntos, partidos):
    """Puntos por partido jugado. None si no hay partidos o puntos."""
    if puntos is None or not partidos:
        return None
    try:
        return round(puntos / partidos, 2)
    except (TypeError, ZeroDivisionError):
        return None


def _temporadas_liga_recientes(seasons, n=2):
    """
    De una lista 'seasons' (tal cual la da la API), devuelve como mucho las
    `n` mas recientes DE LIGA (excluye copa/champions, que aparecen como
    entradas con 'competition'), ordenadas de mas reciente a mas antigua.
    """
    if not seasons or not isinstance(seasons, list):
        return []
    ligas = [s for s in seasons if isinstance(s, dict) and not s.get("competition")]

    def sort_key(s):
        try:
            return int(s.get("id"))
        except (TypeError, ValueError):
            return 0

    return sorted(ligas, key=sort_key, reverse=True)[:n]


def recortar_temporadas_antiguas(detalle):
    """
    Recorta 'seasons' del detalle de un jugador a solo la temporada actual
    y la anterior de LIGA -- es lo unico que usa el analisis (ver
    extract_seasons_points); todo lo mas antiguo, y las copas/champions, se
    descartan. Se aplica antes de guardar en cache para no arrastrar el
    historico completo sin necesidad. Devuelve un detalle nuevo (no muta el
    original).
    """
    seasons = detalle.get("seasons") if isinstance(detalle, dict) else None
    if not isinstance(seasons, list):
        return detalle
    nuevo = dict(detalle)
    nuevo["seasons"] = _temporadas_liga_recientes(seasons, n=2)
    return nuevo


def temporada_liga_mas_comun(detalles: dict):
    """
    Id de la temporada de LIGA 'actual' de verdad, deducido por MAYORIA:
    la temporada mas reciente que tiene registrada cada jugador, la mas
    repetida entre todos. En una liga activa la mayoria de los ~200
    jugadores habran jugado esta temporada, asi que su moda es una
    referencia fiable de cual es "ahora" -- mejor que fiarse ciegamente de
    la temporada mas reciente de CADA jugador por separado, que para
    alguien que lleva tiempo sin jugar (lesion larga, sin equipo, etc.)
    puede ser de hace 2+ temporadas.
    """
    from collections import Counter
    ids = []
    for detalle in detalles.values():
        recientes = _temporadas_liga_recientes(
            detalle.get("seasons") if isinstance(detalle, dict) else None, n=1
        )
        if recientes:
            ids.append(str(recientes[0].get("id")))
    if not ids:
        return None
    return Counter(ids).most_common(1)[0][0]


def extract_seasons_points(detalle, score_id, temporada_actual_id=None, temporada_anterior_id=None):
    """
    Busca en 'seasons' la temporada actual y la anterior DE LIGA (excluye
    copa/champions, que aparecen como entradas con 'competition') y devuelve:
    (pts_actual, partidos_actual, pts_anterior, partidos_anterior,
    etiqueta_actual, etiqueta_anterior).

    Si se pasan temporada_actual_id / temporada_anterior_id (ver
    temporada_liga_mas_comun), exige que la temporada de liga tenga
    exactamente ese id para contar como "actual"/"anterior". Sin esto, a un
    jugador que lleva tiempo sin jugar (su historial mas reciente es de
    hace 2+ temporadas, p.ej. lesion larga o sin equipo) se le atribuirian
    esos puntos viejos como si fueran de esta temporada o la pasada -- eso
    desvirtuaria por completo su Puntuacion potencial. Sin esos parametros,
    usa sin mas las 2 temporadas de liga mas recientes del jugador
    (comportamiento antiguo, menos fiable para casos asi).
    """
    seasons = detalle.get("seasons") if isinstance(detalle, dict) else None

    def label(s):
        return str(first_present(s, ["name", "slug", "id"], "?")) if s else None

    if temporada_actual_id is not None:
        ligas = [s for s in (seasons or []) if isinstance(s, dict) and not s.get("competition")]
        por_id = {str(s.get("id")): s for s in ligas}
        actual = por_id.get(str(temporada_actual_id))
        anterior = por_id.get(str(temporada_anterior_id)) if temporada_anterior_id is not None else None
    else:
        ordered = _temporadas_liga_recientes(seasons)
        actual = ordered[0] if len(ordered) > 0 else None
        anterior = ordered[1] if len(ordered) > 1 else None

    if actual is None and anterior is None:
        return None, None, None, None, None, None

    return (
        puntos_por_score(actual.get("points"), score_id) if actual else None,
        actual.get("games") if actual else None,
        puntos_por_score(anterior.get("points"), score_id) if anterior else None,
        anterior.get("games") if anterior else None,
        label(actual),
        label(anterior),
    )


def calcular_forma(detalle, score_id, decay=0.85, n_partidos=6):
    """
    Media ponderada de los puntos de los ultimos partidos jugados esta
    temporada, dando mas peso a los mas recientes (a partir de
    'reports', que trae un resultado por partido con fecha).

    decay (0-1): cuanto mas bajo, mas se prioriza el ultimo partido frente
    a los de hace unas jornadas. n_partidos: cuantos partidos recientes
    entran en la media.
    """
    reports = detalle.get("reports") if isinstance(detalle, dict) else None
    if not isinstance(reports, list) or not reports:
        return None

    partidos = []
    for r in reports:
        if not isinstance(r, dict):
            continue
        pts = puntos_por_score(r.get("points"), score_id)
        fecha = _dig(r, ["match", "date"])
        if pts is None or fecha is None:
            continue
        partidos.append((fecha, pts))
    if not partidos:
        return None

    partidos.sort(key=lambda t: t[0])  # mas antiguo -> mas reciente
    partidos = partidos[-n_partidos:]

    pesos = [decay ** i for i in range(len(partidos) - 1, -1, -1)]  # ultimo partido -> peso maximo
    total_peso = sum(pesos)
    return round(sum(p * w for (_, p), w in zip(partidos, pesos)) / total_peso, 2)


def puntuacion_potencial(detalle, score_id, decay=0.85, n_partidos=6, peso_temporada_anterior=0.3,
                          temporada_actual_id=None, temporada_anterior_id=None):
    """
    Metrica de rendimiento "potencial" pensada para comparar candidatos a
    fichaje: mezcla la forma reciente (partidos de esta temporada, mas peso
    a los ultimos) con los puntos/partido de la temporada anterior de LaLiga
    (con menos peso, y solo si jugo en primera). Si falta alguno de los dos
    componentes, usa solo el que haya disponible.
    """
    forma = calcular_forma(detalle, score_id, decay=decay, n_partidos=n_partidos)
    _, _, pts_anterior, partidos_anterior, _, _ = extract_seasons_points(
        detalle, score_id, temporada_actual_id, temporada_anterior_id
    )
    ppg_anterior = puntos_por_partido(pts_anterior, partidos_anterior)

    if forma is None and ppg_anterior is None:
        return None
    if forma is None:
        return ppg_anterior
    if ppg_anterior is None:
        return forma
    return round(forma * (1 - peso_temporada_anterior) + ppg_anterior * peso_temporada_anterior, 2)


def esta_disponible(detalle):
    """False si el 'status' del jugador esta en ESTADOS_NO_DISPONIBLE (lesion,
    sancion...). True si esta 'ok' o si no hay status (mejor no descartar
    por error)."""
    estado = detalle.get("status") if isinstance(detalle, dict) else None
    if not estado:
        return True
    return str(estado).lower() not in ESTADOS_NO_DISPONIBLE


def dificultad_proximo_partido(detalle):
    """
    Info del proximo partido de LaLiga del jugador a partir de
    'team.nextGames[0]': rival, si juega en casa, la jornada, y una
    dificultad (0-100, mas alto = mas dificil para su equipo). None si no
    hay proximo partido programado.
    """
    team = detalle.get("team") if isinstance(detalle, dict) else None
    next_games = team.get("nextGames") if isinstance(team, dict) else None
    if not isinstance(next_games, list) or not next_games:
        return None
    partido = next_games[0]
    team_id = team.get("id")
    home = partido.get("home") or {}
    away = partido.get("away") or {}
    es_local = home.get("id") == team_id
    mi_lado, rival_lado = (home, away) if es_local else (away, home)
    return {
        "rival": rival_lado.get("name"),
        "es_local": es_local,
        "dificultad": _dig(mi_lado, ["difficulty", "rating"]),
        "jornada": _dig(partido, ["round", "name"]),
    }


def tendencia_precio(detalle, dias=7):
    """
    % de variacion del valor de mercado en los ultimos `dias` puntos del
    historico diario de precios ('prices': lista [fecha, precio]
    cronologica). Positivo = subiendo, negativo = bajando. None si no hay
    historico suficiente.
    """
    precios = detalle.get("prices") if isinstance(detalle, dict) else None
    if not isinstance(precios, list) or len(precios) < 2:
        return None
    serie = [p[1] for p in precios if isinstance(p, (list, tuple)) and len(p) == 2 and p[1]]
    if len(serie) < 2:
        return None
    ventana = serie[-dias:] if len(serie) >= dias else serie
    inicio, fin = ventana[0], ventana[-1]
    if not inicio:
        return None
    return round((fin - inicio) / inicio * 100, 1)


PARTIDOS_TEMPORADA_COMPLETA = 38  # jornadas de una liga de 20 equipos (ida y vuelta)
VALOR_MINIMO_RELEVANTE = 400_000  # por debajo de esto, se considera "morralla"


def jugador_irrelevante(detalle, score_id, temporada_actual_id=None, temporada_anterior_id=None):
    """
    True si el jugador no aporta nada util al analisis de mercado/rivales:
    - valor de mercado desconocido o <= VALOR_MINIMO_RELEVANTE ("morralla"), o
    - no esta inscrito con ningun equipo de 1a division ('team' vacio en la
      API -- ver el caso de Owono), o
    - no ha debutado esta temporada (sin 'Puntos temporada actual') Y jugo
      menos de la mitad de los partidos la temporada anterior (jugador
      practicamente inactivo o fuera del primer equipo).
    """
    valor = first_present(detalle, CAMPOS_CANDIDATOS["player_price"])
    if valor is None or valor <= VALOR_MINIMO_RELEVANTE:
        return True

    equipo = detalle.get("team") if isinstance(detalle, dict) else None
    if not isinstance(equipo, dict) or not equipo.get("name"):
        return True

    pts_actual, _, _, partidos_anterior, _, _ = extract_seasons_points(
        detalle, score_id, temporada_actual_id, temporada_anterior_id
    )
    if pts_actual is not None:
        return False  # ha debutado esta temporada
    return (partidos_anterior or 0) < (PARTIDOS_TEMPORADA_COMPLETA / 2)


def ratio(points, value):
    """Puntos por millon de valor (de mercado o de clausula). None si no se puede calcular."""
    if points is None or not value:
        return None
    try:
        return round(points / (value / 1_000_000), 2)
    except (TypeError, ZeroDivisionError):
        return None


def construir_fila_jugador(pid, detalle, score_id, extra=None,
                            temporada_actual_id=None, temporada_anterior_id=None):
    extra = extra or {}
    nombre = first_present(detalle, ["name"], f"id:{pid}")
    posicion_id = first_present(detalle, ["position"])
    posicion = POSICIONES.get(posicion_id, posicion_id)
    equipo = first_present(detalle.get("team", {}) if isinstance(detalle.get("team"), dict) else {}, ["name"])
    valor_actual = first_present(detalle, CAMPOS_CANDIDATOS["player_price"])

    pts_actual, partidos_actual, pts_anterior, partidos_anterior, etq_actual, etq_anterior = (
        extract_seasons_points(detalle, score_id, temporada_actual_id, temporada_anterior_id)
    )
    proximo = dificultad_proximo_partido(detalle) or {}

    fila = {
        "player_id": pid,
        "Jugador": nombre,
        "Posicion": posicion,
        "Equipo LaLiga": equipo,
        "Estado": first_present(detalle, ["status"], "ok"),
        "Disponible": esta_disponible(detalle),
        "Valor de mercado": valor_actual,
        "Tendencia precio (7d) %": tendencia_precio(detalle),
        # nombres de columna FIJOS a proposito (no el nombre literal de la
        # temporada, p.ej. "Temporada 2025/2026") -- si no, jugadores sin
        # temporada anterior (debutantes, fichajes de fuera) generan
        # columnas distintas ("Puntos temp. anterior") y el df se llena de
        # columnas casi-duplicadas. 'Temporada actual/anterior' llevan la
        # etiqueta real si hace falta consultarla.
        "Temporada actual": etq_actual,
        "Puntos temporada actual": pts_actual,
        "Partidos temporada actual": partidos_actual,
        "Pts/partido (actual)": puntos_por_partido(pts_actual, partidos_actual),
        "Ratio pts/VM (actual)": ratio(pts_actual, valor_actual),
        "Temporada anterior": etq_anterior,
        "Puntos temporada anterior": pts_anterior,
        "Partidos temporada anterior": partidos_anterior,
        "Pts/partido (anterior)": puntos_por_partido(pts_anterior, partidos_anterior),
        "Ratio pts/VM (anterior)": ratio(pts_anterior, valor_actual),
        "Forma (ult. partidos)": calcular_forma(detalle, score_id),
        "Puntuacion potencial": puntuacion_potencial(
            detalle, score_id, temporada_actual_id=temporada_actual_id, temporada_anterior_id=temporada_anterior_id
        ),
        "Proximo rival": proximo.get("rival"),
        "Local/Visitante": ("Local" if proximo.get("es_local") else "Visitante") if proximo else None,
        "Dificultad proximo partido": proximo.get("dificultad"),
        "Jornada proxima": proximo.get("jornada"),
    }
    fila.update(extra)
    return fila


def sugerir_fichajes(df: pd.DataFrame, balance, top_n=15, solo_disponibles=True, ordenar_por="mejora"):
    """
    Sugiere fichajes (clausula a un rival o compra en el mercado) que
    mejorarian tu equipo, comparando 1 a 1: cada candidato asequible contra
    tu jugador mas flojo (por 'Puntuacion potencial') de su misma posicion.

    df: el DataFrame combinado de construir_dataframe() (necesita las
    columnas 'Origen', 'Posicion', 'Puntuacion potencial' y, para el coste,
    'Precio en mercado' y/o 'Clausula').
    balance: tu saldo disponible (importe maximo a gastar).
    solo_disponibles: si True (por defecto), descarta candidatos lesionados
    o sancionados ('Disponible' == False).
    ordenar_por: 'mejora' (por defecto, mejora absoluta en pts/partido) o
    'eficiencia' (mejora por millon gastado -- prioriza fichajes baratos que
    mejoran poco a poco frente a fichajes carisimos que mejoran mucho).

    Descarta siempre las clausulas de rival que esten bloqueadas ahora
    mismo (tras una compra/clausulazo reciente, la API la rechazaria) --
    ver 'Clausula bloqueada hasta' en el df.
    """
    requeridas = {"Origen", "Posicion", "Puntuacion potencial"}
    faltan = requeridas - set(df.columns)
    if faltan:
        raise ValueError(f"Al DataFrame le faltan columnas: {faltan}")
    if ordenar_por not in ("mejora", "eficiencia"):
        raise ValueError("ordenar_por debe ser 'mejora' o 'eficiencia'")

    mi_equipo = df[(df["Origen"] == "Mi equipo") & df["Puntuacion potencial"].notna()]
    if mi_equipo.empty:
        raise ValueError("No hay jugadores de 'Mi equipo' con Puntuacion potencial calculable.")

    mi_equipo_ordenado = mi_equipo.sort_values("Puntuacion potencial")
    listones = mi_equipo_ordenado.groupby("Posicion")["Puntuacion potencial"].min()
    peor_jugador = mi_equipo_ordenado.groupby("Posicion")["Jugador"].first()
    # cuantos de tus jugadores en cada posicion rinden por debajo de tu
    # propia mediana ahi -- mas de uno indica una posicion con mas de un
    # eslabon debil, no solo el peor
    medianas_propias = mi_equipo.groupby("Posicion")["Puntuacion potencial"].median()
    total_propios = mi_equipo.groupby("Posicion")["Jugador"].count()
    flojos_propios = (
        mi_equipo[mi_equipo["Puntuacion potencial"] <= mi_equipo["Posicion"].map(medianas_propias)]
        .groupby("Posicion")["Jugador"].count()
    )

    candidatos = df[df["Origen"] != "Mi equipo"].copy()
    if solo_disponibles and "Disponible" in candidatos.columns:
        candidatos = candidatos[candidatos["Disponible"]]
    if "Clausula disponible ahora" in candidatos.columns:
        # las filas de Mercado no tienen esta columna (NaN) -- se dejan pasar,
        # solo se descartan clausulas de rival EXPLICITAMENTE bloqueadas ahora
        candidatos = candidatos[candidatos["Clausula disponible ahora"] != False]  # noqa: E712
    coste_mercado = candidatos["Precio en mercado"] if "Precio en mercado" in candidatos else None
    coste_clausula = candidatos["Clausula"] if "Clausula" in candidatos else None
    if coste_mercado is not None and coste_clausula is not None:
        candidatos["Coste fichaje"] = coste_mercado.combine_first(coste_clausula)
    else:
        candidatos["Coste fichaje"] = coste_mercado if coste_mercado is not None else coste_clausula

    candidatos = candidatos[candidatos["Coste fichaje"].notna() & (candidatos["Coste fichaje"] <= balance)]
    candidatos = candidatos[candidatos["Puntuacion potencial"].notna()]
    candidatos = candidatos[candidatos["Posicion"].isin(listones.index)]

    candidatos["Tu listón en esa posicion"] = candidatos["Posicion"].map(listones)
    candidatos["Jugador que reemplazarías"] = candidatos["Posicion"].map(peor_jugador)
    candidatos["Mejora estimada (pts/partido)"] = (
        candidatos["Puntuacion potencial"] - candidatos["Tu listón en esa posicion"]
    )
    candidatos = candidatos[candidatos["Mejora estimada (pts/partido)"] > 0]

    candidatos["Saldo tras fichar"] = balance - candidatos["Coste fichaje"]
    candidatos["Mejora por millon gastado"] = (
        candidatos["Mejora estimada (pts/partido)"] / (candidatos["Coste fichaje"] / 1_000_000)
    ).round(3)
    candidatos["Jugadores flojos en tu posicion"] = (
        candidatos["Posicion"].map(flojos_propios).fillna(0).astype(int)
    )
    candidatos["Total en tu posicion"] = candidatos["Posicion"].map(total_propios).fillna(0).astype(int)

    columna_orden = "Mejora por millon gastado" if ordenar_por == "eficiencia" else "Mejora estimada (pts/partido)"

    columnas = [
        "player_id", "Jugador", "Posicion", "Equipo LaLiga", "Origen", "Coste fichaje",
        "Saldo tras fichar", "Vendedor (id)", "Puntuacion potencial", "Jugador que reemplazarías",
        "Tu listón en esa posicion", "Mejora estimada (pts/partido)", "Mejora por millon gastado",
        "Jugadores flojos en tu posicion", "Total en tu posicion",
    ]
    return (
        candidatos[columnas]
        .sort_values(columna_orden, ascending=False)
        .head(top_n)
        .reset_index(drop=True)
    )


def unico_por_jugador(df: pd.DataFrame) -> pd.DataFrame:
    """
    Un jugador puede aparecer en mas de una fila de construir_dataframe: una
    por estar en la plantilla de un equipo, otra por estar ademas puesto a
    la venta en el mercado por su propio manager (ver la nota en la
    seccion 2 del notebook). Eso es correcto para tablas de compra/venta
    (Mercado vs Rival tienen coste distinto -- precio de mercado vs
    clausula), pero en un RANKING hace que el mismo jugador ocupe varios
    puestos en vez de dejar sitio a otros.

    Esta funcion colapsa esas filas a una por jugador, combinando los
    'Origen' distintos en un solo texto (p.ej. 'Mercado + Rival: Botas FC').
    Usala antes de un top-N cuando no te interesa distinguir por donde se
    consigue, solo el ranking de rendimiento.
    """
    if df.empty or "player_id" not in df.columns:
        return df
    otras_cols = [c for c in df.columns if c not in ("player_id", "Origen")]
    combinado = df.groupby("player_id", as_index=False).agg({
        **{c: "first" for c in otras_cols},
        "Origen": lambda s: " + ".join(sorted(set(s))),
    })
    return combinado[df.columns.tolist()]


def coste_por_punto(df: pd.DataFrame, metrica="Puntuacion potencial", margen_pct=20) -> pd.DataFrame:
    """
    Para cada jugador (cualquier posicion/origen), cuanto cuesta (Valor de
    mercado) por cada punto de rendimiento (`metrica`) que genera, y como
    se compara contra la MEDIANA de TODO EL JUEGO -- todos los jugadores
    validos de tu dataset, sin distinguir posicion (a diferencia de
    calcular_ratio_referencia/estimar_valor_justo, que si separan por
    posicion). Mas bajo = mas barato por punto = mejor.

    Usa unico_por_jugador() primero si quieres un ranking sin duplicados
    por jugadores que estan a la vez en mercado y en una plantilla.

    Se excluyen jugadores con metrica <= 0 (un coste por punto negativo o
    infinito no tiene sentido) o sin valor de mercado.
    """
    validos = df[df[metrica].notna() & (df[metrica] > 0) & df["Valor de mercado"].notna() & (df["Valor de mercado"] > 0)].copy()
    validos["Coste por punto"] = (validos["Valor de mercado"] / validos[metrica]).round(0)

    media_juego = validos["Coste por punto"].median()
    validos["Media del juego (referencia)"] = media_juego
    validos["% vs media del juego"] = ((validos["Coste por punto"] - media_juego) / media_juego * 100).round(1)

    def etiqueta(pct):
        if pct <= -margen_pct:
            return "Muy por debajo (barato)"
        if pct >= margen_pct:
            return "Muy por encima (caro)"
        return "En la media"

    validos["Valoracion"] = validos["% vs media del juego"].apply(etiqueta)

    columnas = [c for c in [
        "player_id", "Jugador", "Posicion", "Origen", "Valor de mercado", metrica,
        "Coste por punto", "Media del juego (referencia)", "% vs media del juego", "Valoracion",
    ] if c in validos.columns]
    return validos[columnas].sort_values("Coste por punto").reset_index(drop=True)


def mejores_jugadores(df: pd.DataFrame, posicion=None, origen=None, top_n=20,
                       metrica="Puntuacion potencial", solo_disponibles=True):
    """
    Ranking general de jugadores por metrica de rendimiento (por defecto
    'Puntuacion potencial'), sin mirar precio ni si te lo puedes permitir.
    Sirve como watchlist: quien esta en mejor forma ahora mismo.

    posicion: filtra por 'Portero'/'Defensa'/'Centrocampista'/'Delantero'.
    origen: filtra por 'Mercado', 'Mi equipo', o el 'Origen' exacto de un
    rival (p.ej. 'Rival: Botas FC'). Por defecto no filtra (incluye todos),
    y en ese caso colapsa jugadores duplicados en varios origenes (ver
    unico_por_jugador) para que no ocupen varios puestos del ranking.
    solo_disponibles: si True (por defecto), descarta lesionados/sancionados.
    """
    resultado = df.dropna(subset=[metrica])
    if solo_disponibles and "Disponible" in resultado.columns:
        resultado = resultado[resultado["Disponible"]]
    if posicion:
        resultado = resultado[resultado["Posicion"] == posicion]
    if origen:
        resultado = resultado[resultado["Origen"] == origen]
    else:
        resultado = unico_por_jugador(resultado)

    columnas = [c for c in [
        "player_id", "Jugador", "Posicion", "Equipo LaLiga", "Origen",
        metrica, "Forma (ult. partidos)", "Proximo rival", "Dificultad proximo partido",
        "Valor de mercado", "Tendencia precio (7d) %", "Clausula", "Precio en mercado",
    ] if c in resultado.columns]

    return (
        resultado[columnas]
        .sort_values(metrica, ascending=False)
        .head(top_n)
        .reset_index(drop=True)
    )


def calcular_ratio_referencia(df: pd.DataFrame, metrica="Puntuacion potencial"):
    """
    Ratio de referencia (metrica / valor de mercado en millones) por
    posicion, calculado con la MEDIANA de todos los jugadores validos
    (mercado + rivales + tu equipo -> muestra amplia). Es la base para
    estimar un valor de mercado "justo" a partir del rendimiento.
    """
    tmp = df.dropna(subset=[metrica, "Valor de mercado"])
    tmp = tmp[tmp["Valor de mercado"] > 0].copy()
    tmp["_ratio"] = tmp[metrica] / (tmp["Valor de mercado"] / 1_000_000)
    return tmp.groupby("Posicion")["_ratio"].median()


def estimar_valor_justo(df: pd.DataFrame, metrica="Puntuacion potencial", margen_pct=20, solo_disponibles=True):
    """
    Para los jugadores EN VENTA en el mercado, estima un "valor justo" a
    partir de como el resto de jugadores de tu liga (de esa misma posicion)
    relacionan su rendimiento con su valor de mercado (ver
    calcular_ratio_referencia), y lo compara con el precio de venta actual
    para detectar chollos/sobreprecios y sugerir una oferta.

    Es una heuristica basada en jugadores similares de tu propia liga, NO
    el precio "real" segun Biwenger -- usalo como orientacion, no como
    verdad absoluta. margen_pct: a partir de que % de diferencia se
    etiqueta como 'Chollo'/'Caro' (por debajo se considera 'Precio justo').
    solo_disponibles: si True (por defecto), descarta lesionados/sancionados
    (un chollo que no puede jugar no es tan chollo).
    """
    ratios_ref = calcular_ratio_referencia(df, metrica=metrica)

    mercado = df[df["Origen"] == "Mercado"].dropna(subset=[metrica, "Precio en mercado"]).copy()
    if solo_disponibles and "Disponible" in mercado.columns:
        mercado = mercado[mercado["Disponible"]]
    mercado["Ratio referencia posicion"] = mercado["Posicion"].map(ratios_ref)
    mercado = mercado.dropna(subset=["Ratio referencia posicion"])
    mercado["Valor justo estimado"] = (
        mercado[metrica] / mercado["Ratio referencia posicion"] * 1_000_000
    ).round(-3)
    mercado["Diferencia vs precio pedido"] = mercado["Valor justo estimado"] - mercado["Precio en mercado"]
    mercado["% sobre precio pedido"] = (
        mercado["Diferencia vs precio pedido"] / mercado["Precio en mercado"] * 100
    ).round(1)
    # nunca tiene sentido ofrecer mas del precio pedido, ni mas de lo que
    # consideramos su valor justo
    mercado["Oferta sugerida"] = mercado[["Valor justo estimado", "Precio en mercado"]].min(axis=1)

    def etiqueta(pct):
        if pct >= margen_pct:
            return "Chollo"
        if pct <= -margen_pct:
            return "Caro"
        return "Precio justo"

    mercado["Valoracion"] = mercado["% sobre precio pedido"].apply(etiqueta)

    columnas = [
        "player_id", "Jugador", "Posicion", "Equipo LaLiga", "Precio en mercado",
        "Valor justo estimado", "Diferencia vs precio pedido", "% sobre precio pedido",
        "Valoracion", "Oferta sugerida", "Vendedor (id)",
    ]
    return (
        mercado[columnas]
        .sort_values("% sobre precio pedido", ascending=False)
        .reset_index(drop=True)
    )


def sugerir_oferta_mercado(df: pd.DataFrame, data: dict, margen_sobre_pedido_pct=5,
                            metrica="Puntuacion potencial", rango_comparables_pct=40):
    """
    Para cada jugador LIBRE del mercado (Origen == 'Mercado'; para
    clausulazos no aplica -- el importe es fijo, no se puja), estima
    cuanto haria falta ofertar de verdad para llevartelo, no solo cuanto
    "vale" segun tu propia heuristica.

    Biwenger no permite pujar por debajo del precio pedido actual, asi que
    ese precio es siempre el SUELO de la oferta recomendada. A partir de
    ahi:

    1. Comparables reales: la MEDIANA de lo que los duenos actuales de
       jugadores similares (misma posicion, valor de mercado dentro de
       +/-rango_comparables_pct%) pagaron de verdad por ellos (precio de
       compra real, no clausula). Si esa mediana supera el precio pedido,
       es la mejor referencia de "lo que hace falta en la practica" -- se
       usa como base en vez del precio pedido.
    2. Competencia real: cuantos RIVALES tienen presupuesto (`maximumBid`
       de standings -- ojo, puede ser mayor que su saldo actual, Biwenger
       permite pujar en descubierto hasta ese limite) para superar esa
       base. Si hay competencia, sube un poco mas (margen_sobre_pedido_pct)
       para no perderlo por una diferencia pequenya.

    Tu valor justo estimado (misma heuristica que estimar_valor_justo) se
    incluye solo como referencia informativa -- para que veas si lo que
    hace falta pujar se sale de lo que tu crees que rinde el jugador -- pero
    ya NO limita la oferta recomendada hacia abajo (un jugador puede exigir
    pujar por encima de tu valor justo y aun asi ser la unica forma real de
    ficharlo).

    Esto es una heuristica orientativa (no sabemos quien esta REALMENTE
    interesado en cada jugador, solo quien podria permitirselo economicamente)
    -- no una prediccion garantizada de lo que hara falta para ganarlo.
    """
    ratios_ref = calcular_ratio_referencia(df, metrica=metrica)

    mercado = df[df["Origen"] == "Mercado"].dropna(subset=[metrica, "Precio en mercado"]).copy()
    mercado["Ratio referencia posicion"] = mercado["Posicion"].map(ratios_ref)
    mercado = mercado.dropna(subset=["Ratio referencia posicion"])
    mercado["Valor justo estimado"] = (
        mercado[metrica] / mercado["Ratio referencia posicion"] * 1_000_000
    ).round(-3)

    # precios de compra REALES (lo que pago cada dueno actual, no la clausula)
    precios_reales = []
    for r in data["rosters_raw"].values():
        for f in r["filas"]:
            precio = f.get("precio_roster")
            if precio:
                precios_reales.append({"player_id": f["player_id"], "precio_pagado": precio})
    precios_reales = pd.DataFrame(precios_reales)
    comparables = df.merge(precios_reales, on="player_id", how="inner") if not precios_reales.empty else pd.DataFrame()

    def precio_comparables(row):
        if comparables.empty:
            return None
        rango = row["Valor de mercado"] * rango_comparables_pct / 100
        similares = comparables[
            (comparables["Posicion"] == row["Posicion"])
            & (comparables["Valor de mercado"] >= row["Valor de mercado"] - rango)
            & (comparables["Valor de mercado"] <= row["Valor de mercado"] + rango)
        ]
        if similares.empty:
            return None
        return similares["precio_pagado"].median()

    mercado["Precio pagado por comparables"] = mercado.apply(precio_comparables, axis=1)

    # cuantos rivales (no tu) tienen presupuesto de sobra para superar el precio pedido
    mi_team_id = data["mi_team_id"]
    max_bids_rivales = [
        s.get("maximumBid") for s in data["standings"]
        if str(s.get("id")) != str(mi_team_id) and s.get("maximumBid") is not None
    ]

    def rivales_que_podrian_pujar_mas(precio_pedido):
        return sum(1 for mb in max_bids_rivales if mb > precio_pedido)

    mercado["Rivales que podrian pujar mas"] = mercado["Precio en mercado"].apply(rivales_que_podrian_pujar_mas)

    def oferta_recomendada(row):
        # suelo: nunca por debajo del precio pedido -- Biwenger no deja pujar menos
        base = row["Precio en mercado"]
        comparable = row["Precio pagado por comparables"]
        if pd.notna(comparable) and comparable > base:
            base = comparable
        if row["Rivales que podrian pujar mas"] > 0:
            base = base * (1 + margen_sobre_pedido_pct / 100)
        return round(base / 1000) * 1000

    mercado["Oferta recomendada"] = mercado.apply(oferta_recomendada, axis=1)

    columnas = [
        "player_id", "Jugador", "Posicion", "Equipo LaLiga", "Precio en mercado",
        "Valor justo estimado", "Precio pagado por comparables", "Rivales que podrian pujar mas",
        "Oferta recomendada", "Vendedor (id)",
    ]
    return (
        mercado[columnas]
        .sort_values("Oferta recomendada", ascending=False)
        .reset_index(drop=True)
    )


def sugerir_ventas(df: pd.DataFrame, data: dict = None, metrica="Puntuacion potencial", margen_pct=20):
    """
    El complementario a estimar_valor_justo pero para TU plantilla: compara
    el valor de mercado OFICIAL de cada jugador tuyo con lo que su
    rendimiento actual justificaria (misma mediana de referencia por
    posicion). Si el valor de mercado es mucho mayor que ese "valor justo",
    el jugador esta sobrevalorado -- venderlo ahora aprovecha un precio que
    probablemente no se sostenga si su rendimiento no mejora.

    Si pasas `data` (lo que devuelve cargar_datos_liga), tambien cruza las
    ofertas que has recibido de un manager IDENTIFICADO (ver
    ofertas(solo_identificadas=True) -- descarta la venta rapida
    automatica que Biwenger genera para cada jugador de tu plantilla, que
    no es demanda real): una oferta real en la mano pesa mas que la
    heuristica -- si alguien identificado ya ofrece igual o mas que el
    valor justo estimado, se recomienda vender aunque el % de
    sobrevaloracion por si solo no llegase al umbral.

    Igual que estimar_valor_justo, es una heuristica orientativa basada en
    tu propia liga, no una prediccion garantizada.
    """
    ratios_ref = calcular_ratio_referencia(df, metrica=metrica)

    mios = df[df["Origen"] == "Mi equipo"].dropna(subset=[metrica, "Valor de mercado"]).copy()
    mios["Ratio referencia posicion"] = mios["Posicion"].map(ratios_ref)
    mios = mios.dropna(subset=["Ratio referencia posicion"])
    mios["Valor justo estimado"] = (
        mios[metrica] / mios["Ratio referencia posicion"] * 1_000_000
    ).round(-3)
    mios["Diferencia (VM - justo)"] = mios["Valor de mercado"] - mios["Valor justo estimado"]
    mios["% sobrevaloracion"] = (
        mios["Diferencia (VM - justo)"] / mios["Valor justo estimado"] * 100
    ).round(1)

    mios["Mejor oferta recibida"] = None
    mios["Nº ofertas recibidas"] = 0
    if data is not None:
        recibidas = ofertas(data, tipo="recibidas", solo_identificadas=True)
        if not recibidas.empty:
            mejor_oferta = recibidas.groupby("player_id")["Importe"].max()
            n_ofertas = recibidas.groupby("player_id")["Importe"].count()
            mios["Mejor oferta recibida"] = mios["player_id"].map(mejor_oferta)
            mios["Nº ofertas recibidas"] = mios["player_id"].map(n_ofertas).fillna(0).astype(int)

    def etiqueta(row):
        oferta = row["Mejor oferta recibida"]
        if pd.notna(oferta) and oferta >= row["Valor justo estimado"]:
            return "Vender ahora (tienes oferta igual o por encima del valor justo)"
        if row["% sobrevaloracion"] >= margen_pct:
            return "Vender ahora"
        if row["% sobrevaloracion"] <= -margen_pct:
            return "Infravalorado (no vender)"
        return "Precio acorde a su rendimiento"

    mios["Recomendacion"] = mios.apply(etiqueta, axis=1)

    columnas = [
        "player_id", "Jugador", "Posicion", "Valor de mercado", "Clausula",
        "Valor justo estimado", "Diferencia (VM - justo)", "% sobrevaloracion",
        "Mejor oferta recibida", "Nº ofertas recibidas", "Recomendacion",
    ]
    columnas = [c for c in columnas if c in mios.columns]
    return (
        mios[columnas]
        .sort_values(["Mejor oferta recibida", "% sobrevaloracion"], ascending=False, na_position="last")
        .reset_index(drop=True)
    )


def login_y_resolver_liga(client: BiwengerClient, league_id=None):
    """Inicia sesion y detecta liga + user_id (team-id) para esa liga.

    Lanza RuntimeError si no ha podido resolver ninguna liga (por ejemplo,
    si BIWENGER_LEAGUE_ID no coincide con ninguna de tu cuenta y tampoco se
    ha encontrado ninguna liga automaticamente).
    """
    client.login()
    leagues, match = client.resolve_league(league_id)
    if not client.league_id:
        raise RuntimeError(
            "No se ha podido detectar tu liga automaticamente. Anade "
            "BIWENGER_LEAGUE_ID a tu .env con el ID de tu liga."
        )
    if not league_id and len(leagues) > 1:
        print("  Tienes varias ligas, uso la primera. Si no es la que "
              "quieres, pon BIWENGER_LEAGUE_ID en tu .env:")
        for l in leagues:
            print(f"    - {l.get('name')}: id={l.get('id')}")
    return leagues, match


def cargar_liga(client: BiwengerClient):
    """Devuelve (standings, score_id) de tu liga."""
    league = client.get_league()
    standings = league.get("standings") or []
    if not standings:
        raise RuntimeError(
            "No se ha encontrado 'standings' en la respuesta de la liga. "
            "Ejecuta 01_explorar_api.py y revisa output/debug/league.json."
        )
    score_id = league.get("scoreID")
    if score_id is None:
        raise RuntimeError(
            "No se ha encontrado 'scoreID' en la respuesta de la liga. "
            "Revisa output/debug/league.json."
        )
    return standings, score_id


def identificar_mi_equipo(standings, client: BiwengerClient, own_team_id=None):
    if own_team_id:
        return own_team_id
    for s in standings:
        uid = first_present(s, ["id"]) or first_present(s.get("user", {}) if isinstance(s.get("user"), dict) else {}, ["id"])
        user_obj = s.get("user") if isinstance(s.get("user"), dict) else {}
        if str(uid) == str(client.user_id) or str(user_obj.get("id")) == str(client.user_id):
            return s.get("id")
    print(
        "  [!] No se ha podido identificar tu propio equipo automaticamente "
        "entre los rivales. Pasa own_team_id (tu team id, visible en "
        "output/debug/sample_team.json o league.json) para forzarlo."
    )
    return None


def cargar_plantilla(client: BiwengerClient, team_id, nombre_equipo):
    team = client.get_team(team_id)
    jugadores = team.get("players") if isinstance(team, dict) else None
    filas = []
    if isinstance(jugadores, list):
        for p in jugadores:
            if not isinstance(p, dict):
                continue
            # el precio de compra y la clausula NO estan en el nivel superior
            # del jugador, sino anidados en 'owner' (ver output/debug/sample_team.json)
            owner = p.get("owner") if isinstance(p.get("owner"), dict) else {}
            filas.append(
                {
                    "player_id": p.get("id"),
                    "equipo_rival": nombre_equipo,
                    "precio_roster": first_present(owner, CAMPOS_CANDIDATOS["roster_price"]),
                    "clausula": first_present(owner, CAMPOS_CANDIDATOS["roster_clause"]),
                    # plazo (epoch) hasta el que esa clausula esta bloqueada
                    # (tipico tras una compra/clausulazo reciente); mientras
                    # dure, la API rechazaria un place_offer sobre este jugador
                    "clausula_bloqueada_hasta": owner.get("clauseLockedUntil"),
                }
            )
    return filas


def _parse_mercado(market: dict):
    lista = None
    for key in CAMPOS_CANDIDATOS["market_list"]:
        v = market.get(key) if isinstance(market, dict) else None
        if isinstance(v, list):
            lista = v
            break
    if lista is None:
        print("  [!] No se ha encontrado la lista de jugadores del mercado. "
              "Revisa output/debug/market.json tras ejecutar 01_explorar_api.py.")
        return []

    filas = []
    for entry in lista:
        if not isinstance(entry, dict):
            continue
        player_id = entry.get("player")
        if isinstance(player_id, dict):
            player_id = player_id.get("id")
        elif "playerId" in entry:
            player_id = entry.get("playerId")
        precio = first_present(entry, CAMPOS_CANDIDATOS["market_price"])
        vendedor_id = _dig(entry, ["user", "id"])
        if player_id:
            filas.append({"player_id": player_id, "precio_mercado": precio, "vendedor_id": vendedor_id})
    return filas


def cargar_mercado(client: BiwengerClient):
    return _parse_mercado(client.get_market())


def _parse_ofertas(market: dict):
    """Extrae la lista de ofertas 'en bruto' del mercado (compra/venta,
    tanto recibidas como enviadas). Ver ofertas() para filtrarlas."""
    offers = market.get("offers") if isinstance(market, dict) else None
    if not isinstance(offers, list):
        return []
    filas = []
    for o in offers:
        if not isinstance(o, dict):
            continue
        filas.append({
            "offer_id": o.get("id"),
            "player_id": (o.get("requestedPlayers") or [None])[0],
            "importe": o.get("amount"),
            "tipo": o.get("type"),
            "estado": o.get("status"),
            "de_id": _dig(o, ["from", "id"]),
            "de_nombre": _dig(o, ["from", "name"]),
            "a_id": _dig(o, ["to", "id"]),
            "a_nombre": _dig(o, ["to", "name"]),
            "expira": o.get("until"),
        })
    return filas


def enriquecer_con_detalle(client: BiwengerClient, player_ids, progreso=True, max_pausas_largas=6,
                            on_progress=None, guardar_cada=10, parar_en_primer_limite=False):
    """
    cf.biwenger.com corta por rate-limit tras ~200 peticiones seguidas.
    Probamos a espaciarlas con una pausa preventiva a mitad de camino, pero
    el corte seguia saltando en el mismo punto exacto (jugador ~200) incluso
    tras una pausa de 5 minutos -- asi que NO es una ventana corta que se
    resetee esperando, sino mas bien un tope de peticiones totales en una
    ventana mas larga. Espaciar peticiones dentro de la misma ejecucion no
    ayuda; lo unico que evita el corte de verdad es no volver a pedir lo
    mismo (ver cargar_datos_liga, que cachea en output/cache/ y solo pide
    lo que aun falte en ejecuciones posteriores).

    Cuando un jugador agota los reintentos rapidos del cliente (ver
    BiwengerClient._get) por un 429, hay dos modos:
    - parar_en_primer_limite=False (por defecto): pausa larga (30s) y
      reintenta ESE MISMO jugador en vez de rendirse y pasar al siguiente
      (que fallaria igual, en cascada). max_pausas_largas limita cuantas
      pausas largas se hacen POR JUGADOR antes de darlo por perdido.
    - parar_en_primer_limite=True: en cuanto se agotan los reintentos
      rapidos del cliente para un jugador, para la descarga entera ahi
      mismo (sin pausas largas ni seguir con el resto) y devuelve lo
      conseguido en esta tanda. Pensado para ir completando la cache poco a
      poco en varias ejecuciones cortas en vez de una tirada larga.

    on_progress: callback opcional que se llama cada `guardar_cada`
    jugadores (y al terminar, si se interrumpe con Ctrl+C, o si se para por
    parar_en_primer_limite) con el dict `detalles` conseguido hasta ese
    momento, para poder guardarlo en cache de forma incremental y no perder
    el progreso.
    """
    detalles = {}
    ids = sorted(player_ids)
    total = len(ids)
    try:
        for i, pid in enumerate(ids, start=1):
            if progreso and (i % 10 == 0 or i == total):
                print(f"    jugador {i}/{total}...")
            pausas = 0
            while True:
                try:
                    detalles[pid] = recortar_temporadas_antiguas(client.get_player_detail(pid))
                    break
                except BiwengerAPIError as e:
                    es_rate_limit = "429" in str(e) or "Demasiados" in str(e)
                    if es_rate_limit and parar_en_primer_limite:
                        print(f"    limite de peticiones alcanzado en el jugador {pid} -- "
                              f"paro aqui ({len(detalles)}/{total} conseguidos en esta tanda).")
                        if on_progress:
                            on_progress(detalles)
                        return detalles
                    if es_rate_limit and pausas < max_pausas_largas:
                        pausas += 1
                        print(f"    (rate-limit sostenido, pausa larga #{pausas} de 30s "
                              f"antes de reintentar el jugador {pid}...)")
                        time.sleep(30)
                        continue
                    print(f"    [!] no se pudo obtener el jugador {pid}: {e}")
                    break
            if on_progress and guardar_cada and i % guardar_cada == 0:
                on_progress(detalles)
    except KeyboardInterrupt:
        if on_progress:
            print(f"    interrumpido -- guardando el progreso conseguido hasta ahora "
                  f"({len(detalles)}/{total} jugadores)...")
            on_progress(detalles)
        raise
    if on_progress:
        on_progress(detalles)
    return detalles


# ---------------------------------------------------------------------------
# Cache local (output/cache/*.json) + orquestacion para el notebook: evita
# volver a pedir todo a la API (con su limite de peticiones) cada vez que
# reejecutas el notebook. Usa forzar_refresco=True cuando quieras datos
# frescos de verdad.
# ---------------------------------------------------------------------------

# archivos de cache que pueden faltar en una cache generada por una version
# anterior de este modulo, sin que eso obligue a re-descargar todo
_CACHE_FILES_OPCIONALES = {"ofertas_raw"}


def _hay_cache():
    return all(
        p.exists() for key, p in _CACHE_FILES.items() if key not in _CACHE_FILES_OPCIONALES
    )


def _guardar_cache(data: dict):
    CACHE_DIR.mkdir(parents=True, exist_ok=True)
    for key, path in _CACHE_FILES.items():
        path.write_text(json.dumps(data[key], ensure_ascii=False), encoding="utf-8")


def _leer_cache():
    data = {}
    for key, path in _CACHE_FILES.items():
        if not path.exists() and key in _CACHE_FILES_OPCIONALES:
            data[key] = []
            continue
        data[key] = json.loads(path.read_text(encoding="utf-8"))
    # los dicts de Python con claves int se guardan como string en JSON
    data["detalles"] = {int(k): v for k, v in data["detalles"].items()}
    return data


def _todos_los_ids(data: dict):
    todos_ids = {f["player_id"] for f in data["mercado_raw"] if f.get("player_id")}
    for r in data["rosters_raw"].values():
        todos_ids |= {f["player_id"] for f in r["filas"] if f.get("player_id")}
    return todos_ids


def _detalles_faltantes(data: dict):
    """ids de jugadores que aparecen en mercado/plantillas pero de los que
    aun no tenemos el detalle descargado en `data['detalles']`."""
    return _todos_los_ids(data) - set(data["detalles"].keys())


def _antiguedad_cache_horas():
    """Horas desde el ultimo guardado de la cache (segun la fecha de
    modificacion de sus ficheros). None si no hay cache."""
    ref = _CACHE_FILES["standings"]
    if not ref.exists():
        return None
    return (time.time() - ref.stat().st_mtime) / 3600


def cargar_datos_liga(email, password, league_id=None, own_team_id=None, forzar_refresco=False,
                       max_antiguedad_horas=24, parar_en_primer_limite=False):
    """
    Punto de entrada principal para el notebook de analisis: descarga (o
    reusa de output/cache/*.json) el mercado, las plantillas de todos los
    rivales, tu plantilla, y el detalle de cada jugador implicado.

    Por defecto reusa la cache local si ya existe Y esta completa Y no tiene
    mas de `max_antiguedad_horas` (24 por defecto) -- pasado ese tiempo se
    refresca sola automaticamente, sin que tengas que acordarte de pasar
    forzar_refresco=True cada dia. Pasa max_antiguedad_horas=None para
    desactivar este chequeo (usar la cache aunque sea vieja, mientras este
    completa). forzar_refresco=True siempre fuerza un refresco, sin mirar
    la antiguedad.

    RESUMIBLE: el detalle de cada jugador (lo mas lento, por el limite de
    peticiones de la API) se guarda en cache de forma incremental segun se
    va consiguiendo -- si esta funcion se interrumpe (Ctrl+C, un corte, etc.)
    a medias, la proxima vez que la llames (sin forzar_refresco) retoma
    donde se quedo: NO vuelve a pedir los jugadores que ya tenia, solo los
    que le faltan. Esto es distinto del refresco automatico por antiguedad:
    una cache INCOMPLETA se retoma siempre (sin importar cuanto tiempo haya
    pasado), porque interrumpida no significa desactualizada.

    parar_en_primer_limite: si True, en cuanto se alcance el limite de
    peticiones de la API (429 sostenido) se sale de la funcion ahi mismo en
    vez de esperar con pausas largas -- lo conseguido hasta ese momento ya
    esta guardado en cache (ver arriba). Util para ir completando la cache
    poco a poco en varias ejecuciones cortas: llamas otra vez mas tarde y
    retoma solo lo que falte.

    En cualquier caso -- completo, incompleto, o parado por limite de
    peticiones -- lo que devuelve esta funcion se relee siempre de la cache
    en disco (no de lo que haya en memoria de esta llamada), porque la
    cache es la combinacion de esta ejecucion con todas las anteriores.
    """
    detalles_previos = {}
    if not forzar_refresco and _hay_cache():
        cache = _leer_cache()
        faltan_cache = _detalles_faltantes(cache)
        if not faltan_cache:
            antiguedad = _antiguedad_cache_horas()
            demasiado_vieja = (
                max_antiguedad_horas is not None
                and antiguedad is not None
                and antiguedad >= max_antiguedad_horas
            )
            if not demasiado_vieja:
                aviso_antiguedad = f" ({antiguedad:.1f}h de antiguedad)" if antiguedad is not None else ""
                print(f"Cargando datos desde cache local ({CACHE_DIR}), completa{aviso_antiguedad}. "
                      "Llama a cargar_datos_liga(..., forzar_refresco=True) para pedir datos frescos.")
                return cache
            print(f"Cache local completa pero tiene {antiguedad:.1f}h de antiguedad "
                  f"(mas de {max_antiguedad_horas}h) -- refrescando automaticamente "
                  "con datos de jugador frescos...")
            # OJO: no reusamos detalles_previos aqui a proposito -- una cache
            # vieja pero completa necesita puntos/precio/forma de verdad
            # actualizados, no solo repetir lo que ya tenia
        else:
            detalles_previos = cache["detalles"]
            print(f"Cache local incompleta: ya hay {len(detalles_previos)} jugadores, "
                  f"faltan {len(faltan_cache)}. Retomando -- no se vuelve a pedir lo ya descargado...")

    print("Descargando datos de la API de Biwenger...")
    client = BiwengerClient(email, password, league_id=league_id)
    login_y_resolver_liga(client, league_id)
    standings, score_id = cargar_liga(client)
    mi_team_id = identificar_mi_equipo(standings, client, own_team_id)

    print("  mercado...")
    market_full = client.get_market()
    mercado_raw = _parse_mercado(market_full)
    ofertas_raw = _parse_ofertas(market_full)
    balance = _dig(market_full, ["status", "balance"])

    print(f"  plantillas de {len(standings)} equipo(s)...")
    rosters_raw = {}
    for s in standings:
        team_id = s.get("id")
        if not team_id:
            continue
        nombre = first_present(s, ["name"], f"equipo:{team_id}")
        es_mio = str(team_id) == str(mi_team_id)
        filas = cargar_plantilla(client, team_id, nombre)
        # el team_id de un standing es el mismo valor que su user_id (el que
        # hace falta como 'to' en place_offer para clausularle un jugador)
        rosters_raw[team_id] = {"team_id": team_id, "es_mio": es_mio, "nombre": nombre, "filas": filas}

    data_base = {
        "standings": standings,
        "score_id": score_id,
        "mi_team_id": mi_team_id,
        "balance": balance,
        "mercado_raw": mercado_raw,
        "rosters_raw": rosters_raw,
        "ofertas_raw": ofertas_raw,
    }

    todos_ids = _todos_los_ids(data_base)
    faltan_ids = todos_ids - set(detalles_previos.keys())
    if detalles_previos:
        print(f"  {len(detalles_previos)} de {len(todos_ids)} jugadores ya en cache; "
              f"pidiendo los {len(faltan_ids)} que faltan...")
    else:
        print(f"  detalle de {len(todos_ids)} jugadores...")

    def guardar_progreso(detalles_parciales):
        data_actual = dict(data_base)
        data_actual["detalles"] = {**detalles_previos, **detalles_parciales}
        _guardar_cache(data_actual)

    nuevos_detalles = enriquecer_con_detalle(
        client, faltan_ids, on_progress=guardar_progreso, parar_en_primer_limite=parar_en_primer_limite,
    )

    data = dict(data_base)
    data["detalles"] = {**detalles_previos, **nuevos_detalles}
    _guardar_cache(data)
    faltan_todavia = len(_detalles_faltantes(data))
    if faltan_todavia:
        print(f"Datos guardados en cache local ({CACHE_DIR}). "
              f"Ojo: aun faltan {faltan_todavia} jugadores -- "
              "vuelve a llamar a esta funcion para reintentar solo esos.")
    else:
        print(f"Datos completos guardados en cache local ({CACHE_DIR}).")
    # siempre releemos de disco: la cache es la combinacion de esta llamada
    # con todas las anteriores, es la fuente de verdad
    return _leer_cache()


def construir_dataframe(data: dict) -> pd.DataFrame:
    """
    Construye el DataFrame combinado (mercado + rivales + tu equipo, cada
    fila con su 'Origen') a partir de lo que devuelve cargar_datos_liga().

    Determina la temporada de liga "actual" por mayoria entre todos los
    jugadores (ver temporada_liga_mas_comun) para que un jugador que lleva
    tiempo sin jugar no arrastre puntos de hace 2+ temporadas disfrazados
    de "esta temporada"/"la anterior".
    """
    detalles = data["detalles"]
    score_id = data["score_id"]
    temporada_actual_id = temporada_liga_mas_comun(detalles)
    temporada_anterior_id = (
        str(int(temporada_actual_id) - 1)
        if temporada_actual_id is not None and temporada_actual_id.isdigit()
        else None
    )
    filas = []

    for f in data["mercado_raw"]:
        # si el jugador tiene dueno (vendedor_id presente), la UNICA via
        # real de ficharlo es el clausulazo -- el "precio de mercado" que
        # trae la API para el no es una compra directa alternativa, es solo
        # informativo (a veces es MENOR que la clausula, a veces MUCHO
        # mayor: no representa un precio de compra real). Ese jugador ya
        # aparece via su fila de Origen 'Rival: X' con la clausula
        # correcta -- no lo dupliques aqui como si fuera comprable suelto.
        if f.get("vendedor_id") is not None:
            continue
        detalle = detalles.get(f["player_id"])
        if not detalle:
            continue
        if jugador_irrelevante(detalle, score_id, temporada_actual_id, temporada_anterior_id):
            continue
        filas.append(construir_fila_jugador(
            f["player_id"], detalle, score_id,
            {
                "Origen": "Mercado",
                "Precio en mercado": f.get("precio_mercado"),
                "Vendedor (id)": f.get("vendedor_id"),
            },
            temporada_actual_id, temporada_anterior_id,
        ))

    for r in data["rosters_raw"].values():
        origen = "Mi equipo" if r["es_mio"] else f"Rival: {r['nombre']}"
        for f in r["filas"]:
            detalle = detalles.get(f["player_id"])
            if not detalle:
                continue
            # el filtro de "jugador irrelevante" NO se aplica a tu propio
            # equipo -- tus jugadores los ves siempre, esten activos o no
            if not r["es_mio"] and jugador_irrelevante(
                detalle, score_id, temporada_actual_id, temporada_anterior_id
            ):
                continue
            pts_actual, partidos_actual, pts_anterior, partidos_anterior, _, _ = (
                extract_seasons_points(detalle, score_id, temporada_actual_id, temporada_anterior_id)
            )
            clausula = f.get("clausula")
            bloqueada_hasta = f.get("clausula_bloqueada_hasta")
            extra = {
                "Origen": origen,
                "Clausula": clausula,
                # dos formas de ver el ratio: por puntos TOTALES acumulados
                # en la temporada, o por puntos MEDIOS (por partido) -- un
                # jugador con pocos partidos jugados puede salir mal en el
                # total pero bien en la media, y viceversa
                "Ratio pts totales/clausula (actual)": ratio(pts_actual, clausula),
                "Ratio pts medios/clausula (actual)": ratio(puntos_por_partido(pts_actual, partidos_actual), clausula),
                "Ratio pts totales/clausula (anterior)": ratio(pts_anterior, clausula),
                "Ratio pts medios/clausula (anterior)": ratio(
                    puntos_por_partido(pts_anterior, partidos_anterior), clausula
                ),
                "Vendedor (id)": None if r["es_mio"] else r.get("team_id"),
                "Clausula bloqueada hasta": (
                    datetime.fromtimestamp(bloqueada_hasta) if bloqueada_hasta else None
                ),
                "Clausula disponible ahora": (
                    not bloqueada_hasta or bloqueada_hasta <= datetime.now().timestamp()
                ),
            }
            filas.append(construir_fila_jugador(
                f["player_id"], detalle, score_id, extra, temporada_actual_id, temporada_anterior_id
            ))

    return pd.DataFrame(filas)


def ofertas(data: dict, tipo="recibidas", solo_identificadas=False) -> pd.DataFrame:
    """
    Ofertas de compra activas en tu mercado, con el nombre del jugador ya
    resuelto. tipo:
    - 'recibidas': ofertas que otros managers han hecho por TUS jugadores
      (util para decidir si aceptar/rechazar desde la app; esta libreria no
      tiene un metodo para aceptar/rechazar, solo para consultarlas).
    - 'enviadas': ofertas que TU has hecho (via place_offer u otro medio).
    - 'todas': sin filtrar.

    IMPORTANTE: Biwenger genera automaticamente una oferta "recibida" para
    CADA jugador de tu plantilla, cercana a su valor de mercado, SIN postor
    real detras (postor 'Desconocido', sin 'de_id') -- es la venta rapida
    de la plataforma, no demanda real de otro manager. Se detecto con
    datos reales: 15/15 jugadores de una plantilla tenian una de estas.
    Pasa solo_identificadas=True para descartarlas y quedarte solo con
    ofertas de un manager identificado (demanda real).
    """
    detalles = data["detalles"]
    mi_team_id = data["mi_team_id"]
    filas = []
    for o in data.get("ofertas_raw", []):
        if tipo == "recibidas" and str(o.get("a_id")) != str(mi_team_id):
            continue
        if tipo == "enviadas" and str(o.get("de_id")) != str(mi_team_id):
            continue
        if solo_identificadas and tipo == "recibidas" and o.get("de_id") is None:
            continue
        pid = o.get("player_id")
        detalle = detalles.get(pid)
        nombre = first_present(detalle, ["name"], f"id:{pid}") if detalle else f"id:{pid}"
        expira = o.get("expira")
        filas.append({
            "offer_id": o.get("offer_id"),
            "player_id": pid,
            "Jugador": nombre,
            "Importe": o.get("importe"),
            "Tipo": o.get("tipo"),
            "Estado": o.get("estado"),
            "De": o.get("de_nombre") or "Desconocido",
            "A": o.get("a_nombre") or "Desconocido",
            "Expira": datetime.fromtimestamp(expira) if expira else None,
        })
    return pd.DataFrame(filas)


# ---------------------------------------------------------------------------
# Historial (output/history/<timestamp>/*.json): copias fechadas de los
# datos, para poder comparar la evolucion de tu equipo/el mercado con el
# tiempo (la cache de output/cache/ en cambio se sobrescribe cada refresco).
# ---------------------------------------------------------------------------

def guardar_snapshot(data: dict):
    """Guarda una copia fechada de `data` (lo que devuelve cargar_datos_liga)
    en output/history/<timestamp>/. Llamalo justo despues de un
    forzar_refresco=True si quieres poder comparar mas adelante."""
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    carpeta = HISTORY_DIR / ts
    carpeta.mkdir(parents=True, exist_ok=True)
    for key in _CACHE_FILES:
        (carpeta / f"{key}.json").write_text(json.dumps(data[key], ensure_ascii=False), encoding="utf-8")
    print(f"Snapshot guardado en {carpeta}")
    return carpeta


def listar_snapshots():
    """Nombres (timestamp) de los snapshots guardados, mas antiguo primero."""
    if not HISTORY_DIR.exists():
        return []
    return sorted(p.name for p in HISTORY_DIR.iterdir() if p.is_dir())


def cargar_snapshot(nombre: str) -> dict:
    """Carga un snapshot guardado por guardar_snapshot() (usa listar_snapshots()
    para ver los nombres disponibles). Devuelve el mismo formato que
    cargar_datos_liga(), listo para pasar a construir_dataframe()."""
    carpeta = HISTORY_DIR / nombre
    data = {}
    for key in _CACHE_FILES:
        data[key] = json.loads((carpeta / f"{key}.json").read_text(encoding="utf-8"))
    data["detalles"] = {int(k): v for k, v in data["detalles"].items()}
    return data
