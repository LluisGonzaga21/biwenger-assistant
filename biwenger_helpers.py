"""
Shared functions for loading and analyzing Biwenger data (market, rival
rosters, your roster, points/value ratios, recent form, and signing
suggestions).

Used by both 02_generar_excel.py and the analysis notebook
(03_analisis.ipynb) to avoid duplicating logic. If you change something
here, it affects both.

IMPORTANT about 'score_id': Biwenger computes each player's points in
parallel under several different scoring systems (1, 2, 3...). Both
'seasons[].points' and 'reports[].points' are dicts {scoreID: pts}, NOT a
plain number. You need to use YOUR league's scoreID (league['scoreID'],
visible in output/debug/league.json) to read the one that actually counts
toward your standings. That's why every function here that reads points
takes score_id explicitly.
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

POSICIONES = {1: "Goalkeeper", 2: "Defender", 3: "Midfielder", 4: "Forward"}

# Player 'status' values that mean they are not available to play. Confirmed
# with real API data: "ok", "doubt", "injured", "sanctioned", "discarded".
# "doubt" (last-minute doubt) is deliberately left OUT -- it usually means
# they might play, not that they definitely won't, so it is not
# automatically excluded from the recommenders (it's still shown in the
# 'Status' column so you can decide for yourself). If you see another new
# value in output/cache/detalles.json, add it here.
ESTADOS_NO_DISPONIBLE = {"injured", "sanctioned", "discarded"}

# ---------------------------------------------------------------------------
# Candidate field names per concept. If something comes out empty in the
# analysis, this is the first place to check/edit after reviewing
# output/debug/*.json (run 01_explorar_api.py to generate them).
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


def points_by_score(points_dict, score_id):
    """Extracts the points value for your score_id from a dict {scoreID: pts}."""
    if not isinstance(points_dict, dict):
        return None
    return points_dict.get(str(score_id))


def points_per_match(puntos, partidos):
    """Points per match played. None if there are no matches or points."""
    if puntos is None or not partidos:
        return None
    try:
        return round(puntos / partidos, 2)
    except (TypeError, ZeroDivisionError):
        return None


def _recent_league_seasons(seasons, n=2):
    """
    From a 'seasons' list (as given by the API), returns at most the `n`
    most recent LEAGUE ones (excludes cup/champions, which appear as
    entries with 'competition'), ordered from most recent to oldest.
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


def trim_old_seasons(detalle):
    """
    Trims 'seasons' in a player's detail down to just the current and
    previous LEAGUE seasons -- that's all the analysis uses (see
    extract_seasons_points); anything older, plus cup/champions, is
    discarded. Applied before saving to cache so the full history isn't
    carried around unnecessarily. Returns a new detail (does not mutate the
    original).
    """
    seasons = detalle.get("seasons") if isinstance(detalle, dict) else None
    if not isinstance(seasons, list):
        return detalle
    nuevo = dict(detalle)
    nuevo["seasons"] = _recent_league_seasons(seasons, n=2)
    return nuevo


def most_common_league_season(detalles: dict):
    """
    Id of the truly 'current' LEAGUE season, inferred by MAJORITY VOTE: the
    most recent season each player has on record, the one most repeated
    across all of them. In an active league, most of the ~200 players will
    have played this season, so their mode is a reliable reference for
    what "now" is -- better than blindly trusting the most recent season of
    EACH player individually, which for someone who hasn't played in a
    while (long injury, no team, etc.) can be from 2+ seasons ago.
    """
    from collections import Counter
    ids = []
    for detalle in detalles.values():
        recientes = _recent_league_seasons(
            detalle.get("seasons") if isinstance(detalle, dict) else None, n=1
        )
        if recientes:
            ids.append(str(recientes[0].get("id")))
    if not ids:
        return None
    return Counter(ids).most_common(1)[0][0]


def extract_seasons_points(detalle, score_id, temporada_actual_id=None, temporada_anterior_id=None):
    """
    Looks in 'seasons' for the current and previous LEAGUE season (excludes
    cup/champions, which appear as entries with 'competition') and returns:
    (pts_actual, partidos_actual, pts_anterior, partidos_anterior,
    etiqueta_actual, etiqueta_anterior).

    If temporada_actual_id / temporada_anterior_id are passed (see
    most_common_league_season), it requires the league season to have
    exactly that id to count as "current"/"previous". Without this, a
    player who hasn't played in a while (their most recent record is from
    2+ seasons ago, e.g. a long injury or no team) would have those old
    points attributed to them as if they were from this season or the last
    one -- that would completely distort their Puntuacion potencial.
    Without these parameters, it simply uses the player's 2 most recent
    league seasons (old behavior, less reliable for cases like that).
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
        ordered = _recent_league_seasons(seasons)
        actual = ordered[0] if len(ordered) > 0 else None
        anterior = ordered[1] if len(ordered) > 1 else None

    if actual is None and anterior is None:
        return None, None, None, None, None, None

    return (
        points_by_score(actual.get("points"), score_id) if actual else None,
        actual.get("games") if actual else None,
        points_by_score(anterior.get("points"), score_id) if anterior else None,
        anterior.get("games") if anterior else None,
        label(actual),
        label(anterior),
    )


def calculate_form(detalle, score_id, decay=0.85, n_partidos=6):
    """
    Weighted average of the points from the last matches played this
    season, giving more weight to the most recent ones (based on
    'reports', which provides one result per match with a date).

    decay (0-1): the lower it is, the more the last match is prioritized
    over ones from a few matchdays back. n_partidos: how many recent
    matches go into the average.
    """
    reports = detalle.get("reports") if isinstance(detalle, dict) else None
    if not isinstance(reports, list) or not reports:
        return None

    partidos = []
    for r in reports:
        if not isinstance(r, dict):
            continue
        pts = points_by_score(r.get("points"), score_id)
        fecha = _dig(r, ["match", "date"])
        if pts is None or fecha is None:
            continue
        partidos.append((fecha, pts))
    if not partidos:
        return None

    partidos.sort(key=lambda t: t[0])  # oldest -> most recent
    partidos = partidos[-n_partidos:]

    pesos = [decay ** i for i in range(len(partidos) - 1, -1, -1)]  # last match -> maximum weight
    total_peso = sum(pesos)
    return round(sum(p * w for (_, p), w in zip(partidos, pesos)) / total_peso, 2)


def potential_score(detalle, score_id, decay=0.85, n_partidos=6, peso_temporada_anterior=0.3,
                          temporada_actual_id=None, temporada_anterior_id=None):
    """
    "Potential" performance metric designed to compare signing candidates:
    it blends recent form (matches from this season, weighting the latest
    ones more) with points/match from the previous LaLiga season (weighted
    less, and only if they played in the top division). If either
    component is missing, it just uses whichever one is available.
    """
    forma = calculate_form(detalle, score_id, decay=decay, n_partidos=n_partidos)
    _, _, pts_anterior, partidos_anterior, _, _ = extract_seasons_points(
        detalle, score_id, temporada_actual_id, temporada_anterior_id
    )
    ppg_anterior = points_per_match(pts_anterior, partidos_anterior)

    if forma is None and ppg_anterior is None:
        return None
    if forma is None:
        return ppg_anterior
    if ppg_anterior is None:
        return forma
    return round(forma * (1 - peso_temporada_anterior) + ppg_anterior * peso_temporada_anterior, 2)


def is_available(detalle):
    """False if the player's 'status' is in ESTADOS_NO_DISPONIBLE (injury,
    suspension...). True if it's 'ok' or if there's no status (better not
    to exclude by mistake)."""
    estado = detalle.get("status") if isinstance(detalle, dict) else None
    if not estado:
        return True
    return str(estado).lower() not in ESTADOS_NO_DISPONIBLE


def next_match_difficulty(detalle):
    """
    Info about the player's next LaLiga match based on 'team.nextGames[0]':
    opponent, whether it's home, the matchday, and a difficulty rating
    (0-100, higher = harder for their team). None if there is no upcoming
    match scheduled.
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


def price_trend(detalle, dias=7):
    """
    % change in market value over the last `dias` points of the daily price
    history ('prices': chronological list of [date, price]). Positive =
    rising, negative = falling. None if there isn't enough history.
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


PARTIDOS_TEMPORADA_COMPLETA = 38  # matchdays in a 20-team league (home and away)
VALOR_MINIMO_RELEVANTE = 400_000  # below this, considered "clutter"


def is_irrelevant_player(detalle, score_id, temporada_actual_id=None, temporada_anterior_id=None):
    """
    True if the player adds nothing useful to the market/rivals analysis:
    - unknown market value or <= VALOR_MINIMO_RELEVANTE ("clutter"), or
    - not registered with any top-division team ('team' empty in the API --
      see the Owono case), or
    - hasn't debuted this season (no 'Points current season') AND played
      less than half the matches last season (essentially inactive player
      or out of the first team).
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
        return False  # has debuted this season
    return (partidos_anterior or 0) < (PARTIDOS_TEMPORADA_COMPLETA / 2)


def ratio(points, value):
    """Points per million of value (market value or buyout clause). None if it can't be computed."""
    if points is None or not value:
        return None
    try:
        return round(points / (value / 1_000_000), 2)
    except (TypeError, ZeroDivisionError):
        return None


def build_player_row(pid, detalle, score_id, extra=None,
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
    proximo = next_match_difficulty(detalle) or {}

    fila = {
        "player_id": pid,
        "Player": nombre,
        "Position": posicion,
        "LaLiga Team": equipo,
        "Status": first_present(detalle, ["status"], "ok"),
        "Available": is_available(detalle),
        "Market value": valor_actual,
        "Price trend (7d) %": price_trend(detalle),
        # column names are FIXED on purpose (not the literal season name,
        # e.g. "Temporada 2025/2026") -- otherwise players without a
        # previous season (debutants, signings from abroad) would generate
        # different columns ("Points prev. season") and the df would fill
        # up with near-duplicate columns. 'Current season'/'Previous season' carry
        # the real label if you ever need to check it.
        "Current season": etq_actual,
        "Points current season": pts_actual,
        "Matches current season": partidos_actual,
        "Pts/match (current)": points_per_match(pts_actual, partidos_actual),
        "Ratio pts/MV (current)": ratio(pts_actual, valor_actual),
        "Previous season": etq_anterior,
        "Points previous season": pts_anterior,
        "Matches previous season": partidos_anterior,
        "Pts/match (previous)": points_per_match(pts_anterior, partidos_anterior),
        "Ratio pts/MV (previous)": ratio(pts_anterior, valor_actual),
        "Form (last matches)": calculate_form(detalle, score_id),
        "Potential score": potential_score(
            detalle, score_id, temporada_actual_id=temporada_actual_id, temporada_anterior_id=temporada_anterior_id
        ),
        "Next opponent": proximo.get("rival"),
        "Home/Away": ("Home" if proximo.get("es_local") else "Away") if proximo else None,
        "Next match difficulty": proximo.get("dificultad"),
        "Next matchday": proximo.get("jornada"),
    }
    fila.update(extra)
    return fila


def suggest_signings(df: pd.DataFrame, balance, top_n=15, solo_disponibles=True, ordenar_por="mejora"):
    """
    Suggests signings (a buyout clause on a rival's player, or a market
    purchase) that would improve your team, comparing 1 to 1: each
    affordable candidate against your weakest player (by 'Puntuacion
    potencial') in the same position.

    df: the combined DataFrame from build_dataframe() (needs the
    columns 'Source', 'Position', 'Potential score' and, for the cost,
    'Market price' and/or 'Clause').
    balance: your available balance (maximum amount to spend).
    solo_disponibles: if True (default), discards injured or suspended
    candidates ('Available' == False).
    ordenar_por: 'mejora' (default, absolute improvement in pts/match) or
    'eficiencia' (improvement per million spent -- prioritizes cheap
    signings that improve you a little over very expensive signings that
    improve you a lot).

    Always discards rival buyout clauses that are currently locked (after a
    recent purchase/clause buyout, the API would reject it) -- see
    'Clause locked until' in the df.
    """
    requeridas = {"Source", "Position", "Potential score"}
    faltan = requeridas - set(df.columns)
    if faltan:
        raise ValueError(f"The DataFrame is missing columns: {faltan}")
    if ordenar_por not in ("mejora", "eficiencia"):
        raise ValueError("ordenar_por must be 'mejora' or 'eficiencia'")

    mi_equipo = df[(df["Source"] == "My team") & df["Potential score"].notna()]
    if mi_equipo.empty:
        raise ValueError("There are no 'My team' players with a computable Puntuacion potencial.")

    mi_equipo_ordenado = mi_equipo.sort_values("Potential score")
    listones = mi_equipo_ordenado.groupby("Position")["Potential score"].min()
    peor_jugador = mi_equipo_ordenado.groupby("Position")["Player"].first()
    # how many of your players in each position perform below your own
    # median there -- more than one indicates a position with more than one
    # weak link, not just the worst one
    medianas_propias = mi_equipo.groupby("Position")["Potential score"].median()
    total_propios = mi_equipo.groupby("Position")["Player"].count()
    flojos_propios = (
        mi_equipo[mi_equipo["Potential score"] <= mi_equipo["Position"].map(medianas_propias)]
        .groupby("Position")["Player"].count()
    )

    candidatos = df[df["Source"] != "My team"].copy()
    if solo_disponibles and "Available" in candidatos.columns:
        candidatos = candidatos[candidatos["Available"]]
    if "Clause available now" in candidatos.columns:
        # Mercado rows don't have this column (NaN) -- they're let through,
        # only rival clauses EXPLICITLY locked right now are discarded
        candidatos = candidatos[candidatos["Clause available now"] != False]  # noqa: E712
    coste_mercado = candidatos["Market price"] if "Market price" in candidatos else None
    coste_clausula = candidatos["Clause"] if "Clause" in candidatos else None
    if coste_mercado is not None and coste_clausula is not None:
        candidatos["Signing cost"] = coste_mercado.combine_first(coste_clausula)
    else:
        candidatos["Signing cost"] = coste_mercado if coste_mercado is not None else coste_clausula

    candidatos = candidatos[candidatos["Signing cost"].notna() & (candidatos["Signing cost"] <= balance)]
    candidatos = candidatos[candidatos["Potential score"].notna()]
    candidatos = candidatos[candidatos["Position"].isin(listones.index)]

    candidatos["Your baseline in that position"] = candidatos["Position"].map(listones)
    candidatos["Player you would replace"] = candidatos["Position"].map(peor_jugador)
    candidatos["Estimated improvement (pts/match)"] = (
        candidatos["Potential score"] - candidatos["Your baseline in that position"]
    )
    candidatos = candidatos[candidatos["Estimated improvement (pts/match)"] > 0]

    candidatos["Balance after signing"] = balance - candidatos["Signing cost"]
    candidatos["Improvement per million spent"] = (
        candidatos["Estimated improvement (pts/match)"] / (candidatos["Signing cost"] / 1_000_000)
    ).round(3)
    candidatos["Weak players in your position"] = (
        candidatos["Position"].map(flojos_propios).fillna(0).astype(int)
    )
    candidatos["Total in your position"] = candidatos["Position"].map(total_propios).fillna(0).astype(int)

    columna_orden = "Improvement per million spent" if ordenar_por == "eficiencia" else "Estimated improvement (pts/match)"

    columnas = [
        "player_id", "Player", "Position", "LaLiga Team", "Source", "Signing cost",
        "Balance after signing", "Seller (id)", "Potential score", "Player you would replace",
        "Your baseline in that position", "Estimated improvement (pts/match)", "Improvement per million spent",
        "Weak players in your position", "Total in your position",
    ]
    return (
        candidatos[columnas]
        .sort_values(columna_orden, ascending=False)
        .head(top_n)
        .reset_index(drop=True)
    )


def unique_by_player(df: pd.DataFrame) -> pd.DataFrame:
    """
    A player can appear in more than one row from build_dataframe: one
    for being on a team's roster, another for also being listed for sale on
    the market by their own manager (see the note in section 2 of the
    notebook). That's correct for buy/sell tables (Mercado vs Rival have
    different costs -- market price vs buyout clause), but in a RANKING it
    makes the same player take up several spots instead of leaving room for
    others.

    This function collapses those rows into one per player, combining the
    different 'Source' values into a single text (e.g. 'Mercado + Rival:
    Botas FC'). Use it before a top-N when you don't care where they can be
    obtained from, only the performance ranking.
    """
    if df.empty or "player_id" not in df.columns:
        return df
    otras_cols = [c for c in df.columns if c not in ("player_id", "Source")]
    combinado = df.groupby("player_id", as_index=False).agg({
        **{c: "first" for c in otras_cols},
        "Source": lambda s: " + ".join(sorted(set(s))),
    })
    return combinado[df.columns.tolist()]


def cost_per_point(df: pd.DataFrame, metrica="Potential score", margen_pct=20) -> pd.DataFrame:
    """
    For each player (any position/origin), how much it costs (Valor de
    mercado) per point of performance (`metrica`) they generate, and how
    that compares against the MEDIAN of THE WHOLE GAME -- all valid players
    in your dataset, without distinguishing by position (unlike
    calculate_reference_ratio/estimate_fair_value, which do split by
    position). Lower = cheaper per point = better.

    Use unique_by_player() first if you want a ranking without duplicates
    for players who are both on the market and on a roster.

    Players with metrica <= 0 (a negative or infinite cost per point makes
    no sense) or without a market value are excluded.
    """
    validos = df[df[metrica].notna() & (df[metrica] > 0) & df["Market value"].notna() & (df["Market value"] > 0)].copy()
    validos["Cost per point"] = (validos["Market value"] / validos[metrica]).round(0)

    media_juego = validos["Cost per point"].median()
    validos["Game average (reference)"] = media_juego
    validos["% vs game average"] = ((validos["Cost per point"] - media_juego) / media_juego * 100).round(1)

    def label(pct):
        if pct <= -margen_pct:
            return "Well below (cheap)"
        if pct >= margen_pct:
            return "Well above (expensive)"
        return "Average"

    validos["Rating"] = validos["% vs game average"].apply(label)

    columnas = [c for c in [
        "player_id", "Player", "Position", "Source", "Market value", metrica,
        "Cost per point", "Game average (reference)", "% vs game average", "Rating",
    ] if c in validos.columns]
    return validos[columnas].sort_values("Cost per point").reset_index(drop=True)


def best_players(df: pd.DataFrame, posicion=None, origen=None, top_n=20,
                       metrica="Potential score", solo_disponibles=True):
    """
    General player ranking by performance metric (default 'Puntuacion
    potencial'), without looking at price or whether you can afford it.
    Works as a watchlist: who's in the best form right now.

    posicion: filters by 'Goalkeeper'/'Defender'/'Midfielder'/'Forward'.
    origen: filters by 'Market', 'My team', or a rival's exact 'Source'
    (e.g. 'Rival: Botas FC'). By default it doesn't filter (includes all),
    and in that case it collapses players duplicated across several
    origins (see unique_by_player) so they don't take up several spots in
    the ranking.
    solo_disponibles: if True (default), discards injured/suspended players.
    """
    resultado = df.dropna(subset=[metrica])
    if solo_disponibles and "Available" in resultado.columns:
        resultado = resultado[resultado["Available"]]
    if posicion:
        resultado = resultado[resultado["Position"] == posicion]
    if origen:
        resultado = resultado[resultado["Source"] == origen]
    else:
        resultado = unique_by_player(resultado)

    columnas = [c for c in [
        "player_id", "Player", "Position", "LaLiga Team", "Source",
        metrica, "Form (last matches)", "Next opponent", "Next match difficulty",
        "Market value", "Price trend (7d) %", "Clause", "Market price",
    ] if c in resultado.columns]

    return (
        resultado[columnas]
        .sort_values(metrica, ascending=False)
        .head(top_n)
        .reset_index(drop=True)
    )


def calculate_reference_ratio(df: pd.DataFrame, metrica="Potential score"):
    """
    Reference ratio (metric / market value in millions) per position,
    computed with the MEDIAN of all valid players (market + rivals + your
    team -> broad sample). This is the basis for estimating a "fair"
    market value from performance.
    """
    tmp = df.dropna(subset=[metrica, "Market value"])
    tmp = tmp[tmp["Market value"] > 0].copy()
    tmp["_ratio"] = tmp[metrica] / (tmp["Market value"] / 1_000_000)
    return tmp.groupby("Position")["_ratio"].median()


def estimate_fair_value(df: pd.DataFrame, metrica="Potential score", margen_pct=20, solo_disponibles=True):
    """
    For players FOR SALE on the market, estimates a "fair value" based on
    how the rest of the players in your league (in that same position)
    relate their performance to their market value (see
    calculate_reference_ratio), and compares it with the current asking
    price to spot bargains/overpriced players and suggest an offer.

    This is a heuristic based on similar players in your own league, NOT
    the "real" price according to Biwenger -- use it as guidance, not
    absolute truth. margen_pct: the % difference threshold above/below
    which it's labeled 'Bargain'/'Expensive' (below that it's considered 'Precio
    justo').
    solo_disponibles: if True (default), discards injured/suspended players
    (a bargain that can't play isn't much of a bargain).
    """
    ratios_ref = calculate_reference_ratio(df, metrica=metrica)

    mercado = df[df["Source"] == "Market"].dropna(subset=[metrica, "Market price"]).copy()
    if solo_disponibles and "Available" in mercado.columns:
        mercado = mercado[mercado["Available"]]
    mercado["Position reference ratio"] = mercado["Position"].map(ratios_ref)
    mercado = mercado.dropna(subset=["Position reference ratio"])
    mercado["Estimated fair value"] = (
        mercado[metrica] / mercado["Position reference ratio"] * 1_000_000
    ).round(-3)
    mercado["Difference vs asking price"] = mercado["Estimated fair value"] - mercado["Market price"]
    mercado["% over asking price"] = (
        mercado["Difference vs asking price"] / mercado["Market price"] * 100
    ).round(1)
    # it never makes sense to offer more than the asking price, nor more
    # than what we consider their fair value
    mercado["Suggested offer"] = mercado[["Estimated fair value", "Market price"]].min(axis=1)

    def label(pct):
        if pct >= margen_pct:
            return "Bargain"
        if pct <= -margen_pct:
            return "Expensive"
        return "Fair price"

    mercado["Rating"] = mercado["% over asking price"].apply(label)

    columnas = [
        "player_id", "Player", "Position", "LaLiga Team", "Market price",
        "Estimated fair value", "Difference vs asking price", "% over asking price",
        "Rating", "Suggested offer", "Seller (id)",
    ]
    return (
        mercado[columnas]
        .sort_values("% over asking price", ascending=False)
        .reset_index(drop=True)
    )


def suggest_market_offer(df: pd.DataFrame, data: dict, margen_sobre_pedido_pct=5,
                            metrica="Potential score", rango_comparables_pct=40):
    """
    For each FREE-AGENT player on the market (Origen == 'Market'; doesn't
    apply to buyout clauses -- the amount there is fixed, not bid on),
    estimates how much you'd actually need to bid to land them, not just
    how much they're "worth" per your own heuristic.

    Biwenger doesn't allow bidding below the current asking price, so that
    price is always the FLOOR of the recommended offer. From there:

    1. Real comparables: the MEDIAN of what the current owners of similar
       players (same position, market value within +/-rango_comparables_pct%)
       actually paid for them (real purchase price, not the buyout clause).
       If that median exceeds the asking price, it's the best reference for
       "what it actually takes in practice" -- it's used as the base
       instead of the asking price.
    2. Real competition: how many RIVALS have the budget (`maximumBid` from
       standings -- note this can be higher than their current balance,
       Biwenger allows bidding into the red up to that limit) to beat that
       base. If there's competition, it bumps it up a bit more
       (margen_sobre_pedido_pct) so you don't lose them over a small
       difference.

    Your estimated fair value (same heuristic as estimate_fair_value) is
    included only as an informative reference -- so you can see whether
    what it takes to bid is out of line with what you think the player is
    worth -- but it no longer caps the recommended offer from below (a
    player may require bidding above your fair value and still be the only
    real way to sign them).

    This is a guiding heuristic (we don't know who is REALLY interested in
    each player, only who could afford them financially) -- not a
    guaranteed prediction of what it will take to win them.
    """
    ratios_ref = calculate_reference_ratio(df, metrica=metrica)

    mercado = df[df["Source"] == "Market"].dropna(subset=[metrica, "Market price"]).copy()
    mercado["Position reference ratio"] = mercado["Position"].map(ratios_ref)
    mercado = mercado.dropna(subset=["Position reference ratio"])
    mercado["Estimated fair value"] = (
        mercado[metrica] / mercado["Position reference ratio"] * 1_000_000
    ).round(-3)

    # REAL purchase prices (what each current owner actually paid, not the buyout clause)
    precios_reales = []
    for r in data["rosters_raw"].values():
        for f in r["filas"]:
            precio = f.get("precio_roster")
            if precio:
                precios_reales.append({"player_id": f["player_id"], "precio_pagado": precio})
    precios_reales = pd.DataFrame(precios_reales)
    comparables = df.merge(precios_reales, on="player_id", how="inner") if not precios_reales.empty else pd.DataFrame()

    def comparable_price(row):
        if comparables.empty:
            return None
        rango = row["Market value"] * rango_comparables_pct / 100
        similares = comparables[
            (comparables["Position"] == row["Position"])
            & (comparables["Market value"] >= row["Market value"] - rango)
            & (comparables["Market value"] <= row["Market value"] + rango)
        ]
        if similares.empty:
            return None
        return similares["precio_pagado"].median()

    mercado["Price paid by comparables"] = mercado.apply(comparable_price, axis=1)

    # how many rivals (not you) have enough budget spare to beat the asking price
    mi_team_id = data["mi_team_id"]
    max_bids_rivales = [
        s.get("maximumBid") for s in data["standings"]
        if str(s.get("id")) != str(mi_team_id) and s.get("maximumBid") is not None
    ]

    def rivals_who_could_bid_more(precio_pedido):
        return sum(1 for mb in max_bids_rivales if mb > precio_pedido)

    mercado["Rivals who could bid more"] = mercado["Market price"].apply(rivals_who_could_bid_more)

    def recommended_offer(row):
        # floor: never below the asking price -- Biwenger won't allow a lower bid
        base = row["Market price"]
        comparable = row["Price paid by comparables"]
        if pd.notna(comparable) and comparable > base:
            base = comparable
        if row["Rivals who could bid more"] > 0:
            base = base * (1 + margen_sobre_pedido_pct / 100)
        return round(base / 1000) * 1000

    mercado["Recommended offer"] = mercado.apply(recommended_offer, axis=1)

    columnas = [
        "player_id", "Player", "Position", "LaLiga Team", "Market price",
        "Estimated fair value", "Price paid by comparables", "Rivals who could bid more",
        "Recommended offer", "Seller (id)",
    ]
    return (
        mercado[columnas]
        .sort_values("Recommended offer", ascending=False)
        .reset_index(drop=True)
    )


def suggest_sales(df: pd.DataFrame, data: dict = None, metrica="Potential score", margen_pct=20):
    """
    The counterpart to estimate_fair_value but for YOUR roster: compares
    the OFFICIAL market value of each of your players with what their
    current performance would justify (same reference median by position).
    If the market value is much higher than that "fair value", the player
    is overvalued -- selling now takes advantage of a price that probably
    won't hold if their performance doesn't improve.

    If you pass `data` (what load_league_data returns), it also cross-
    references the offers you've received from an IDENTIFIED manager (see
    offers(solo_identificadas=True) -- discards the automatic quick-sale
    offer Biwenger generates for every player on your roster, which isn't
    real demand): a real offer in hand outweighs the heuristic -- if an
    identified manager already offers at or above the estimated fair
    value, selling is recommended even if the overvaluation % alone
    wouldn't reach the threshold.

    Same as estimate_fair_value, this is a guiding heuristic based on your
    own league, not a guaranteed prediction.
    """
    ratios_ref = calculate_reference_ratio(df, metrica=metrica)

    mios = df[df["Source"] == "My team"].dropna(subset=[metrica, "Market value"]).copy()
    mios["Position reference ratio"] = mios["Position"].map(ratios_ref)
    mios = mios.dropna(subset=["Position reference ratio"])
    mios["Estimated fair value"] = (
        mios[metrica] / mios["Position reference ratio"] * 1_000_000
    ).round(-3)
    mios["Difference (MV - fair)"] = mios["Market value"] - mios["Estimated fair value"]
    mios["% overvaluation"] = (
        mios["Difference (MV - fair)"] / mios["Estimated fair value"] * 100
    ).round(1)

    mios["Best offer received"] = None
    mios["Offers received"] = 0
    if data is not None:
        recibidas = offers(data, tipo="recibidas", solo_identificadas=True)
        if not recibidas.empty:
            mejor_oferta = recibidas.groupby("player_id")["Amount"].max()
            n_ofertas = recibidas.groupby("player_id")["Amount"].count()
            mios["Best offer received"] = mios["player_id"].map(mejor_oferta)
            mios["Offers received"] = mios["player_id"].map(n_ofertas).fillna(0).astype(int)

    def label(row):
        oferta = row["Best offer received"]
        if pd.notna(oferta) and oferta >= row["Estimated fair value"]:
            return "Sell now (you have an offer at or above fair value)"
        if row["% overvaluation"] >= margen_pct:
            return "Sell now"
        if row["% overvaluation"] <= -margen_pct:
            return "Undervalued (don't sell)"
        return "Price matches performance"

    mios["Recommendation"] = mios.apply(label, axis=1)

    columnas = [
        "player_id", "Player", "Position", "Market value", "Clause",
        "Estimated fair value", "Difference (MV - fair)", "% overvaluation",
        "Best offer received", "Offers received", "Recommendation",
    ]
    columnas = [c for c in columnas if c in mios.columns]
    return (
        mios[columnas]
        .sort_values(["Best offer received", "% overvaluation"], ascending=False, na_position="last")
        .reset_index(drop=True)
    )


def login_and_resolve_league(client: BiwengerClient, league_id=None):
    """Logs in and detects the league + user_id (team-id) for that league.

    Raises RuntimeError if it couldn't resolve any league (for example, if
    BIWENGER_LEAGUE_ID doesn't match any league on your account and no
    league could be found automatically either).
    """
    client.login()
    leagues, match = client.resolve_league(league_id)
    if not client.league_id:
        raise RuntimeError(
            "Could not detect your league automatically. Add "
            "BIWENGER_LEAGUE_ID to your .env with your league's ID."
        )
    if not league_id and len(leagues) > 1:
        print("  You have several leagues, using the first one. If it's not "
              "the one you want, set BIWENGER_LEAGUE_ID in your .env:")
        for l in leagues:
            print(f"    - {l.get('name')}: id={l.get('id')}")
    return leagues, match


def load_league(client: BiwengerClient):
    """Returns (standings, score_id) for your league."""
    league = client.get_league()
    standings = league.get("standings") or []
    if not standings:
        raise RuntimeError(
            "Could not find 'standings' in the league response. "
            "Run 01_explorar_api.py and check output/debug/league.json."
        )
    score_id = league.get("scoreID")
    if score_id is None:
        raise RuntimeError(
            "Could not find 'scoreID' in the league response. "
            "Check output/debug/league.json."
        )
    return standings, score_id


def identify_my_team(standings, client: BiwengerClient, own_team_id=None):
    if own_team_id:
        return own_team_id
    for s in standings:
        uid = first_present(s, ["id"]) or first_present(s.get("user", {}) if isinstance(s.get("user"), dict) else {}, ["id"])
        user_obj = s.get("user") if isinstance(s.get("user"), dict) else {}
        if str(uid) == str(client.user_id) or str(user_obj.get("id")) == str(client.user_id):
            return s.get("id")
    print(
        "  [!] Could not automatically identify your own team among the "
        "rivals. Pass own_team_id (your team id, visible in "
        "output/debug/sample_team.json or league.json) to force it."
    )
    return None


def load_squad(client: BiwengerClient, team_id, nombre_equipo):
    team = client.get_team(team_id)
    jugadores = team.get("players") if isinstance(team, dict) else None
    filas = []
    if isinstance(jugadores, list):
        for p in jugadores:
            if not isinstance(p, dict):
                continue
            # the purchase price and the buyout clause are NOT at the top
            # level of the player, but nested inside 'owner' (see
            # output/debug/sample_team.json)
            owner = p.get("owner") if isinstance(p.get("owner"), dict) else {}
            filas.append(
                {
                    "player_id": p.get("id"),
                    "equipo_rival": nombre_equipo,
                    "precio_roster": first_present(owner, CAMPOS_CANDIDATOS["roster_price"]),
                    "clausula": first_present(owner, CAMPOS_CANDIDATOS["roster_clause"]),
                    # deadline (epoch) until which that buyout clause is
                    # locked (typical right after a recent purchase/buyout);
                    # while it lasts, the API would reject a place_offer on
                    # this player
                    "clausula_bloqueada_hasta": owner.get("clauseLockedUntil"),
                }
            )
    return filas


def _parse_market(market: dict):
    lista = None
    for key in CAMPOS_CANDIDATOS["market_list"]:
        v = market.get(key) if isinstance(market, dict) else None
        if isinstance(v, list):
            lista = v
            break
    if lista is None:
        print("  [!] Could not find the market player list. "
              "Check output/debug/market.json after running 01_explorar_api.py.")
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


def load_market(client: BiwengerClient):
    return _parse_market(client.get_market())


def _parse_offers(market: dict):
    """Extracts the 'raw' list of market offers (buy/sell, both received
    and sent). See offers() to filter them."""
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


def enrich_with_details(client: BiwengerClient, player_ids, progreso=True, max_pausas_largas=6,
                            on_progress=None, guardar_cada=10, parar_en_primer_limite=False):
    """
    cf.biwenger.com cuts you off with a rate limit after ~200 requests in a
    row. We tried spacing them out with a preventive pause halfway through,
    but the cutoff kept hitting at the exact same point (player ~200) even
    after a 5-minute pause -- so it's NOT a short window that resets by
    waiting, but rather more like a cap on total requests within a longer
    window. Spacing out requests within the same run doesn't help; the only
    thing that really avoids the cutoff is not requesting the same thing
    again (see load_league_data, which caches to output/cache/ and only
    requests what's still missing on later runs).

    When a player exhausts the client's fast retries (see
    BiwengerClient._get) due to a 429, there are two modes:
    - parar_en_primer_limite=False (default): a long pause (30s) and it
      retries THAT SAME player instead of giving up and moving to the next
      one (which would fail the same way, in a cascade). max_pausas_largas
      limits how many long pauses happen PER PLAYER before giving up on
      them.
    - parar_en_primer_limite=True: as soon as the client's fast retries are
      exhausted for a player, it stops the whole download right there (no
      long pauses, doesn't continue with the rest) and returns what was
      obtained in this batch. Meant for gradually completing the cache
      across several short runs instead of one long run.

    on_progress: optional callback called every `guardar_cada` players (and
    at the end, if interrupted with Ctrl+C, or if it stops due to
    parar_en_primer_limite) with the `detalles` dict obtained so far, so it
    can be saved to cache incrementally without losing progress.
    """
    detalles = {}
    ids = sorted(player_ids)
    total = len(ids)
    try:
        for i, pid in enumerate(ids, start=1):
            if progreso and (i % 10 == 0 or i == total):
                print(f"    player {i}/{total}...")
            pausas = 0
            while True:
                try:
                    detalles[pid] = trim_old_seasons(client.get_player_detail(pid))
                    break
                except BiwengerAPIError as e:
                    es_rate_limit = "429" in str(e) or "Demasiados" in str(e)
                    if es_rate_limit and parar_en_primer_limite:
                        print(f"    rate limit reached at player {pid} -- "
                              f"stopping here ({len(detalles)}/{total} obtained in this batch).")
                        if on_progress:
                            on_progress(detalles)
                        return detalles
                    if es_rate_limit and pausas < max_pausas_largas:
                        pausas += 1
                        print(f"    (sustained rate-limit, long pause #{pausas} of 30s "
                              f"before retrying player {pid}...)")
                        time.sleep(30)
                        continue
                    print(f"    [!] could not fetch player {pid}: {e}")
                    break
            if on_progress and guardar_cada and i % guardar_cada == 0:
                on_progress(detalles)
    except KeyboardInterrupt:
        if on_progress:
            print(f"    interrupted -- saving progress obtained so far "
                  f"({len(detalles)}/{total} players)...")
            on_progress(detalles)
        raise
    if on_progress:
        on_progress(detalles)
    return detalles


# ---------------------------------------------------------------------------
# Local cache (output/cache/*.json) + orchestration for the notebook: avoids
# requesting everything from the API again (with its rate limit) every time
# you rerun the notebook. Use forzar_refresco=True when you really want
# fresh data.
# ---------------------------------------------------------------------------

# cache files that may be missing from a cache generated by an older
# version of this module, without that forcing a full re-download
_CACHE_FILES_OPCIONALES = {"ofertas_raw"}


def _cache_exists():
    return all(
        p.exists() for key, p in _CACHE_FILES.items() if key not in _CACHE_FILES_OPCIONALES
    )


def _save_cache(data: dict):
    CACHE_DIR.mkdir(parents=True, exist_ok=True)
    for key, path in _CACHE_FILES.items():
        path.write_text(json.dumps(data[key], ensure_ascii=False), encoding="utf-8")


def _read_cache():
    data = {}
    for key, path in _CACHE_FILES.items():
        if not path.exists() and key in _CACHE_FILES_OPCIONALES:
            data[key] = []
            continue
        data[key] = json.loads(path.read_text(encoding="utf-8"))
    # Python dicts with int keys are saved as strings in JSON
    data["detalles"] = {int(k): v for k, v in data["detalles"].items()}
    return data


def _all_ids(data: dict):
    todos_ids = {f["player_id"] for f in data["mercado_raw"] if f.get("player_id")}
    for r in data["rosters_raw"].values():
        todos_ids |= {f["player_id"] for f in r["filas"] if f.get("player_id")}
    return todos_ids


def _missing_details(data: dict):
    """ids of players that appear in the market/rosters but whose detail we
    haven't downloaded yet in `data['detalles']`."""
    return _all_ids(data) - set(data["detalles"].keys())


def _cache_age_hours():
    """Hours since the cache was last saved (based on the modification date
    of its files). None if there's no cache."""
    ref = _CACHE_FILES["standings"]
    if not ref.exists():
        return None
    return (time.time() - ref.stat().st_mtime) / 3600


def load_league_data(email, password, league_id=None, own_team_id=None, forzar_refresco=False,
                       max_antiguedad_horas=24, parar_en_primer_limite=False):
    """
    Main entry point for the analysis notebook: downloads (or reuses from
    output/cache/*.json) the market, all rivals' rosters, your roster, and
    the detail for every player involved.

    By default it reuses the local cache if it already exists AND is
    complete AND is no more than `max_antiguedad_horas` old (24 by default)
    -- past that it refreshes itself automatically, so you don't have to
    remember to pass forzar_refresco=True every day. Pass
    max_antiguedad_horas=None to disable this check (use the cache even if
    it's old, as long as it's complete). forzar_refresco=True always forces
    a refresh, regardless of age.

    RESUMABLE: each player's detail (the slowest part, due to the API rate
    limit) is saved to cache incrementally as it's obtained -- if this
    function is interrupted (Ctrl+C, a cutoff, etc.) partway through, the
    next time you call it (without forzar_refresco) it picks up where it
    left off: it does NOT re-request players it already had, only the ones
    still missing. This is different from the automatic age-based refresh:
    an INCOMPLETE cache is always resumed (no matter how much time has
    passed), because interrupted doesn't mean outdated.

    parar_en_primer_limite: if True, as soon as the API rate limit is hit
    (sustained 429) it exits the function right there instead of waiting
    with long pauses -- what was obtained up to that point is already saved
    to cache (see above). Useful for gradually completing the cache across
    several short runs: you call it again later and it resumes only what's
    missing.

    In any case -- complete, incomplete, or stopped due to the rate limit
    -- what this function returns is always re-read from the cache on disk
    (not from whatever is in memory for this call), because the cache is
    the combination of this run with all previous ones.
    """
    detalles_previos = {}
    if not forzar_refresco and _cache_exists():
        cache = _read_cache()
        faltan_cache = _missing_details(cache)
        if not faltan_cache:
            antiguedad = _cache_age_hours()
            demasiado_vieja = (
                max_antiguedad_horas is not None
                and antiguedad is not None
                and antiguedad >= max_antiguedad_horas
            )
            if not demasiado_vieja:
                aviso_antiguedad = f" ({antiguedad:.1f}h old)" if antiguedad is not None else ""
                print(f"Loading data from local cache ({CACHE_DIR}), complete{aviso_antiguedad}. "
                      "Call load_league_data(..., forzar_refresco=True) to request fresh data.")
                return cache
            print(f"Local cache is complete but is {antiguedad:.1f}h old "
                  f"(more than {max_antiguedad_horas}h) -- refreshing automatically "
                  "with fresh player data...")
            # NOTE: we deliberately don't reuse detalles_previos here -- an
            # old but complete cache needs truly updated points/price/form,
            # not just repeating what it already had
        else:
            detalles_previos = cache["detalles"]
            print(f"Local cache is incomplete: already have {len(detalles_previos)} players, "
                  f"{len(faltan_cache)} missing. Resuming -- won't re-request what's already downloaded...")

    print("Downloading data from the Biwenger API...")
    client = BiwengerClient(email, password, league_id=league_id)
    login_and_resolve_league(client, league_id)
    standings, score_id = load_league(client)
    mi_team_id = identify_my_team(standings, client, own_team_id)

    print("  market...")
    market_full = client.get_market()
    mercado_raw = _parse_market(market_full)
    ofertas_raw = _parse_offers(market_full)
    balance = _dig(market_full, ["status", "balance"])

    print(f"  rosters for {len(standings)} team(s)...")
    rosters_raw = {}
    for s in standings:
        team_id = s.get("id")
        if not team_id:
            continue
        nombre = first_present(s, ["name"], f"equipo:{team_id}")
        es_mio = str(team_id) == str(mi_team_id)
        filas = load_squad(client, team_id, nombre)
        # a standing's team_id is the same value as its user_id (the one
        # needed as 'to' in place_offer to buy out a player's clause)
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

    todos_ids = _all_ids(data_base)
    faltan_ids = todos_ids - set(detalles_previos.keys())
    if detalles_previos:
        print(f"  {len(detalles_previos)} of {len(todos_ids)} players already cached; "
              f"requesting the {len(faltan_ids)} still missing...")
    else:
        print(f"  detail for {len(todos_ids)} players...")

    def save_progress(detalles_parciales):
        data_actual = dict(data_base)
        data_actual["detalles"] = {**detalles_previos, **detalles_parciales}
        _save_cache(data_actual)

    nuevos_detalles = enrich_with_details(
        client, faltan_ids, on_progress=save_progress, parar_en_primer_limite=parar_en_primer_limite,
    )

    data = dict(data_base)
    data["detalles"] = {**detalles_previos, **nuevos_detalles}
    _save_cache(data)
    faltan_todavia = len(_missing_details(data))
    if faltan_todavia:
        print(f"Data saved to local cache ({CACHE_DIR}). "
              f"Note: {faltan_todavia} players are still missing -- "
              "call this function again to retry just those.")
    else:
        print(f"Complete data saved to local cache ({CACHE_DIR}).")
    # always re-read from disk: the cache is the combination of this call
    # with all previous ones, it's the source of truth
    return _read_cache()


def build_dataframe(data: dict) -> pd.DataFrame:
    """
    Builds the combined DataFrame (market + rivals + your team, each row
    with its 'Source') from what load_league_data() returns.

    Determines the "current" league season by majority vote across all
    players (see most_common_league_season) so that a player who hasn't
    played in a while doesn't carry over points from 2+ seasons ago
    disguised as "this season"/"the previous one".
    """
    detalles = data["detalles"]
    score_id = data["score_id"]
    temporada_actual_id = most_common_league_season(detalles)
    temporada_anterior_id = (
        str(int(temporada_actual_id) - 1)
        if temporada_actual_id is not None and temporada_actual_id.isdigit()
        else None
    )
    filas = []

    for f in data["mercado_raw"]:
        # if the player has an owner (vendedor_id present), the ONLY real
        # way to sign them is a buyout clause -- the "market price" the API
        # gives for them isn't an alternative direct purchase, it's purely
        # informative (sometimes it's LOWER than the clause, sometimes MUCH
        # higher: it doesn't represent a real purchase price). That player
        # already appears via their 'Rival: X' Origen row with the correct
        # clause -- don't duplicate them here as if they were purchasable
        # on their own.
        if f.get("vendedor_id") is not None:
            continue
        detalle = detalles.get(f["player_id"])
        if not detalle:
            continue
        if is_irrelevant_player(detalle, score_id, temporada_actual_id, temporada_anterior_id):
            continue
        filas.append(build_player_row(
            f["player_id"], detalle, score_id,
            {
                "Source": "Market",
                "Market price": f.get("precio_mercado"),
                "Seller (id)": f.get("vendedor_id"),
            },
            temporada_actual_id, temporada_anterior_id,
        ))

    for r in data["rosters_raw"].values():
        origen = "My team" if r["es_mio"] else f"Rival: {r['nombre']}"
        for f in r["filas"]:
            detalle = detalles.get(f["player_id"])
            if not detalle:
                continue
            # the "irrelevant player" filter is NOT applied to your own
            # team -- you always see your own players, whether active or not
            if not r["es_mio"] and is_irrelevant_player(
                detalle, score_id, temporada_actual_id, temporada_anterior_id
            ):
                continue
            pts_actual, partidos_actual, pts_anterior, partidos_anterior, _, _ = (
                extract_seasons_points(detalle, score_id, temporada_actual_id, temporada_anterior_id)
            )
            clausula = f.get("clausula")
            bloqueada_hasta = f.get("clausula_bloqueada_hasta")
            extra = {
                "Source": origen,
                "Clause": clausula,
                # two ways to look at the ratio: by TOTAL points accumulated
                # this season, or by AVERAGE points (per match) -- a player
                # with few matches played can look bad on the total but
                # good on the average, and vice versa
                "Ratio total pts/clause (current)": ratio(pts_actual, clausula),
                "Ratio avg pts/clause (current)": ratio(points_per_match(pts_actual, partidos_actual), clausula),
                "Ratio total pts/clause (previous)": ratio(pts_anterior, clausula),
                "Ratio avg pts/clause (previous)": ratio(
                    points_per_match(pts_anterior, partidos_anterior), clausula
                ),
                "Seller (id)": None if r["es_mio"] else r.get("team_id"),
                "Clause locked until": (
                    datetime.fromtimestamp(bloqueada_hasta) if bloqueada_hasta else None
                ),
                "Clause available now": (
                    not bloqueada_hasta or bloqueada_hasta <= datetime.now().timestamp()
                ),
            }
            filas.append(build_player_row(
                f["player_id"], detalle, score_id, extra, temporada_actual_id, temporada_anterior_id
            ))

    return pd.DataFrame(filas)


def offers(data: dict, tipo="recibidas", solo_identificadas=False) -> pd.DataFrame:
    """
    Active buy offers on your market, with the player's name already
    resolved. tipo:
    - 'recibidas': offers other managers have made for YOUR players
      (useful for deciding whether to accept/reject from the app; this
      library has no method to accept/reject, only to query them).
    - 'enviadas': offers YOU have made (via place_offer or another means).
    - 'todas': unfiltered.

    IMPORTANT: Biwenger automatically generates a "received" offer for
    EVERY player on your roster, close to their market value, with NO real
    bidder behind it (bidder 'Unknown', no 'de_id') -- it's the
    platform's quick sale, not real demand from another manager. Detected
    with real data: 15/15 players on a roster had one of these. Pass
    solo_identificadas=True to discard them and keep only offers from an
    identified manager (real demand).
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
            "Player": nombre,
            "Amount": o.get("importe"),
            "Type": o.get("tipo"),
            "Status": o.get("estado"),
            "From": o.get("de_nombre") or "Unknown",
            "To": o.get("a_nombre") or "Unknown",
            "Expires": datetime.fromtimestamp(expira) if expira else None,
        })
    return pd.DataFrame(filas)


# ---------------------------------------------------------------------------
# History (output/history/<timestamp>/*.json): dated copies of the data, so
# you can compare how your team/the market evolves over time (the
# output/cache/ cache, by contrast, is overwritten on every refresh).
# ---------------------------------------------------------------------------

def save_snapshot(data: dict):
    """Saves a dated copy of `data` (what load_league_data returns) to
    output/history/<timestamp>/. Call it right after a forzar_refresco=True
    if you want to be able to compare later on."""
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    carpeta = HISTORY_DIR / ts
    carpeta.mkdir(parents=True, exist_ok=True)
    for key in _CACHE_FILES:
        (carpeta / f"{key}.json").write_text(json.dumps(data[key], ensure_ascii=False), encoding="utf-8")
    print(f"Snapshot saved to {carpeta}")
    return carpeta


def list_snapshots():
    """Names (timestamps) of saved snapshots, oldest first."""
    if not HISTORY_DIR.exists():
        return []
    return sorted(p.name for p in HISTORY_DIR.iterdir() if p.is_dir())


def load_snapshot(nombre: str) -> dict:
    """Loads a snapshot saved by save_snapshot() (use list_snapshots()
    to see the available names). Returns the same format as
    load_league_data(), ready to pass to build_dataframe()."""
    carpeta = HISTORY_DIR / nombre
    data = {}
    for key in _CACHE_FILES:
        data[key] = json.loads((carpeta / f"{key}.json").read_text(encoding="utf-8"))
    data["detalles"] = {int(k): v for k, v in data["detalles"].items()}
    return data
