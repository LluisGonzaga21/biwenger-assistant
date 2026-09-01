# Biwenger Analysis

Python tools to analyze your [Biwenger](https://biwenger.as.com/) (fantasy football) league beyond what the app shows you: performance-per-price ratios, a "form" score weighted by recent matches, fair market value estimation (bargains/overpriced players), buy/sell suggestions that tell you exactly who they'd replace, and even placing bids/clause buyouts straight from code (in a safe, no-surprises mode).

Everything is built on top of the **unofficial Biwenger API** (there's no public documentation), by inspecting real responses. If Biwenger changes something, this is the first place to look — see [How it works under the hood](#how-it-works-under-the-hood).

> ⚠️ This is a personal project, not affiliated with Biwenger/AS.com. Use at your own risk: it's an undocumented API and can break if Biwenger changes its backend. The "operations" part (bidding/clause buyouts) moves real money in your team — read the [safety](#safety-when-biddingbuying-out-a-clause) section before using it.

## Table of contents

- [What does it actually do?](#what-does-it-actually-do)
- [Installation](#installation)
- [Configuration](#configuration)
- [Quick start](#quick-start)
- [The notebook, section by section](#the-notebook-section-by-section)
- [The DataFrame: columns explained](#the-dataframe-columns-explained)
- [Function reference](#function-reference-biwenger_helperspy)
- [Safety when bidding/buying out a clause](#safety-when-biddingbuying-out-a-clause)
- [How it works under the hood](#how-it-works-under-the-hood)
- [Project structure](#project-structure)
- [Troubleshooting](#troubleshooting)
- [Contact / contributing](#contact--contributing)

## What does it actually do?

Biwenger shows you one player at a time. This project downloads **everything** relevant to your league (the market, every rival's squad, the detail of every player involved) and merges it into a single table you can sort, filter and cross-reference however you want. From there:

- **Performance-per-price ratios**: points / market value, points / clause price (current and previous season), both in total and per match.
- **`Puntuacion potencial` ("potential score")**: a custom metric that blends recent form (last matches, weighted — the most recent counts more) with last season's performance, so it doesn't put all its trust in a 2-match hot streak.
- **Fair market value**: for every player on the market, estimates whether they're cheap or expensive by comparing them to similar players in your own league (not an external "official" price).
- **How much to bid**: combines your estimated fair value, how many rivals could outbid you, and what's actually been paid for similar players in your league — never suggests bidding below the asking price (Biwenger doesn't allow that anyway).
- **Signing suggestions**: compares every affordable candidate (market or clause) against your weakest player, with no "this position is already covered" bias — it measures real improvement and tells you exactly who you'd be replacing.
- **Sell suggestions**: who in your squad is currently overvalued, cross-checked against real offers you've already received.
- **Charts**: bargains/overvalued players with outliers labeled, best/worst by position, performance distribution.
- **Resumable cache**: the initial download (market + every squad + every player's detail) can take several minutes because of the API's rate limit. It's saved to disk incrementally, so if it gets interrupted, the next run picks up only what's missing.
- **Real bidding/clause buyouts**: dry-run by default (just shows you what would be sent), with an explicit `confirm=True` to actually execute it.

## Installation

You need **Python 3.10+**.

```bash
git clone <your-fork-url>
cd biwenger-analisis
python -m venv .venv

# Windows
.venv\Scripts\activate
# macOS / Linux
source .venv/bin/activate

pip install -r requirements.txt
```

## Configuration

Copy the example file and fill in your credentials:

```bash
copy .env.example .env      # Windows
cp .env.example .env        # macOS / Linux
```

Open `.env` in a text editor:

```env
BIWENGER_EMAIL=your_email@example.com
BIWENGER_PASSWORD=your_password

# Optional: only needed if your account is in several leagues
BIWENGER_LEAGUE_ID=

# Optional: only needed if auto-detect can't figure out which team is yours
BIWENGER_OWN_TEAM_ID=
```

`.env` is in `.gitignore` — **it's never uploaded anywhere**, it stays only on your machine. Login happens on every run because Biwenger's session token expires.

## Quick start

### Option 1: the interactive notebook (recommended)

```bash
python 01_explorar_api.py   # confirms login and league detection work (run once, or if something breaks)
```

Open `03_analisis.ipynb` in VS Code or Jupyter, select the `.venv` interpreter as the kernel, and run the cells top to bottom. The first data load takes a few minutes (see [resumable cache](#how-it-works-under-the-hood)); after that it's instant.

### Option 2: static Excel export

```bash
python 02_generar_excel.py            # uses the cache if it already exists
python 02_generar_excel.py --refresh  # forces a full re-download
```

Generates `output/analisis_biwenger_<date>.xlsx` with three sheets (Market, Rivals, My team), frozen columns and conditional formatting (red→yellow→green) on the ratios. Quick to share or print, but no charts or recommendations — use the notebook for that. It shares the same cache as the notebook, so if you already loaded data from one of the two, the other reuses it instantly.

## The notebook, section by section

### 1. Load data

```python
data = biwenger_helpers.load_league_data(
    EMAIL, PASSWORD, league_id=LEAGUE_ID, own_team_id=OWN_TEAM_ID,
    forzar_refresco=False, parar_en_primer_limite=True,
)
```

The first run downloads the market + every manager's squad + the detail of every player involved, and saves it to `output/cache/*.json`. It's **resumable**: with `parar_en_primer_limite=True` the load stops itself as soon as the API cuts off due to too many requests (instead of sitting through long backoff pauses), and the next time you run the cell it picks up only what's missing. If the cache is older than 24h (`max_antiguedad_horas`) it refreshes itself automatically.

### 2. Explore

`df` has one row per player — market + every rival + your own team — with the `Origen` ("source") column to filter by. Examples:

```python
# Top 20 market bargains by points/market-value ratio
df[df["Origen"] == "Mercado"].dropna(subset=["Ratio pts/VM (actual)"]) \
    .sort_values("Ratio pts/VM (actual)", ascending=False).head(20)

# Your squad sorted by performance (to spot your weakest players)
df[df["Origen"] == "Mi equipo"].sort_values("Puntuacion potencial")

# Best rival clause buyouts by points/clause ratio
df[df["Origen"].str.startswith("Rival:")] \
    .dropna(subset=["Ratio pts totales/clausula (actual)"]) \
    .sort_values("Ratio pts totales/clausula (actual)", ascending=False).head(20)
```

**Important note about `Origen`**: `'Mercado'` means a *genuinely free* player (no owner, directly biddable). Market entries that DO have an owner (someone listed their own player for sale) do **not** count as `'Mercado'` — the only real way to sign them is a clause buyout, so they show up as `'Rival: <name>'` with their `Clausula` column filled in. On top of that, market/rivals already exclude "clutter" automatically (no top-flight team, value ≤ €400,000, or inactive players); your own squad is never filtered.

### 3. Charts

- **Top 20 by points/market-value ratio** (horizontal bar chart).
- **Ratio distribution by position** (boxplot).
- **Value vs. performance**: scatter plot with a trend line; the 8 players that deviate the most above (bargains) and below (overvalued) the trend are labeled with their name.
- **Best and worst 5 per position**: four panels (Goalkeeper/Defender/Midfielder/Forward), green on top, red on the bottom.

### 4. Signing suggestions

```python
biwenger_helpers.suggest_signings(df, data["balance"], top_n=15)

# Prioritizing efficiency (improvement per million) instead of raw improvement
biwenger_helpers.suggest_signings(df, data["balance"], top_n=15, ordenar_por="eficiencia")
```

Compares every affordable, available candidate against your weakest player **in that same position**, and only shows candidates that would actually improve on that player. It doesn't apply any "this line is already covered" bias — it measures real improvement and explicitly tells you who you'd be replacing (`Jugador que reemplazarías` column). Clauses that are temporarily locked after a recent purchase are automatically excluded.

Key columns: `Saldo tras fichar` (balance after signing), `Jugadores flojos en tu posicion` / `Total en tu posicion` (how many of your players in that position perform below your own median there), `Mejora por millon gastado` (improvement per million spent).

> 💡 This is a rough suggestion based on the `balance` you pass in — it doesn't account for Biwenger's real `maximumBid` (which can be significantly higher than your balance, since Biwenger lets you bid into the red using your squad's value as collateral). If you want to use your real bidding limit, pass it in yourself instead of `data["balance"]`.

### 5. General ranking

```python
biwenger_helpers.best_players(df, top_n=20)                                    # top 20 overall
biwenger_helpers.best_players(df, posicion="Delantero", origen="Mercado")       # filtered
```

A watchlist of who's performing best right now (by `Puntuacion potencial`), regardless of price or whether you can afford them. Excludes injured/suspended players by default.

### 6. Fair market value

```python
valor_justo = biwenger_helpers.estimate_fair_value(df, margen_pct=20)
valor_justo[valor_justo["Valoracion"] == "Chollo"]

# How much to bid for a genuinely free market player
biwenger_helpers.suggest_market_offer(df, data)
```

For every player on the market, `estimate_fair_value` calculates the median performance/price relationship of similar players in your league (same position) and uses it as a reference to label Bargain/Expensive/Fair.

`suggest_market_offer` goes a step further and estimates how much it would actually take to win the player (not just what they're "worth" by your heuristic): it combines the asking price (**mandatory floor** — Biwenger won't let you bid below it), what's genuinely been paid for comparable players in your league, and how many rivals have budget to spare to outbid you. It's a rough guide — it has no way of knowing who's *actually* interested in each player, only who could afford them; a rival can end up paying well above any reasonable estimate.

### 7. Sell suggestions

```python
ventas = biwenger_helpers.suggest_sales(df, data=data, margen_pct=20)
ventas[ventas["Recomendacion"] == "Vender ahora"]
```

The counterpart to section 6 for your own squad: players whose market value has shot up well beyond what their actual performance would justify. It also cross-checks real offers you've already received (filtering out Biwenger's automatic "Unknown" instant-sell offers, which aren't real demand).

### 8. Offers

```python
biwenger_helpers.offers(data, tipo="recibidas")
biwenger_helpers.offers(data, tipo="enviadas")
```

Lists active buy offers in the market without having to open the app (doesn't include accepting/rejecting — that's still done in Biwenger).

### 9. History

```python
biwenger_helpers.save_snapshot(data)          # saves a dated snapshot
biwenger_helpers.list_snapshots()
biwenger_helpers.load_snapshot("20260828_120000")
```

Unlike the cache (which gets overwritten), `save_snapshot` saves a dated copy under `output/history/`, useful for comparing how your team or the market evolves over time.

### 10. Operations (bidding / clause buyouts)

Real actions on your league — read the [safety](#safety-when-biddingbuying-out-a-clause) section before touching this.

```python
cliente_operaciones = BiwengerClient(EMAIL, PASSWORD, league_id=LEAGUE_ID)
biwenger_helpers.login_and_resolve_league(cliente_operaciones, LEAGUE_ID)

# 'Vendedor (id)' from the tables above is what goes into 'to'
cliente_operaciones.place_offer(player_id, importe, to=vendedor_id)                    # dry-run: only prints the payload
cliente_operaciones.place_offer(player_id, importe, to=vendedor_id, confirm=True)      # actually executes it
```

## The DataFrame: columns explained

Column names are kept in Spanish (matching the live Biwenger data and the rest of the code); this table explains what each one means.

| Column | What it is |
|---|---|
| `player_id` | Biwenger's internal ID for the player |
| `Jugador`, `Posicion`, `Equipo LaLiga` | Name, position, LaLiga club |
| `Estado`, `Disponible` | `Disponible=False` if injured/suspended/discarded (`'doubt'` does not count as unavailable) |
| `Valor de mercado` | Current official market value, in euros |
| `Tendencia precio (7d) %` | Market value change over the last 7 days |
| `Temporada actual` / `anterior` | Real season label (e.g. "Temporada 2025/2026"); the column name is always the same, computed by majority vote across all players so it doesn't misalign for someone who hasn't played recently |
| `Puntos temporada actual` / `anterior` | Total accumulated points |
| `Partidos temporada actual` / `anterior` | Matches played |
| `Pts/partido (actual)` / `(anterior)` | Total points ÷ matches |
| `Ratio pts/VM (actual)` / `(anterior)` | Total points ÷ market value (in millions) |
| `Forma (ult. partidos)` | Weighted average of the last few matches, weighting more recent ones higher (exponential decay) |
| `Puntuacion potencial` | `Forma` (70%) + previous season's points per match (30%) — the metric used by the rankings and recommenders |
| `Proximo rival`, `Local/Visitante`, `Dificultad proximo partido`, `Jornada proxima` | Their next LaLiga match |
| `Origen` | `'Mercado'` (free agent), `'Mi equipo'` (your team), or `'Rival: <name>'` |
| `Precio en mercado`, `Vendedor (id)` | Only when `Origen == 'Mercado'` |
| `Clausula`, `Clausula bloqueada hasta`, `Clausula disponible ahora` | Only for rival/your own players |
| `Ratio pts totales/clausula` / `Ratio pts medios/clausula` (`actual`/`anterior`) | Same ratio as above but against the clause price instead of market value — by total points or by points-per-match (a player with few matches played can look bad in the total but fine in the average) |

## Function reference (`biwenger_helpers.py`)

Function names are in English; their parameters stay in Spanish, matching the DataFrame columns they filter/compare against.

### Loading and cache

| Function | What it does |
|---|---|
| `load_league_data(email, password, league_id=None, own_team_id=None, forzar_refresco=False, max_antiguedad_horas=24, parar_en_primer_limite=False)` | Orchestrates everything: login, market, squads, player details, resumable cache. Returns the `data` dict used by every other function |
| `build_dataframe(data)` | Turns `data` into the combined DataFrame used throughout the notebook |
| `save_snapshot(data)` / `list_snapshots()` / `load_snapshot(nombre)` | Dated history under `output/history/` |

### Analysis and recommendations

| Function | What it does |
|---|---|
| `best_players(df, posicion=None, origen=None, top_n=20, metrica="Puntuacion potencial", solo_disponibles=True)` | Filterable ranking/watchlist |
| `estimate_fair_value(df, metrica="Puntuacion potencial", margen_pct=20, solo_disponibles=True)` | Bargain/Expensive/Fair labeling for the market |
| `suggest_market_offer(df, data, margen_sobre_pedido_pct=5, metrica="Puntuacion potencial", rango_comparables_pct=40)` | How much to bid for a genuinely free market player |
| `suggest_signings(df, balance, top_n=15, solo_disponibles=True, ordenar_por="mejora")` | Candidates that improve on your weakest player in their position (`ordenar_por="eficiencia"` to prioritize cost/benefit instead) |
| `suggest_sales(df, data=None, metrica="Puntuacion potencial", margen_pct=20)` | Your overvalued players, cross-checked against real offers |
| `cost_per_point(df, metrica="Puntuacion potencial", margen_pct=20)` | Cost per point for every player against the median of the ENTIRE game (all positions combined), to see if they're cheap/expensive in absolute terms |
| `offers(data, tipo="recibidas", solo_identificadas=False)` | Active buy offers (`solo_identificadas=True` drops Biwenger's automatic offers, which aren't real demand) |

### Internal building blocks (in case you tweak something)

| Function | What it does |
|---|---|
| `potential_score(detalle, score_id, decay=0.85, n_partidos=6, peso_temporada_anterior=0.3, ...)` | The base metric — adjust the weights here if you want recent form to count for more or less |
| `calculate_form(detalle, score_id, decay=0.85, n_partidos=6)` | Weighted average of the last N matches |
| `is_irrelevant_player(detalle, score_id, ...)` | The "clutter" filter (no top-flight team, low value, inactive) |
| `most_common_league_season(detalles)` | Detects the real "current" season by majority vote across all players, so it doesn't drag in stale seasons for someone who hasn't debuted yet |

## Safety when bidding/buying out a clause

`client.place_offer(...)` runs in **dry-run mode by default**: without `confirm=True` nothing is sent to Biwenger, it just prints the payload that would be sent (player, amount, to whom). Always double-check that payload before confirming — passing `confirm=True` actually executes the offer, spends or commits your balance, and **cannot be undone** from here.

The `to` field is the `user_id` of the player's current owner (matches the `team_id` you see in the standings). For a market player use `client.find_market_seller(player_id)` or the `Vendedor (id)` column from the notebook's tables; for a clause buyout on a rival, use their `team_id` directly.

The offer type (`offer_type="purchase"` by default) is confirmed for market bids by several unofficial Biwenger clients, but is **not** confirmed for clause buyouts on a rival — before confirming a clause buyout for the first time, compare the dry-run payload against the real request you see in your browser's DevTools when doing it from biwenger.com.

## How it works under the hood

- **Unofficial API**: uses `biwenger.as.com/api/v2` (authenticated, for league/market/squads/bids) and `cf.biwenger.com/api/v2` (public, for each player's detail). There's no official documentation; field names come from inspecting real responses — if something stops working, run `01_explorar_api.py` and compare `output/debug/*.json` against what `biwenger_helpers.py` expects.
- **`x-user` is the team ID, not the account ID**: the API requires the `x-user` header to be your team ID *within that specific league* (`leagues[].user.id`), not the global account ID from the login token. `resolve_league()` resolves this automatically.
- **`cf.biwenger.com` rejects session headers**: player-detail requests explicitly null out `Authorization`/`x-league`/`x-user` — otherwise it responds with a 403.
- **Rate limit**: `cf.biwenger.com` cuts off around request ~200-211 of a single run. It's not a short window that resets if you wait it out (we tried a preemptive 5-minute pause and the cutoff still hit at the exact same point) — it's a cap on total requests. The only real way to avoid it entirely is to not repeat requests, hence the shared, resumable cache: each player's detail is saved as soon as it's fetched, so an interrupted load only picks up what's missing next time.
- **`maximumBid` can exceed your balance**: Biwenger lets you bid into the red using your squad's value as collateral. `suggest_signings` uses whatever `balance` you pass it as-is — if you want your real bidding limit, pass it `maximumBid` (from `standings`) instead of `data["balance"]`.
- **Points by scoring system**: every match report has points broken down by `scoreID` (different leagues can score differently). All analysis automatically uses your league's own `scoreID`.

## Project structure

```
biwenger_client.py      API client: login, reading data, and writing (bids/clause buyouts)
biwenger_helpers.py     All the analysis logic and data loading/caching
01_explorar_api.py      Diagnostics: confirms login/league works and dumps real responses to output/debug/
02_generar_excel.py     Generates a static Excel export (market + rivals + your team)
03_analisis.ipynb       Interactive notebook — the recommended day-to-day way to use this
.env.example            Credentials template (copy to .env, never uploaded)
output/                 Generated locally (cache, debug dumps, Excel, snapshots) — in .gitignore
```

## Troubleshooting

- **`401 Invalid email or password`**: check `.env`. If your credentials are correct and it still fails, it might be a temporary lockout from too many logins in a row — wait a few minutes.
- **`400 X-League and X-User headers required` / `401 Invalid user`**: usually resolved automatically by `resolve_league()` / `login_and_resolve_league()`. If you're in multiple leagues, set `BIWENGER_LEAGUE_ID` in `.env`.
- **"Field not found" warnings**: run `01_explorar_api.py`, look at the real JSON under `output/debug/`, and adjust `CAMPOS_CANDIDATOS` in `biwenger_helpers.py`.
- **Data loading cuts off / is slow**: that's the rate limit (see [above](#how-it-works-under-the-hood)) — nothing special to do, just re-run the cell/script later and it'll resume from where it left off.
- **`ModuleNotFoundError`**: make sure the active interpreter is the one in `.venv` (`.venv\Scripts\python.exe` on Windows), not a global Python install.
- **The notebook isn't picking up changes to `biwenger_helpers.py`**: restart the kernel — Python caches modules that are already imported.

## Contact / contributing

This started as a personal project, but issues, pull requests and suggestions are welcome. If you run into a bug, have an idea, or just want to ask something about how it works, reach out: **lluisgonzaga21@gmail.com**.
