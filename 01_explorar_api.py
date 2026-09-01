"""
Script de diagnostico (ejecutar UNA VEZ antes que 02_generar_excel.py).

Inicia sesion contra la API real de Biwenger y vuelca respuestas reales
a output/debug/*.json. El objetivo es confirmar juntos los nombres de
campo exactos -sobre todo si el detalle de jugador trae de verdad el
historico de la temporada anterior en 'seasons'- antes de dar por bueno
el calculo de ratios.

Uso:
    1. Copia .env.example a .env y rellena tu email/password de Biwenger.
    2. python 01_explorar_api.py
    3. Revisa los .json generados en output/debug/ (o mandamelos/pegamelos)
       y seguimos ajustando 02_generar_excel.py si algun nombre de campo
       no es el esperado.
"""

import json
import os
import sys
from pathlib import Path

from dotenv import load_dotenv

from biwenger_client import BiwengerClient

# evita que un simbolo no-ASCII en un print tumbe el script en consolas que
# no usan UTF-8 (p.ej. cmd.exe con la codepage por defecto) -- lo sustituye
# en vez de lanzar UnicodeEncodeError
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(errors="replace")

load_dotenv()

EMAIL = os.getenv("BIWENGER_EMAIL")
PASSWORD = os.getenv("BIWENGER_PASSWORD")
LEAGUE_ID = os.getenv("BIWENGER_LEAGUE_ID") or None

OUT_DIR = Path(__file__).parent / "output" / "debug"


def dump(name: str, data) -> None:
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    path = OUT_DIR / f"{name}.json"
    path.write_text(json.dumps(data, indent=2, ensure_ascii=False), encoding="utf-8")
    print(f"  -> guardado {path}")


def main() -> None:
    if not EMAIL or not PASSWORD:
        raise SystemExit(
            "Faltan BIWENGER_EMAIL / BIWENGER_PASSWORD.\n"
            "Copia .env.example a .env y rellenalo con tus datos de Biwenger."
        )

    client = BiwengerClient(EMAIL, PASSWORD, league_id=LEAGUE_ID)

    print("1) Iniciando sesion...")
    login_payload = client.login()
    dump("login_response", login_payload)
    print(f"   user_id detectado: {client.user_id}")

    print("2) Buscando ligas de tu cuenta...")
    try:
        leagues, match = client.resolve_league(LEAGUE_ID)
        dump("account_leagues", leagues)
        print(f"   {len(leagues)} liga(s) encontradas")
        if not client.league_id:
            print(
                "   No se ha encontrado ninguna liga automaticamente.\n"
                "   Anade BIWENGER_LEAGUE_ID a tu .env con el ID de tu liga "
                "(esta en la URL cuando entras a tu liga en biwenger.com) "
                "y vuelve a ejecutar este script."
            )
            return
        if match is None and LEAGUE_ID:
            print(
                f"   [!] No se ha podido confirmar la liga {LEAGUE_ID} contra tu "
                "cuenta; el x-user puede no ser correcto para ella."
            )
        else:
            print(f"   usando la liga: {client.league_id} (user_id liga: {client.user_id})")
    except Exception as e:
        print(f"   (fallo al listar ligas automaticamente: {e})")
        print(
            "   Anade BIWENGER_LEAGUE_ID a tu .env con el ID de tu liga y "
            "vuelve a ejecutar este script."
        )
        return

    print("3) Consultando la liga (clasificacion / standings)...")
    league = client.get_league()
    dump("league", league)

    print("4) Consultando el mercado...")
    market = client.get_market()
    dump("market", market)

    standings = league.get("standings") if isinstance(league, dict) else None
    if not standings:
        print(
            "   [!] No se ha encontrado 'standings' en la respuesta de la liga. "
            "Revisa output/debug/league.json a mano para localizar donde "
            "estan los equipos/rivales."
        )
        return

    sample = standings[0]
    sample_team_id = sample.get("id") or _dig(sample, ["team", "id"])
    if not sample_team_id:
        print("   [!] No se ha podido sacar un team_id de muestra de standings[0].")
        return

    print(f"5) Consultando un equipo de muestra (id={sample_team_id})...")
    team = client.get_team(sample_team_id)
    dump("sample_team", team)

    players_field = team.get("players") if isinstance(team, dict) else None
    player_ids = []
    if isinstance(players_field, list):
        player_ids = [p.get("id") for p in players_field if isinstance(p, dict) and p.get("id")]

    if not player_ids:
        print("   [!] No se han encontrado jugadores en el equipo de muestra.")
        return

    print(f"6) Consultando el detalle de un jugador de muestra (id={player_ids[0]})...")
    player = client.get_player_detail(player_ids[0])
    dump("sample_player", player)

    if isinstance(player, dict) and player.get("seasons"):
        print("   'seasons' SI aparece en el detalle del jugador -> pinta bien")
        print(f"   contenido de 'seasons': {json.dumps(player['seasons'], ensure_ascii=False)[:500]}")
    else:
        print("   [!] 'seasons' no aparece (o viene vacio) en el detalle del jugador.")
        print("   Revisa output/debug/sample_player.json entero por si el historico")
        print("   esta bajo otro nombre de campo.")

    print("\nListo. Revisa los ficheros en output/debug/ y cuentame lo que veas,")
    print("sobre todo en sample_player.json, para calibrar 02_generar_excel.py.")


def _dig(d, path):
    cur = d
    for p in path:
        if not isinstance(cur, dict) or p not in cur:
            return None
        cur = cur[p]
    return cur


if __name__ == "__main__":
    main()
