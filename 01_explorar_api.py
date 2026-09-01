"""
Diagnostic script (run it ONCE before 02_generar_excel.py).

Logs into the real Biwenger API and dumps real responses to
output/debug/*.json. The goal is to confirm together the exact field
names -especially whether the player detail actually brings the
previous season's history in 'seasons'- before trusting the ratio
calculations.

Usage:
    1. Copy .env.example to .env and fill in your Biwenger email/password.
    2. python 01_explorar_api.py
    3. Check the generated .json files in output/debug/ (or send/paste them
       to me) and we'll keep adjusting 02_generar_excel.py if any field
       name isn't the one expected.
"""

import json
import os
import sys
from pathlib import Path

from dotenv import load_dotenv

from biwenger_client import BiwengerClient

# prevents a non-ASCII character in a print from crashing the script on
# consoles that don't use UTF-8 (e.g. cmd.exe with the default codepage) --
# it replaces the character instead of raising UnicodeEncodeError
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
    print(f"  -> saved {path}")


def main() -> None:
    if not EMAIL or not PASSWORD:
        raise SystemExit(
            "Missing BIWENGER_EMAIL / BIWENGER_PASSWORD.\n"
            "Copy .env.example to .env and fill it in with your Biwenger credentials."
        )

    client = BiwengerClient(EMAIL, PASSWORD, league_id=LEAGUE_ID)

    print("1) Logging in...")
    login_payload = client.login()
    dump("login_response", login_payload)
    print(f"   detected user_id: {client.user_id}")

    print("2) Looking up your account's leagues...")
    try:
        leagues, match = client.resolve_league(LEAGUE_ID)
        dump("account_leagues", leagues)
        print(f"   {len(leagues)} league(s) found")
        if not client.league_id:
            print(
                "   No league could be found automatically.\n"
                "   Add BIWENGER_LEAGUE_ID to your .env with your league's ID "
                "(it's in the URL when you open your league on biwenger.com) "
                "and run this script again."
            )
            return
        if match is None and LEAGUE_ID:
            print(
                f"   [!] Could not confirm league {LEAGUE_ID} against your "
                "account; the x-user might not be correct for it."
            )
        else:
            print(f"   using league: {client.league_id} (league user_id: {client.user_id})")
    except Exception as e:
        print(f"   (failed to list leagues automatically: {e})")
        print(
            "   Add BIWENGER_LEAGUE_ID to your .env with your league's ID and "
            "run this script again."
        )
        return

    print("3) Fetching the league (standings)...")
    league = client.get_league()
    dump("league", league)

    print("4) Querying the market...")
    market = client.get_market()
    dump("market", market)

    standings = league.get("standings") if isinstance(league, dict) else None
    if not standings:
        print(
            "   [!] 'standings' was not found in the league response. "
            "Check output/debug/league.json by hand to locate where "
            "the teams/opponents are."
        )
        return

    sample = standings[0]
    sample_team_id = sample.get("id") or _dig(sample, ["team", "id"])
    if not sample_team_id:
        print("   [!] Could not get a sample team_id from standings[0].")
        return

    print(f"5) Fetching a sample team (id={sample_team_id})...")
    team = client.get_team(sample_team_id)
    dump("sample_team", team)

    players_field = team.get("players") if isinstance(team, dict) else None
    player_ids = []
    if isinstance(players_field, list):
        player_ids = [p.get("id") for p in players_field if isinstance(p, dict) and p.get("id")]

    if not player_ids:
        print("   [!] No players found in the sample team.")
        return

    print(f"6) Fetching the detail of a sample player (id={player_ids[0]})...")
    player = client.get_player_detail(player_ids[0])
    dump("sample_player", player)

    if isinstance(player, dict) and player.get("seasons"):
        print("   'seasons' DOES appear in the player detail -> looks good")
        print(f"   'seasons' content: {json.dumps(player['seasons'], ensure_ascii=False)[:500]}")
    else:
        print("   [!] 'seasons' is missing (or empty) in the player detail.")
        print("   Check the whole output/debug/sample_player.json in case the history")
        print("   is under a different field name.")

    print("\nDone. Check the files in output/debug/ and tell me what you see,")
    print("especially in sample_player.json, so we can calibrate 02_generar_excel.py.")


def _dig(d, path):
    cur = d
    for p in path:
        if not isinstance(cur, dict) or p not in cur:
            return None
        cur = cur[p]
    return cur


if __name__ == "__main__":
    main()
