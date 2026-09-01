"""
Unofficial client for the Biwenger API.

Biwenger does not publish official documentation for its API, so this
client is built from the endpoint structure used by several community
projects (biwenger.as.com / cf.biwenger.com). That's why it is
deliberately defensive: if some expected field is missing from a
response, it doesn't crash the whole program, it warns and continues
with what it can.

Before fully trusting the calculations in 02_generar_excel.py, run
01_explorar_api.py once: it dumps real responses to output/debug/ so
you can confirm the exact field names (especially the historical data
from previous seasons in 'seasons').
"""

from __future__ import annotations

import base64
import json
import time
from typing import Any, Optional

import requests

BASE_URL = "https://biwenger.as.com/api/v2"
CF_URL = "https://cf.biwenger.com/api/v2"

DEFAULT_HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/124.0 Safari/537.36"
    ),
    "Content-Type": "application/json",
}


class BiwengerAPIError(RuntimeError):
    pass


def _jwt_claim(token: str, claim: str):
    """Decode (without verifying the signature) the payload of a JWT and return a claim.

    Biwenger's login response sometimes only includes the 'token' (without a
    'user' object), but the user id lives inside the JWT itself as the 'iss'
    claim.
    """
    try:
        parts = token.split(".")
        if len(parts) != 3:
            return None
        payload_b64 = parts[1]
        padding = "=" * (-len(payload_b64) % 4)
        payload = json.loads(base64.urlsafe_b64decode(payload_b64 + padding))
        return payload.get(claim)
    except Exception:
        return None


def _dig(d: Any, path: list):
    """Navigate a nested dict, returning None if any step doesn't exist."""
    cur = d
    for p in path:
        if not isinstance(cur, dict) or p not in cur:
            return None
        cur = cur[p]
    return cur


class BiwengerClient:
    def __init__(
        self,
        email: str,
        password: str,
        league_id: Optional[str] = None,
        lang: str = "es",
        request_delay: float = 0.3,
    ):
        self.email = email
        self.password = password
        self.lang = lang
        self.league_id = league_id
        self.request_delay = request_delay

        self.token: Optional[str] = None
        self.user_id: Optional[str] = None

        self.session = requests.Session()
        self.session.headers.update(DEFAULT_HEADERS)

    # ---------------------------------------------------------------- auth

    def login(self) -> dict:
        resp = self.session.post(
            f"{BASE_URL}/auth/login",
            json={"email": self.email, "password": self.password},
        )
        self._raise_for_status(resp, "login")
        payload = resp.json()
        data = payload.get("data", payload)

        token = _dig(data, ["token"])
        if isinstance(token, dict):
            token = token.get("token")
        if not token:
            raise BiwengerAPIError(
                "Could not extract the token from the login response. "
                f"Response received: {json.dumps(payload, ensure_ascii=False)[:800]}"
            )
        self.token = token

        user_id = _dig(data, ["user", "id"]) or _dig(data, ["id"])
        if not user_id:
            user_id = _jwt_claim(token, "iss")
        self.user_id = user_id

        self._refresh_headers()
        return payload

    def _refresh_headers(self):
        headers = {"x-lang": self.lang, "Authorization": f"Bearer {self.token}"}
        if self.league_id:
            headers["x-league"] = str(self.league_id)
        if self.user_id:
            headers["x-user"] = str(self.user_id)
        self.session.headers.update(headers)

    def set_league(self, league_id, user_id=None):
        self.league_id = league_id
        if user_id:
            self.user_id = user_id
        self._refresh_headers()

    def resolve_league(self, league_id=None):
        """Detect the league (and the team-id that must be sent in x-user for
        that specific league) from /account.

        The API requires x-user to be the TEAM id within the league
        (leagues[].user.id), not the global account id from the login JWT.
        Without this, /league returns 401 "Invalid user".
        """
        leagues = self.discover_leagues()
        match = None
        if league_id:
            for l in leagues:
                if str(l.get("id")) == str(league_id):
                    match = l
                    break
            if match is None:
                # Could not confirm this against /account (a league from a
                # friend's account?): proceed with the given league_id as-is.
                self.set_league(league_id)
                return leagues, None
        elif leagues:
            match = leagues[0]

        if match is not None:
            self.set_league(match.get("id"), _dig(match, ["user", "id"]))
        return leagues, match

    # ------------------------------------------------------------- generic

    def _get(self, url, params=None, retries=4, headers=None):
        last_error = None
        for attempt in range(retries):
            resp = self.session.get(url, params=params, headers=headers)
            if resp.status_code == 429:
                wait = 5 * (attempt + 1)
                print(f"    (rate limit hit, waiting {wait}s...)")
                time.sleep(wait)
                continue
            try:
                self._raise_for_status(resp, url)
            except BiwengerAPIError as e:
                last_error = e
                break
            payload = resp.json()
            if self.request_delay:
                time.sleep(self.request_delay)
            return payload.get("data", payload)
        raise last_error or BiwengerAPIError(f"Too many 429s requesting {url}")

    def _post(self, url, json_body=None, method="POST", headers=None):
        resp = self.session.request(method, url, json=json_body, headers=headers)
        self._raise_for_status(resp, url)
        payload = resp.json()
        if self.request_delay:
            time.sleep(self.request_delay)
        return payload.get("data", payload)

    @staticmethod
    def _raise_for_status(resp, context):
        if resp.status_code >= 400:
            raise BiwengerAPIError(
                f"Error {resp.status_code} in {context}: {resp.text[:500]}"
            )

    # ------------------------------------------------------- account/league

    def get_account(self) -> dict:
        return self._get(f"{BASE_URL}/account", params={"fields": "*,leagues"})

    def discover_leagues(self) -> list:
        account = self.get_account()
        leagues = account.get("leagues") if isinstance(account, dict) else None
        return leagues or []

    def get_league(self) -> dict:
        return self._get(
            f"{BASE_URL}/league",
            params={"include": "all", "fields": "*,standings,group,settings(description)"},
        )

    # ------------------------------------------------------ market/teams

    def get_market(self) -> dict:
        return self._get(f"{BASE_URL}/market")

    def find_market_seller(self, player_id, market=None):
        """Search the market for who is selling `player_id` and return their
        user_id (the one to pass as `to` in place_offer to bid for them).

        If you already have the result of get_market() at hand, pass it in
        `market` to avoid fetching it again.
        """
        market = market if market is not None else self.get_market()
        sales = market.get("sales") if isinstance(market, dict) else None
        if not isinstance(sales, list):
            return None
        for sale in sales:
            if _dig(sale, ["player", "id"]) == int(player_id):
                return _dig(sale, ["user", "id"])
        return None

    def place_offer(self, player_id, amount, to=None, offer_type="purchase", confirm=False) -> dict:
        """Send an offer/bid for a player: buying a player that is on the
        market, or a clause buyout for a player belonging to another manager
        in your league who is NOT on the market.

        - player_id: the player's id.
        - amount: the offer amount (or the clause price, for a buyout).
        - to: user_id of the player's current owner. For a player on the
          market you can get it with find_market_seller(player_id); for a
          buyout it's the user.id you see in standings/your roster.
        - offer_type: "purchase" is the type confirmed by several unofficial
          clients for market bids. For buyouts it is NOT confirmed to be the
          same type -- this API has no official documentation, so before
          passing confirm=True for a buyout, compare the dry-run payload
          with the actual request you see in your browser's DevTools when
          buying out a clause from biwenger.com.
        - confirm: defaults to False -> nothing is sent, it only prints and
          returns the payload that would be sent (dry run). Pass
          confirm=True to actually execute the offer (spends/commits your
          balance).
        """
        payload = {
            "amount": int(amount),
            "requestedPlayers": [int(player_id)],
            "to": to,
            "type": offer_type,
        }
        if not confirm:
            print("[DRY RUN] Nothing has been sent yet. Payload that would "
                  "be sent to POST /offers/:")
            print(json.dumps(payload, indent=2, ensure_ascii=False))
            print("Call again with confirm=True to execute the real offer.")
            return {"dry_run": True, "payload": payload}
        return self._post(f"{BASE_URL}/offers/", json_body=payload)

    def get_team(self, team_id) -> dict:
        return self._get(
            f"{BASE_URL}/user/{team_id}",
            params={
                "fields": (
                    "*,account(id),players(id,owner),"
                    "lineups(round,points,count,position),"
                    "league(id,name,competition,mode,scoreID),"
                    "market,seasons,offers,lastPositions"
                )
            },
        )

    # -------------------------------------------------------------- players

    # cf.biwenger.com is public (no login required) and returns 403 if it
    # receives the biwenger.as.com session headers (Authorization/x-league/
    # x-user), so they must be explicitly cleared in these requests.
    _CF_HEADERS = {"Authorization": None, "x-league": None, "x-user": None}

    def get_all_players_snapshot(self) -> dict:
        """Public listing (no login required) of all LaLiga players."""
        return self._get(
            f"{CF_URL}/competitions/la-liga/data",
            params={"lang": self.lang, "score": 2},
            headers=self._CF_HEADERS,
        )

    def get_player_detail(self, player_id) -> dict:
        """Player detail. Includes 'seasons' with the historical data."""
        return self._get(
            f"{CF_URL}/players/la-liga/{player_id}",
            params={
                "fields": (
                    "*,team,fitness,reports(points,home,events,"
                    "status(status,statusInfo),match(*,round,home,away),star),"
                    "prices,competition,seasons,news,threads"
                ),
                "score": 5,
                "lang": self.lang,
            },
            headers=self._CF_HEADERS,
        )
