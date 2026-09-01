"""
Cliente no oficial para la API de Biwenger.

Biwenger no publica documentacion oficial de su API, asi que este cliente
esta construido a partir de la estructura de endpoints que usan varios
proyectos de la comunidad (biwenger.as.com / cf.biwenger.com). Por eso es
deliberadamente defensivo: si algun campo esperado no aparece en una
respuesta, no revienta el programa entero, avisa y sigue con lo que puede.

Antes de confiar del todo en los calculos de 02_generar_excel.py, ejecuta
01_explorar_api.py una vez: vuelca respuestas reales a output/debug/ para
poder confirmar los nombres de campo exactos (sobre todo el historico de
temporadas anteriores en 'seasons').
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
    """Decodifica (sin verificar firma) el payload de un JWT y devuelve un claim.

    La respuesta de login de Biwenger a veces solo trae el 'token' (sin objeto
    'user'), pero el id de usuario va dentro del propio JWT como claim 'iss'.
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
    """Navega un dict anidado devolviendo None si algun paso no existe."""
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
                "No se ha podido extraer el token de la respuesta de login. "
                f"Respuesta recibida: {json.dumps(payload, ensure_ascii=False)[:800]}"
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
        """Detecta la liga (y el team-id que hay que mandar en x-user para
        esa liga concreta) a partir de /account.

        La API exige que x-user sea el id de EQUIPO dentro de la liga
        (leagues[].user.id), no el id de cuenta global que trae el JWT del
        login. Sin esto, /league devuelve 401 "Invalid user".
        """
        leagues = self.discover_leagues()
        match = None
        if league_id:
            for l in leagues:
                if str(l.get("id")) == str(league_id):
                    match = l
                    break
            if match is None:
                # No lo hemos podido confirmar contra /account (¿liga de otra
                # cuenta amiga?): seguimos con el league_id dado tal cual.
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
                print(f"    (limite de peticiones, esperando {wait}s...)")
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
        raise last_error or BiwengerAPIError(f"Demasiados 429 al pedir {url}")

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
                f"Error {resp.status_code} en {context}: {resp.text[:500]}"
            )

    # ------------------------------------------------------- account/liga

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

    # ------------------------------------------------------ mercado/equipos

    def get_market(self) -> dict:
        return self._get(f"{BASE_URL}/market")

    def find_market_seller(self, player_id, market=None):
        """Busca en el mercado quien vende `player_id` y devuelve su user_id
        (el que hay que pasar como `to` en place_offer para pujar por el).

        Si ya tienes el resultado de get_market() a mano pasalo en `market`
        para no volver a pedirlo.
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
        """Envia una oferta/puja por un jugador: compra de un jugador que
        esta en el mercado, o clausulazo a un jugador de otro manager de tu
        liga que NO esta en el mercado.

        - player_id: id del jugador.
        - amount: importe de la oferta (o el precio de la clausula, para un
          clausulazo).
        - to: user_id del actual propietario del jugador. Para uno que esta
          en el mercado puedes obtenerlo con find_market_seller(player_id);
          para un clausulazo es el user.id que ves en standings/tu roster.
        - offer_type: "purchase" es el tipo confirmado por varios clientes
          no oficiales para pujas de mercado. Para clausulazos NO esta
          confirmado que sea el mismo tipo -- esta API no tiene
          documentacion oficial, asi que antes de pasar confirm=True para
          un clausulazo, compara el payload en dry-run con la peticion real
          que ves en el DevTools de tu navegador al clausular desde
          biwenger.com.
        - confirm: por defecto False -> no se manda nada, solo se imprime
          y devuelve el payload que se enviaria (dry-run). Pasa confirm=True
          para ejecutar la oferta de verdad (gasta/compromete tu saldo).
        """
        payload = {
            "amount": int(amount),
            "requestedPlayers": [int(player_id)],
            "to": to,
            "type": offer_type,
        }
        if not confirm:
            print("[DRY RUN] No se ha enviado nada todavia. Payload que se "
                  "mandaria a POST /offers/:")
            print(json.dumps(payload, indent=2, ensure_ascii=False))
            print("Llama de nuevo con confirm=True para ejecutar la oferta real.")
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

    # -------------------------------------------------------------- jugadores

    # cf.biwenger.com es publico (sin login) y devuelve 403 si le llegan las
    # cabeceras de sesion de biwenger.as.com (Authorization/x-league/x-user),
    # asi que hay que anularlas explicitamente en estas peticiones.
    _CF_HEADERS = {"Authorization": None, "x-league": None, "x-user": None}

    def get_all_players_snapshot(self) -> dict:
        """Listado publico (sin login) de todos los jugadores de LaLiga."""
        return self._get(
            f"{CF_URL}/competitions/la-liga/data",
            params={"lang": self.lang, "score": 2},
            headers=self._CF_HEADERS,
        )

    def get_player_detail(self, player_id) -> dict:
        """Detalle de un jugador. Incluye 'seasons' con el historico."""
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
