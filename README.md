# Análisis Biwenger

Herramientas en Python para analizar tu liga de [Biwenger](https://biwenger.as.com/) (fútbol fantasy) más allá de lo que ofrece la app: ratios de rendimiento por precio, un "estado de forma" ponderado por partidos recientes, estimación de valor justo de mercado (chollos/sobreprecios), sugerencias de fichajes y ventas con explicación de a quién sustituyen, y hasta ejecutar pujas/clausulazos directamente desde código (en modo seguro, sin sorpresas).

Todo se construye a partir de la **API no oficial de Biwenger** (no hay documentación pública), inspeccionando respuestas reales. Si Biwenger cambia algo, es el sitio por el que empezar a mirar — ver [Cómo funciona por dentro](#cómo-funciona-por-dentro).

> ⚠️ Este es un proyecto personal, no afiliado a Biwenger/AS.com. Úsalo bajo tu responsabilidad: es una API no documentada y puede romperse si Biwenger cambia su backend. La parte de "operaciones" (pujar/clausular) mueve dinero real de tu equipo — lee la sección de [seguridad](#seguridad-al-pujarclausular) antes de usarla.

## Índice

- [¿Qué hace exactamente?](#qué-hace-exactamente)
- [Instalación](#instalación)
- [Configuración](#configuración)
- [Guía rápida](#guía-rápida)
- [El notebook, sección a sección](#el-notebook-sección-a-sección)
- [El DataFrame: columnas explicadas](#el-dataframe-columnas-explicadas)
- [Referencia de funciones](#referencia-de-funciones-biwenger_helperspy)
- [Seguridad al pujar/clausular](#seguridad-al-pujarclausular)
- [Cómo funciona por dentro](#cómo-funciona-por-dentro)
- [Estructura del proyecto](#estructura-del-proyecto)
- [Solución de problemas](#solución-de-problemas)

## ¿Qué hace exactamente?

Biwenger te enseña un jugador a la vez. Este proyecto descarga **todo** lo relevante de tu liga (mercado, plantilla de cada rival, detalle de cada jugador) y lo junta en una única tabla que puedes ordenar, filtrar y cruzar como quieras. A partir de ahí:

- **Ratios de rendimiento por precio**: puntos / valor de mercado, puntos / clausula (temporada actual y anterior), tanto en total como por partido.
- **`Puntuación potencial`**: una métrica propia que mezcla la forma reciente (últimos partidos, ponderados — cuenta más el más reciente) con el rendimiento de la temporada pasada, para no fiarlo todo a una racha de 2 partidos.
- **Valor justo de mercado**: para cada jugador en venta, estima si está barato o caro comparándolo con jugadores similares de tu propia liga (no un precio "oficial" externo).
- **Cuánto ofertar**: combina tu valor justo estimado, cuántos rivales podrían pujar más que tú, y lo que se ha pagado de verdad por jugadores parecidos en tu liga — nunca sugiere pujar por debajo del precio pedido (Biwenger no lo permite).
- **Sugerencia de fichajes**: compara cada candidato asequible (mercado o clausula) contra tu jugador más flojo, sin sesgo de "esta posición ya está cubierta" — mide mejora real y dice a quién sustituirías.
- **Sugerencia de ventas**: quién de tu plantilla está sobrevalorado ahora mismo, cruzado con ofertas reales que ya has recibido.
- **Gráficos**: chollos/sobrevalorados con outliers marcados, mejores/peores por posición, distribución de rendimiento.
- **Caché resumible**: la descarga inicial (mercado + cada plantilla + detalle de cada jugador) puede tardar varios minutos por el límite de peticiones de la API. Se guarda en disco de forma incremental, así que si se corta a medias, la siguiente ejecución retoma solo lo que falta.
- **Pujar/clausular de verdad**: en modo *dry-run* por defecto (solo te enseña qué se mandaría), con `confirm=True` explícito para ejecutar.

## Instalación

Necesitas **Python 3.10+**.

```bash
git clone <la-url-de-tu-fork>
cd biwenger-analisis
python -m venv .venv

# Windows
.venv\Scripts\activate
# macOS / Linux
source .venv/bin/activate

pip install -r requirements.txt
```

## Configuración

Copia el fichero de ejemplo y rellena tus credenciales:

```bash
copy .env.example .env      # Windows
cp .env.example .env        # macOS / Linux
```

Abre `.env` con un editor de texto:

```env
BIWENGER_EMAIL=tu_email@ejemplo.com
BIWENGER_PASSWORD=tu_contraseña

# Opcional: solo hace falta si tu cuenta esta en varias ligas
BIWENGER_LEAGUE_ID=

# Opcional: solo hace falta si el auto-detect no encuentra cual eres tu
BIWENGER_OWN_TEAM_ID=
```

`.env` está en `.gitignore` — **nunca se sube a ningún sitio**, se queda solo en tu máquina. El login se hace en cada ejecución porque el token de sesión de Biwenger caduca.

## Guía rápida

### Opción 1: el notebook interactivo (recomendado)

```bash
python 01_explorar_api.py   # valida que el login y la liga funcionan (una vez, o si algo falla)
```

Abre `03_analisis.ipynb` en VS Code o Jupyter, selecciona el intérprete de `.venv` como kernel, y ejecuta las celdas de arriba a abajo. La primera carga de datos tarda unos minutos (ver [caché resumible](#cómo-funciona-por-dentro)); a partir de ahí es instantánea.

### Opción 2: Excel estático

```bash
python 02_generar_excel.py            # usa la cache si ya existe
python 02_generar_excel.py --refresh  # fuerza descarga completa
```

Genera `output/analisis_biwenger_<fecha>.xlsx` con tres hojas (Mercado, Rivales, Mi equipo), columnas fijadas y formato condicional (rojo→amarillo→verde) sobre los ratios. Rápido de compartir o imprimir, pero sin gráficos ni recomendaciones — para eso usa el notebook. Comparte la misma caché que el notebook, así que si ya cargaste datos desde uno de los dos, el otro los reutiliza al instante.

## El notebook, sección a sección

### 1. Cargar datos

```python
data = biwenger_helpers.cargar_datos_liga(
    EMAIL, PASSWORD, league_id=LEAGUE_ID, own_team_id=OWN_TEAM_ID,
    forzar_refresco=False, parar_en_primer_limite=True,
)
```

La primera vez descarga mercado + plantillas de todos los managers + detalle de cada jugador implicado, y lo guarda en `output/cache/*.json`. Es **resumible**: con `parar_en_primer_limite=True` la carga se para sola en cuanto la API corta por exceso de peticiones (en vez de quedarse esperando en pausas largas), y la siguiente vez que ejecutes la celda retoma solo lo que falte. Si la caché tiene más de 24h (`max_antiguedad_horas`) se refresca sola.

### 2. Explorar

`df` es una fila por jugador — mercado + cada rival + tu equipo — con la columna `Origen` para filtrar. Ejemplos:

```python
# Top 20 gangas del mercado por ratio puntos/valor de mercado
df[df["Origen"] == "Mercado"].dropna(subset=["Ratio pts/VM (actual)"]) \
    .sort_values("Ratio pts/VM (actual)", ascending=False).head(20)

# Tu plantilla ordenada por rendimiento (para ver tus jugadores más flojos)
df[df["Origen"] == "Mi equipo"].sort_values("Puntuacion potencial")

# Mejores clausulas de rivales por ratio puntos/clausula
df[df["Origen"].str.startswith("Rival:")] \
    .dropna(subset=["Ratio pts totales/clausula (actual)"]) \
    .sort_values("Ratio pts totales/clausula (actual)", ascending=False).head(20)
```

**Importante sobre `Origen`**: `'Mercado'` significa jugador *libre de verdad* (sin dueño, comprable directo pujando). Las entradas del mercado que sí tienen dueño (alguien puso a su jugador en venta) **no** cuentan como `'Mercado'` — la única vía real de ficharlos es el clausulazo, así que aparecen como `'Rival: <nombre>'` con su columna `Clausula`. Además, mercado/rivales ya excluyen "morralla" automáticamente (sin equipo de 1ª, valor ≤ 400.000€, o inactivos); tu propio equipo nunca se filtra.

### 3. Gráficos

- **Top 20 por ratio pts/VM** (barras horizontales).
- **Distribución del ratio por posición** (boxplot).
- **Valor vs. rendimiento**: scatter con línea de tendencia; los 8 jugadores que más se desvían por arriba (chollos) y por abajo (sobrevalorados) salen rotulados con su nombre.
- **Mejores y peores 5 por posición**: cuatro paneles (Portero/Defensa/Centrocampista/Delantero), verde arriba y rojo abajo.

### 4. Sugerencia de fichajes

```python
biwenger_helpers.sugerir_fichajes(df, data["balance"], top_n=15)

# Priorizando eficiencia (mejora por millón) en vez de mejora absoluta
biwenger_helpers.sugerir_fichajes(df, data["balance"], top_n=15, ordenar_por="eficiencia")
```

Compara cada candidato asequible y disponible contra tu jugador más flojo **de esa misma posición**, y solo muestra candidatos que realmente mejorarían a ese jugador. No mete ningún sesgo de "esta línea ya está cubierta" — mide mejora real y te dice explícitamente a quién sustituirías (columna `Jugador que reemplazarías`). Descarta automáticamente clausulas bloqueadas temporalmente tras una compra reciente.

Columnas clave: `Saldo tras fichar`, `Jugadores flojos en tu posicion` / `Total en tu posicion` (cuántos de los tuyos rinden por debajo de tu propia mediana ahí), `Mejora por millon gastado`.

> 💡 Es una sugerencia orientativa a partir del `balance` que le pases — no tiene en cuenta el `maximumBid` real de Biwenger (que puede ser bastante mayor que tu saldo, ya que Biwenger permite pujar en descubierto usando el valor de tu plantilla como colateral). Si quieres usar tu límite real de puja, pásalo tú mismo en vez de `data["balance"]`.

### 5. Ranking general

```python
biwenger_helpers.mejores_jugadores(df, top_n=20)                                    # top 20 general
biwenger_helpers.mejores_jugadores(df, posicion="Delantero", origen="Mercado")       # filtrado
```

Una watchlist de quién está rindiendo mejor ahora mismo (por `Puntuacion potencial`), sin mirar precio ni si te lo puedes permitir. Excluye lesionados/sancionados por defecto.

### 6. Valor justo de mercado

```python
valor_justo = biwenger_helpers.estimar_valor_justo(df, margen_pct=20)
valor_justo[valor_justo["Valoracion"] == "Chollo"]

# Cuánto ofertar por un jugador libre del mercado
biwenger_helpers.sugerir_oferta_mercado(df, data)
```

Para cada jugador en venta, `estimar_valor_justo` calcula la mediana de rendimiento/precio de jugadores similares de tu liga (misma posición) y la usa de referencia para etiquetar Chollo/Caro/En la media.

`sugerir_oferta_mercado` va un paso más allá y calcula cuánto haría falta ofertar de verdad para llevarte al jugador (no solo cuánto "vale" según tu heurística): combina el precio pedido (**suelo obligatorio** — Biwenger no deja pujar por debajo), lo que se ha pagado de verdad por jugadores comparables en tu liga, y cuántos rivales tienen presupuesto de sobra para pujar más que tú. Es orientativo — no sabe quién está *realmente* interesado en cada jugador, solo quién podría permitírselo económicamente; un rival puede pagar muy por encima de cualquier estimación razonable.

### 7. Sugerencia de ventas

```python
ventas = biwenger_helpers.sugerir_ventas(df, data=data, margen_pct=20)
ventas[ventas["Recomendacion"] == "Vender ahora"]
```

El complementario de la sección 6 para tu plantilla: jugadores cuyo valor de mercado se ha disparado muy por encima de lo que su rendimiento real justificaría. Cruza también ofertas reales ya recibidas (filtrando las ofertas automáticas "Desconocido" que Biwenger genera para cada jugador tuyo, que no son demanda real).

### 8. Ofertas

```python
biwenger_helpers.ofertas(data, tipo="recibidas")
biwenger_helpers.ofertas(data, tipo="enviadas")
```

Lista las ofertas de compra activas en el mercado sin tener que entrar en la app (no incluye aceptar/rechazar, eso se hace desde Biwenger).

### 9. Historial

```python
biwenger_helpers.guardar_snapshot(data)          # guarda una copia fechada
biwenger_helpers.listar_snapshots()
biwenger_helpers.cargar_snapshot("20260828_120000")
```

A diferencia de la caché (que se sobrescribe), `guardar_snapshot` guarda una copia con fecha en `output/history/`, útil para comparar la evolución de tu equipo o del mercado con el tiempo.

### 10. Operaciones (pujar / clausular)

Acciones reales sobre tu liga — ver la sección de [seguridad](#seguridad-al-pujarclausular) antes de tocar esto.

```python
cliente_operaciones = BiwengerClient(EMAIL, PASSWORD, league_id=LEAGUE_ID)
biwenger_helpers.login_y_resolver_liga(cliente_operaciones, LEAGUE_ID)

# 'Vendedor (id)' de las tablas de arriba es lo que va en 'to'
cliente_operaciones.place_offer(player_id, importe, to=vendedor_id)                    # dry-run: solo imprime el payload
cliente_operaciones.place_offer(player_id, importe, to=vendedor_id, confirm=True)      # ejecuta de verdad
```

## El DataFrame: columnas explicadas

| Columna | Qué es |
|---|---|
| `player_id` | ID interno de Biwenger para el jugador |
| `Jugador`, `Posicion`, `Equipo LaLiga` | Datos básicos |
| `Estado`, `Disponible` | `Disponible=False` si está lesionado/sancionado/descartado (`'doubt'` no cuenta como no disponible) |
| `Valor de mercado` | Valor de mercado oficial actual, en euros |
| `Tendencia precio (7d) %` | Variación del valor de mercado en los últimos 7 días |
| `Temporada actual` / `anterior` | Etiqueta real de la temporada (p.ej. "Temporada 2025/2026"); el nombre de columna es siempre el mismo, calculado por mayoría entre todos los jugadores para no desalinear a quien no jugó recientemente |
| `Puntos temporada actual` / `anterior` | Puntos totales acumulados |
| `Partidos temporada actual` / `anterior` | Partidos jugados |
| `Pts/partido (actual)` / `(anterior)` | Puntos totales ÷ partidos |
| `Ratio pts/VM (actual)` / `(anterior)` | Puntos totales ÷ valor de mercado (en millones) |
| `Forma (ult. partidos)` | Media ponderada de los últimos partidos, dando más peso a los más recientes (decay exponencial) |
| `Puntuacion potencial` | `Forma` (70%) + puntos/partido de la temporada anterior (30%) — la métrica que usan los rankings y recomendadores |
| `Proximo rival`, `Local/Visitante`, `Dificultad proximo partido`, `Jornada proxima` | Su siguiente partido de LaLiga |
| `Origen` | `'Mercado'` (libre), `'Mi equipo'`, o `'Rival: <nombre>'` |
| `Precio en mercado`, `Vendedor (id)` | Solo si `Origen == 'Mercado'` |
| `Clausula`, `Clausula bloqueada hasta`, `Clausula disponible ahora` | Solo para jugadores de rivales/tuyos |
| `Ratio pts totales/clausula` / `Ratio pts medios/clausula` (`actual`/`anterior`) | El mismo ratio de antes pero contra la clausula en vez del valor de mercado — por puntos totales o por puntos medios (un jugador con pocos partidos puede salir mal en el total pero bien en la media) |

## Referencia de funciones (`biwenger_helpers.py`)

### Carga y caché

| Función | Qué hace |
|---|---|
| `cargar_datos_liga(email, password, league_id=None, own_team_id=None, forzar_refresco=False, max_antiguedad_horas=24, parar_en_primer_limite=False)` | Orquesta todo: login, mercado, plantillas, detalle de jugadores, caché resumible. Devuelve el dict `data` que usan el resto de funciones |
| `construir_dataframe(data)` | Convierte `data` en el DataFrame combinado que se usa en todo el notebook |
| `guardar_snapshot(data)` / `listar_snapshots()` / `cargar_snapshot(nombre)` | Histórico fechado en `output/history/` |

### Análisis y recomendación

| Función | Qué hace |
|---|---|
| `mejores_jugadores(df, posicion=None, origen=None, top_n=20, metrica="Puntuacion potencial", solo_disponibles=True)` | Ranking/watchlist filtrable |
| `estimar_valor_justo(df, metrica="Puntuacion potencial", margen_pct=20, solo_disponibles=True)` | Chollo/Caro/En la media para el mercado |
| `sugerir_oferta_mercado(df, data, margen_sobre_pedido_pct=5, metrica="Puntuacion potencial", rango_comparables_pct=40)` | Cuánto ofertar por un jugador libre del mercado |
| `sugerir_fichajes(df, balance, top_n=15, solo_disponibles=True, ordenar_por="mejora")` | Candidatos que mejoran a tu jugador más flojo de su posición (`ordenar_por="eficiencia"` para priorizar coste/beneficio) |
| `sugerir_ventas(df, data=None, metrica="Puntuacion potencial", margen_pct=20)` | Jugadores tuyos sobrevalorados, cruzado con ofertas reales |
| `coste_por_punto(df, metrica="Puntuacion potencial", margen_pct=20)` | Coste por punto de cada jugador frente a la mediana de TODO el juego (todas las posiciones), para saber si está caro/barato en términos absolutos |
| `ofertas(data, tipo="recibidas", solo_identificadas=False)` | Ofertas de compra activas (`solo_identificadas=True` descarta las automáticas de Biwenger, que no son demanda real) |

### Piezas internas (por si tocas algo)

| Función | Qué hace |
|---|---|
| `puntuacion_potencial(detalle, score_id, decay=0.85, n_partidos=6, peso_temporada_anterior=0.3, ...)` | La métrica base — ajusta aquí los pesos si quieres que pese más o menos la forma reciente |
| `calcular_forma(detalle, score_id, decay=0.85, n_partidos=6)` | Media ponderada de los últimos N partidos |
| `jugador_irrelevante(detalle, score_id, ...)` | El filtro de "morralla" (sin equipo de 1ª, valor bajo, inactivo) |
| `temporada_liga_mas_comun(detalles)` | Detecta la temporada "actual" real por mayoría entre todos los jugadores, para no arrastrar temporadas antiguas de quien no ha debutado |

## Seguridad al pujar/clausular

`client.place_offer(...)` funciona en **modo dry-run por defecto**: sin `confirm=True` no se envía nada a Biwenger, solo se imprime el payload que se mandaría (jugador, importe, a quién). Repasa siempre ese payload antes de confirmar — pasar `confirm=True` ejecuta la oferta de verdad, gasta o compromete tu saldo, y **no se puede deshacer** desde aquí.

El campo `to` es el `user_id` del propietario actual del jugador (coincide con el `team_id` que ves en la clasificación). Para un jugador del mercado usa `client.find_market_seller(player_id)` o la columna `Vendedor (id)` de las tablas del notebook; para un clausulazo a un rival, usa directamente su `team_id`.

El tipo de oferta (`offer_type="purchase"` por defecto) está confirmado para pujas de mercado por varios clientes no oficiales de Biwenger, pero **no** para clausulazos a un rival — antes de confirmar un clausulazo por primera vez, compara el payload del dry-run con la petición real que ves en el DevTools del navegador al hacerlo desde biwenger.com.

## Cómo funciona por dentro

- **API no oficial**: se usa `biwenger.as.com/api/v2` (autenticado, para liga/mercado/plantillas/pujas) y `cf.biwenger.com/api/v2` (público, para el detalle de cada jugador). No hay documentación oficial; los nombres de campo están sacados de inspeccionar respuestas reales — si algo deja de funcionar, ejecuta `01_explorar_api.py` y compara `output/debug/*.json` con lo que espera `biwenger_helpers.py`.
- **`x-user` es el ID de equipo, no de cuenta**: la API exige que la cabecera `x-user` sea tu id de equipo *dentro de esa liga* (`leagues[].user.id`), no el id de cuenta global del token de login. `resolve_league()` lo resuelve automáticamente.
- **`cf.biwenger.com` rechaza cabeceras de sesión**: las peticiones de detalle de jugador anulan explícitamente `Authorization`/`x-league`/`x-user`, si no, responde 403.
- **Límite de peticiones**: `cf.biwenger.com` corta alrededor de la petición ~200-211 de una tirada. No es una ventana corta que se resetee esperando (probamos una pausa preventiva de 5 minutos y el corte seguía saltando en el mismo punto) — es un tope de peticiones totales. La única forma real de evitarlo del todo es no repetir peticiones, de ahí la caché compartida y resumible: el detalle de cada jugador se guarda en cuanto se consigue, así que una carga interrumpida retoma solo lo que falta.
- **`maximumBid` puede superar tu saldo**: Biwenger te deja pujar en descubierto usando el valor de tu plantilla como colateral. `sugerir_fichajes` usa el `balance` que le pases tal cual — si quieres tu límite real de puja, pásale `maximumBid` (de `standings`) en vez de `data["balance"]`.
- **Puntos por sistema de puntuación**: cada partido trae los puntos desglosados por `scoreID` (distintas ligas pueden puntuar distinto). Todo el análisis usa automáticamente el `scoreID` de tu liga.

## Estructura del proyecto

```
biwenger_client.py      Cliente de la API: login, lectura, y escritura (pujar/clausular)
biwenger_helpers.py     Toda la lógica de análisis y carga/caché de datos
01_explorar_api.py      Diagnóstico: valida login/liga y vuelca respuestas reales a output/debug/
02_generar_excel.py     Genera un Excel estático (mercado + rivales + tu equipo)
03_analisis.ipynb       Notebook interactivo — la forma recomendada de usar esto día a día
.env.example            Plantilla de credenciales (copia a .env, nunca se sube)
output/                 Generado en local (caché, debug, Excel, snapshots) — en .gitignore
```

## Solución de problemas

- **`401 Invalid email or password`**: revisa `.env`. Si tus credenciales son correctas y sigue fallando, puede ser un bloqueo temporal por demasiados logins seguidos — espera unos minutos.
- **`400 X-League and X-User headers required` / `401 Invalid user`**: normalmente se resuelve solo con `resolve_league()` / `login_y_resolver_liga()`. Si tienes varias ligas, fija `BIWENGER_LEAGUE_ID` en `.env`.
- **Avisos de "campo no encontrado"**: ejecuta `01_explorar_api.py`, mira el JSON real en `output/debug/`, y ajusta `CAMPOS_CANDIDATOS` en `biwenger_helpers.py`.
- **La carga de datos se corta / va lenta**: es el límite de peticiones (ver [arriba](#cómo-funciona-por-dentro)) — no hay que hacer nada especial, vuelve a ejecutar la celda/script más tarde y retomará solo lo que falte.
- **`ModuleNotFoundError`**: asegúrate de que el intérprete activo es el de `.venv` (`.venv\Scripts\python.exe` en Windows) y no un Python global.
- **El notebook no refleja cambios en `biwenger_helpers.py`**: reinicia el kernel — Python cachea los módulos ya importados.
