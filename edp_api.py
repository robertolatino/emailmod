"""
edp_api.py — Cliente de la API interna del backoffice de Edelvives Digital Plus.

QUÉ ES ESTE MÓDULO
------------------
El backoffice (https://publisher.edelvivesdigitalplus.com) es una aplicación web
de una sola página (SPA, hecha en React). Todo lo que aparece en pantalla lo pide
el navegador a una API JSON del mismo dominio, bajo /api/. Por ejemplo, al pulsar
"Guardar" en la ficha de un usuario, el navegador envía PUT /api/users/<guid>
con los datos del formulario.

Este módulo hace esas mismas peticiones directamente, con el cliente HTTP que
trae Playwright (APIRequestContext). El efecto es idéntico a usar la web a mano,
pero sin abrir páginas ni hacer clics: es mucho más rápido y no se rompe si cambia
el diseño de las pantallas.

Los endpoints se obtuvieron leyendo el código JavaScript público del backoffice y
se comprobaron contra la plataforma real con usuarios de prueba. El detalle está
en DOCUMENTACION_TECNICA.md.

ENDPOINTS QUE SE USAN
---------------------
  POST /api/login                   Inicio de sesión. Devuelve un token (JWT).
  GET  /api/users/whoami            Datos de la cuenta con la que se ha entrado.
  GET  /api/users?usernames[]=...   Búsqueda exacta por correo de acceso, sin distinguir
                                    mayúsculas. Hasta 156 correos por petición.
  GET  /api/users/{guid}            Ficha completa de un usuario.
  PUT  /api/users/{guid}            Guardar la ficha (username, email, name, lastname,
                                    role_guid, password, change_password...). Si el
                                    correo ya lo usa otro usuario: HTTP 409 USERNAME_EXIST.
  GET  /api/users/{guid}/license-personals
                                    Licencias redimidas por el usuario (pestaña
                                    "Licencias" de su ficha), con su fecha de caducidad.

PIEZAS PRINCIPALES
------------------
  ClienteEDP           La clase que habla con la API: login, búsqueda, lectura de
                       fichas y licencias, y edición. Anota cada petición en `registro`.
  ErrorLogin           No se ha podido iniciar sesión.
  ErrorAPI             La API ha respondido con un error.
  SesionRevocada       Alguien ha entrado con la misma cuenta y ha cerrado esta sesión.
  cargar_env, agregar_args_conexion, conectar
                       Utilidades que comparten los scripts para leer las credenciales
                       del fichero .env e iniciar sesión.

USO DIRECTO (solo lectura, para comprobar la conexión)
------------------------------------------------------
  python3 edp_api.py correo1@dominio.com correo2@dominio.com
"""
from __future__ import annotations

import argparse
import json
import os
import sys
import time
from datetime import datetime, timezone
from pathlib import Path
from typing import Iterable
from urllib.parse import urlencode

from playwright.sync_api import Error as PlaywrightError
from playwright.sync_api import Playwright, sync_playwright

URL_DEFECTO = "https://publisher.edelvivesdigitalplus.com"

# Roles con los que la pantalla de login del backoffice deja entrar
# (admin, helpdesk, admin fixed layout y publisher). La API los exige en el login.
ROLES_BACKOFFICE = ["R04", "R05", "R06", "R07"]

# Códigos de rol de la plataforma y su nombre legible (para los informes).
NOMBRES_ROL = {"R01": "alumno", "R02": "profesor", "R03": "admin colegio", "R04": "admin",
               "R05": "helpdesk", "R06": "admin fixed layout", "R07": "publisher"}

# Correos por petición de búsqueda. La propia web limita a 156; usamos 150 por margen.
MAX_POR_BUSQUEDA = 150

# Respuestas que indican un problema temporal del servidor: merece la pena reintentar.
ESTADOS_REINTENTABLES = {429, 500, 502, 503, 504}


# --- Errores ------------------------------------------------------------------

class ErrorLogin(Exception):
    """No se ha podido iniciar sesión (credenciales, sesión abierta en otro sitio, red...)."""


class ErrorAPI(Exception):
    """La API ha respondido con un error.

    Atributos:
        status:    código HTTP de la respuesta (409, 404...).
        cuerpo:    cuerpo de la respuesta tal cual (normalmente un dict).
        error_key: código de error de la plataforma, p. ej. "USERNAME_EXIST".
        errno:     código de error de base de datos, si lo hay (1062 = duplicado).
    """

    def __init__(self, status: int, cuerpo, metodo: str, ruta: str):
        self.status = status
        self.cuerpo = cuerpo
        info = extraer_error(cuerpo)
        self.error_key = info.get("errorKey")
        self.errno = info.get("errno")
        detalle = " - ".join(str(v) for v in (self.error_key or info.get("code"), info.get("message"))
                             if v) or cuerpo
        super().__init__(f"{metodo} /api/{ruta} -> HTTP {status}: {detalle}")

    @property
    def usuario_ya_existe(self) -> bool:
        """True si el error significa "ese correo ya lo usa otro usuario".

        La web muestra "El usuario ya existe" con cualquiera de estos dos códigos.
        En la plataforma real se recibe HTTP 409 con errorKey USERNAME_EXIST.
        """
        return self.error_key == "USERNAME_EXIST" or str(self.errno) == "1062"


class SesionRevocada(ErrorAPI):
    """La plataforma solo admite una sesión por cuenta. Si alguien entra con la misma
    cuenta en otro sitio (por ejemplo, en el navegador) forzando el acceso, la sesión
    del script se cierra y la API responde 403 SESSION_REVOKED. Los scripts paran
    ordenadamente y guardan el informe hasta ese punto."""


def extraer_error(cuerpo) -> dict:
    """Saca {errorKey, errno, message, code} del cuerpo de una respuesta de error.

    La API usa varias formas según el tipo de error:
      {"status": "fail", "error": {"code": 409, "errorKey": "USERNAME_EXIST"}}
      {"message": "Authorization failed", "status": "fail", "data": {"error": 401}}
      {"code": "SCHEMA_VALIDATION_FAILED", "message": ..., "results": {"errors": [...]}}
    """
    if not isinstance(cuerpo, dict):
        return {"message": str(cuerpo)[:300]} if cuerpo else {}
    info = dict(cuerpo["error"]) if isinstance(cuerpo.get("error"), dict) else {}
    for clave in ("errorKey", "errno", "message", "code"):
        if cuerpo.get(clave) is not None:
            info.setdefault(clave, cuerpo[clave])
    errores = (cuerpo.get("results") or {}).get("errors") or []
    if errores:  # errores de validación: indicamos qué campo ha fallado y por qué
        info["message"] = "; ".join(f"{'.'.join(map(str, e.get('path', [])))}: {e.get('message')}"
                                    for e in errores)
    return info


def token_rechazado(status: int, cuerpo) -> bool:
    """True si la API rechaza el token (caducado o ausente).

    Ojo: en ese caso la API no responde 401 sino HTTP 403 con
    {"status": "fail", "data": {"error": 401}}.
    """
    datos = cuerpo.get("data") if isinstance(cuerpo, dict) else None
    return status == 401 or (status == 403 and isinstance(datos, dict) and datos.get("error") == 401)


def _json(respuesta):
    """Cuerpo de la respuesta como JSON, o como texto si no es JSON."""
    try:
        return respuesta.json()
    except Exception:
        return respuesta.text()[:500]


# --- Datos de usuario -----------------------------------------------------------

def datos_edicion(usuario: dict) -> dict:
    """Campos que la ficha del backoffice reenvía siempre al pulsar "Guardar", con sus
    valores actuales.

    Al editar enviamos estos campos sin tocar más los que queremos cambiar. Así el
    resto de datos (nombre, apellidos, rol, demo, país) se reenvía con su mismo
    valor y no cambia. Se ha comprobado comparando la ficha completa antes y después.
    Los campos vacíos (None) no se envían.
    """
    datos = {
        "username": usuario.get("username"),
        "name": usuario.get("name"),
        "lastname": usuario.get("lastname"),
        "role_guid": usuario.get("role_guid"),
        "is_demo": 1 if usuario.get("is_demo") else 0,
        "country_guid": usuario.get("country_guid"),
    }
    return {k: v for k, v in datos.items() if v is not None}


def describir_cambios(actual: dict, cambios: dict) -> str:
    """Texto legible de un cambio para el registro, p. ej. "username: a@x.com → b@x.com".
    La contraseña nunca se escribe."""
    partes = []
    for campo, valor in cambios.items():
        if campo == "password":
            partes.append("contraseña: (oculta)")
        else:
            anterior = actual.get(campo)
            partes.append(f"{campo}: {'(vacío)' if anterior in (None, '') else anterior} → {valor}")
    return "; ".join(partes)


def normalizar_correo(valor) -> str:
    """Correo en minúsculas y sin espacios alrededor ("" si está vacío). Las búsquedas
    de la API no distinguen mayúsculas, así que comparamos siempre en minúsculas."""
    return str(valor or "").strip().lower()


def licencia_activa(licencia: dict, ahora: datetime) -> bool:
    """True si la licencia no ha caducado.

    Es el mismo criterio que la columna "Estado" de la pestaña Licencias de la ficha,
    que muestra los días que quedan hasta expired_at o "Caducada" si ya ha pasado.
    expired_at viene en UTC, p. ej. "2027-09-30T00:00:00.000Z". Sin fecha de caducidad,
    o con una fecha que no se entiende, se considera activa: ante la duda, mejor no
    tocar al usuario.
    """
    texto = str(licencia.get("expired_at") or "").strip()
    if not texto:
        return True
    try:
        caduca = datetime.fromisoformat(texto.replace("Z", "+00:00"))
    except ValueError:
        return True
    if caduca.tzinfo is None:
        caduca = caduca.replace(tzinfo=timezone.utc)
    return caduca > ahora


# --- Cliente --------------------------------------------------------------------

class ClienteEDP:
    """Cliente de la API del backoffice.

    Guarda el token de sesión y lo envía en la cabecera Authorization de cada
    petición. Cada petición queda anotada en `self.registro` (sin contraseñas) para
    que el informe Excel muestre exactamente qué se hizo y cuándo.

    Uso típico:
        with sync_playwright() as pw:
            cliente = ClienteEDP(pw)
            cliente.login("admin@edelvives.es", "contraseña")
            usuarios = cliente.buscar_por_correos(["a@x.com", "b@x.com"])
            cliente.cambiar_username(usuarios["a@x.com"]["guid"], "nuevo@x.com")

    Atributos:
        url_base:       dirección del backoffice.
        api:            url_base + "/api/".
        pausa:          segundos de espera tras cada escritura, para no saturar el servidor.
        registro:       lista de dicts, una entrada por petición (ver _anotar).
        usuario_sesion: correo de la cuenta con la que se ha iniciado sesión.
    """

    def __init__(self, pw: Playwright, url_base: str = URL_DEFECTO, pausa: float = 0.2,
                 timeout_ms: int = 30000):
        self.pw = pw
        self.url_base = url_base.rstrip("/")
        self.api = self.url_base + "/api/"
        self.pausa = pausa
        self.timeout_ms = timeout_ms
        self.registro: list = []
        self.usuario_sesion = ""
        # Credenciales en memoria para poder renovar el token si caduca a mitad de proceso.
        self._credenciales = None
        self.ctx = pw.request.new_context(timeout=timeout_ms)

    # --- Autenticación -------------------------------------------------------

    def login(self, email: str, password: str, forzar: bool = False) -> None:
        """Inicia sesión con correo y contraseña, igual que la pantalla de login.

        Si la cuenta ya tiene una sesión abierta (por ejemplo, en el navegador), la API
        responde ACTIVE_SESSION_EXISTS. Con forzar=True se entra igualmente y la otra
        sesión se cierra.
        """
        datos = {"email": email, "password": password, "source": "backoffice",
                 "admitted_roles": ROLES_BACKOFFICE}
        if forzar:
            datos["force_login"] = True
        accion = f"Iniciar sesión como {email}" + (" (forzando)" if forzar else "")
        inicio, estado = time.monotonic(), {"http": "", "intentos": 1}
        try:
            respuesta = self.ctx.post(self.api + "login", data=datos)
        except PlaywrightError as e:
            mensaje = f"No se pudo conectar con {self.url_base}: {e.message.splitlines()[0]}"
            self._anotar(accion, "POST", "login", estado, inicio, f"ERROR: {mensaje}")
            raise ErrorLogin(mensaje) from None
        estado["http"] = respuesta.status
        cuerpo = _json(respuesta)
        token = ((cuerpo.get("data") or {}) if isinstance(cuerpo, dict) else {}).get("token")
        if respuesta.ok and token:
            self._anotar(accion, "POST", "login", estado, inicio, "OK")
            self._fijar_token(token)
            self._credenciales = (email, password, forzar)
            return
        info = extraer_error(cuerpo)
        self._anotar(accion, "POST", "login", estado, inicio,
                     f"ERROR: {info.get('errorKey') or info.get('message')}")
        if info.get("errorKey") == "ACTIVE_SESSION_EXISTS":
            raise ErrorLogin(
                "Ya hay una sesión abierta con esta cuenta (por ejemplo, en tu navegador). "
                "Vuelve a lanzarlo con --forzar-login (cerrará esa otra sesión) o usa --login-navegador.")
        raise ErrorLogin(f"Login fallido (HTTP {respuesta.status}): "
                         f"{info.get('errorKey') or info.get('message') or cuerpo}")

    def login_navegador(self, espera_s: int = 300) -> None:
        """Abre Chromium para que la persona inicie sesión a mano. Sirve también si la
        cuenta entra con SSO (Microsoft o Google) o si no se quiere guardar la contraseña.

        Tras entrar, la web hace peticiones a /api/ con la cabecera Authorization. De
        ahí se copia el token y se cierra el navegador.
        """
        capturado = {}

        def al_pedir(peticion):
            auth = peticion.headers.get("authorization") or ""
            if "/api/" in peticion.url and len(auth) >= 20 and auth not in ("undefined", "null"):
                capturado["token"] = auth

        inicio = time.monotonic()
        navegador = self.pw.chromium.launch(headless=False)
        try:
            pagina = navegador.new_page()
            pagina.on("request", al_pedir)
            pagina.goto(self.url_base + "/auth/login")
            print(f"Inicia sesión en la ventana de Chromium (tienes {espera_s // 60} minutos)...")
            limite = time.time() + espera_s
            while "token" not in capturado:
                if time.time() > limite:
                    raise ErrorLogin("No se detectó el inicio de sesión a tiempo.")
                try:
                    pagina.wait_for_timeout(500)
                except PlaywrightError:
                    raise ErrorLogin("Se cerró el navegador antes de iniciar sesión.") from None
        finally:
            navegador.close()
        self._anotar("Iniciar sesión a mano en el navegador", "-", "login", {"http": "", "intentos": 1},
                     inicio, "OK")
        self._fijar_token(capturado["token"])

    def quien_soy(self) -> str:
        """Comprueba que la sesión funciona y devuelve el correo de la cuenta."""
        datos = self._peticion("GET", "users/whoami", accion="Comprobar la sesión").get("data") or {}
        self.usuario_sesion = datos.get("username") or datos.get("email") or ""
        return self.usuario_sesion

    def _fijar_token(self, token: str) -> None:
        """Crea un contexto HTTP nuevo que envía el token en todas las peticiones."""
        self.ctx.dispose()
        self.ctx = self.pw.request.new_context(timeout=self.timeout_ms,
                                               extra_http_headers={"Authorization": token})

    def cerrar(self) -> None:
        self.ctx.dispose()

    # --- Peticiones y registro -------------------------------------------------

    def _peticion(self, metodo: str, ruta: str, *, accion: str, datos=None, query: str = "",
                  reintentos: int = 3):
        """Hace una petición a /api/<ruta> y devuelve el JSON de respuesta.

        Todas las peticiones (salvo el login) pasan por aquí, y cada una queda anotada
        en self.registro con la descripción `accion`, haya ido bien o mal.
        Ver _enviar para los reintentos y la gestión de errores.
        """
        url = self.api + ruta + (f"?{query}" if query else "")
        inicio, estado = time.monotonic(), {"http": "", "intentos": 0}
        try:
            cuerpo = self._enviar(metodo, url, ruta, datos, reintentos, estado)
        except Exception as e:
            self._anotar(accion, metodo, ruta, estado, inicio, f"ERROR: {e}")
            raise
        self._anotar(accion, metodo, ruta, estado, inicio, "OK")
        return cuerpo

    def _enviar(self, metodo: str, url: str, ruta: str, datos, reintentos: int, estado: dict):
        """Envía la petición con reintentos. Va actualizando `estado` (http, intentos).

        - Fallo de red o respuesta 429/5xx (problema temporal): espera 2, 4... segundos
          y reintenta, hasta `reintentos` veces.
        - Token caducado: vuelve a iniciar sesión con las mismas credenciales y repite.
        - 403 SESSION_REVOKED: lanza SesionRevocada (no tiene sentido seguir).
        - Cualquier otra respuesta de error: lanza ErrorAPI.
        """
        for intento in range(1, reintentos + 1):
            estado["intentos"] = intento
            try:
                respuesta = self.ctx.fetch(url, method=metodo, data=datos)
            except PlaywrightError:  # timeout o error de red
                if intento == reintentos:
                    raise
                time.sleep(2 ** intento)
                continue
            estado["http"] = respuesta.status
            if respuesta.status in ESTADOS_REINTENTABLES and intento < reintentos:
                time.sleep(2 ** intento)
                continue
            cuerpo = _json(respuesta)
            if respuesta.status == 403 and extraer_error(cuerpo).get("errorKey") == "SESSION_REVOKED":
                raise SesionRevocada(respuesta.status, cuerpo, metodo, ruta)
            if token_rechazado(respuesta.status, cuerpo) and self._credenciales and intento < reintentos:
                try:
                    self.login(*self._credenciales)
                except ErrorLogin:
                    raise SesionRevocada(respuesta.status, cuerpo, metodo, ruta) from None
                continue
            if not respuesta.ok or (isinstance(cuerpo, dict) and cuerpo.get("status") in ("error", "fail")):
                raise ErrorAPI(respuesta.status, cuerpo, metodo, ruta)
            return cuerpo
        raise AssertionError("inalcanzable")

    def _anotar(self, accion: str, metodo: str, ruta: str, estado: dict, inicio: float,
                resultado: str) -> None:
        """Añade una entrada al registro de peticiones (hoja "Registro API" del informe)."""
        self.registro.append({
            "fecha_hora": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "accion": accion,
            "metodo": metodo,
            "ruta": "/api/" + ruta,
            "http": estado["http"],
            "intentos": estado["intentos"],
            "duracion_ms": round((time.monotonic() - inicio) * 1000),
            "resultado": resultado,
        })

    # --- Usuarios ------------------------------------------------------------

    def buscar_por_correos(self, correos: Iterable[str]) -> dict:
        """Busca usuarios por correo de acceso, en lotes de 150 correos por petición.

        Devuelve {correo en minúsculas: usuario} solo con los correos que existen. Cada
        usuario es un dict con guid, username, name, lastname, role_guid, email, etc.
        La búsqueda de la API ya es exacta; aun así filtramos por coincidencia exacta
        por seguridad.
        """
        unicos = sorted({normalizar_correo(c) for c in correos} - {""})
        lotes = [unicos[i:i + MAX_POR_BUSQUEDA] for i in range(0, len(unicos), MAX_POR_BUSQUEDA)]
        encontrados = {}
        for n, lote in enumerate(lotes, start=1):
            # El navegador manda la lista como usernames[]=a&usernames[]=b (formato de axios).
            query = urlencode([("usernames[]", c) for c in lote] + [("pageSize", 156)])
            cuerpo = self._peticion("GET", "users", query=query,
                                    accion=f"Buscar {len(lote)} correos (lote {n} de {len(lotes)})")
            buscados = set(lote)
            for usuario in (cuerpo.get("data") or {}).get("users") or []:
                correo = normalizar_correo(usuario.get("username"))
                if correo in buscados:
                    encontrados[correo] = usuario
        return encontrados

    def obtener_usuario(self, guid: str) -> dict:
        """Ficha completa y actual de un usuario."""
        return self._peticion("GET", f"users/{guid}", accion="Leer la ficha del usuario")["data"]

    def licencias_activas(self, guid: str) -> list:
        """Licencias redimidas por un usuario que no han caducado (ver licencia_activa).

        Son las de la pestaña "Licencias" de su ficha. Cada una es un dict con code,
        linked_at (fecha de redención), expired_at (caducidad), role_guid, programs...
        """
        licencias = self._peticion("GET", f"users/{guid}/license-personals",
                                   accion="Leer las licencias del usuario").get("data") or []
        ahora = datetime.now(timezone.utc)
        return [licencia for licencia in licencias if licencia_activa(licencia, ahora)]

    def actualizar_usuario(self, guid: str, cambios: dict, accion: str) -> dict:
        """Guarda la ficha de un usuario con `cambios` aplicados, igual que el botón
        "Guardar" (ver datos_edicion). Devuelve la ficha tal como estaba antes del cambio."""
        actual = self.obtener_usuario(guid)
        self._guardar(guid, actual, cambios, accion)
        return actual

    def cambiar_username(self, guid: str, nuevo: str) -> bool:
        """Cambia el correo de acceso (username) de un usuario.

        Los usuarios tienen además un campo `email` de contacto, que a veces coincide
        con el username y a veces está vacío. Si coincidía, se cambia también, para que
        no se quede apuntando al correo antiguo. Si estaba vacío o era otro, no se toca.

        Devuelve True si también se ha cambiado el email de contacto.
        """
        actual = self.obtener_usuario(guid)
        cambios = {"username": nuevo}
        if normalizar_correo(actual.get("email")) == normalizar_correo(actual.get("username")):
            cambios["email"] = nuevo
        self._guardar(guid, actual, cambios, "Cambiar correo de acceso")
        return "email" in cambios

    def _guardar(self, guid: str, actual: dict, cambios: dict, accion: str) -> None:
        datos = {**datos_edicion(actual), **cambios}
        self._peticion("PUT", f"users/{guid}", datos=datos,
                       accion=f"{accion} ({describir_cambios(actual, cambios)})")
        time.sleep(self.pausa)


# --- Utilidades compartidas por los scripts -------------------------------------

def cargar_env(ruta: Path = Path(__file__).with_name(".env")) -> None:
    """Lee el fichero .env (junto a este script) y carga sus variables CLAVE=valor en el
    entorno. Las variables que ya existan en el entorno tienen prioridad."""
    if not ruta.exists():
        return
    for linea in ruta.read_text(encoding="utf-8").splitlines():
        linea = linea.strip()
        if linea and not linea.startswith("#") and "=" in linea:
            clave, valor = linea.split("=", 1)
            os.environ.setdefault(clave.strip(), valor.strip().strip("\"'"))


def agregar_args_conexion(parser: argparse.ArgumentParser) -> None:
    """Añade al script las opciones de conexión comunes (--url, --login-navegador...)."""
    grupo = parser.add_argument_group("conexión")
    grupo.add_argument("--url", default=os.environ.get("EDP_URL", URL_DEFECTO),
                       help=f"URL del backoffice (por defecto {URL_DEFECTO})")
    grupo.add_argument("--login-navegador", action="store_true",
                       help="iniciar sesión a mano en una ventana de Chromium en vez de usar .env")
    grupo.add_argument("--forzar-login", action="store_true",
                       help="entrar aunque haya otra sesión abierta con la cuenta (la cierra)")
    grupo.add_argument("--pausa", type=float, default=0.2,
                       help="segundos de espera tras cada escritura (por defecto 0.2)")


def conectar(pw: Playwright, args) -> ClienteEDP:
    """Crea el cliente e inicia sesión según las opciones. Si falla, termina con un mensaje."""
    cliente = ClienteEDP(pw, args.url, pausa=args.pausa)
    try:
        if args.login_navegador:
            cliente.login_navegador()
        else:
            email, password = os.environ.get("EDP_EMAIL"), os.environ.get("EDP_PASSWORD")
            if not email or not password:
                sys.exit("Faltan EDP_EMAIL y EDP_PASSWORD en el fichero .env (o usa --login-navegador).")
            cliente.login(email, password, forzar=args.forzar_login)
        cliente.quien_soy()
    except (ErrorLogin, ErrorAPI) as e:
        sys.exit(str(e))
    print(f"Sesión iniciada en {cliente.url_base} como {cliente.usuario_sesion}")
    return cliente


def main() -> None:
    """Comprobación de conexión: inicia sesión y muestra la ficha de los correos dados."""
    cargar_env()
    parser = argparse.ArgumentParser(
        description="Comprueba la conexión y busca usuarios por correo (solo lectura).")
    parser.add_argument("correos", nargs="+")
    agregar_args_conexion(parser)
    args = parser.parse_args()
    with sync_playwright() as pw:
        cliente = conectar(pw, args)
        encontrados = cliente.buscar_por_correos(args.correos)
        for correo in args.correos:
            usuario = encontrados.get(normalizar_correo(correo))
            if not usuario:
                print(f"\n{correo}: NO EXISTE")
                continue
            ficha = cliente.obtener_usuario(usuario["guid"])
            activas = cliente.licencias_activas(usuario["guid"])
            print(f"\n{correo}: guid={ficha.get('guid')} · {ficha.get('name')} {ficha.get('lastname')}"
                  f" · {NOMBRES_ROL.get(ficha.get('role_guid'), ficha.get('role_guid'))}"
                  f" · email de contacto: {ficha.get('email') or '(vacío)'}"
                  f" · licencias activas: {len(activas)}")
            print("  datos que se reenvían al guardar:", json.dumps(datos_edicion(ficha), ensure_ascii=False))
        cliente.cerrar()


if __name__ == "__main__":
    main()
