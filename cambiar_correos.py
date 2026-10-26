"""
cambiar_correos.py — Cambia el correo de acceso de usuarios en el backoffice de Edelvives Digital Plus.

QUÉ HACE
--------
Lee un Excel con parejas (correo antiguo, correo nuevo) y deja a cada usuario con el
correo nuevo como correo de acceso. Según lo que ya exista en la plataforma, cada
fila cae en uno de estos cuatro casos (se comprueban en este orden):

  Caso 1. El antiguo existe y el nuevo está libre
          -> Se pone el correo nuevo al usuario antiguo.
  Caso 2. El antiguo existe y el nuevo ya está registrado (la persona se registró
          por su cuenta con el correo nuevo)
          -> Se miran las licencias del usuario que tiene el correo nuevo:
             - si tiene alguna activa, no se toca a ninguno de los dos usuarios;
             - si no, pasa a <correo nuevo>_old (p. ej. ana@x.com_old) y después se
               aplica el caso 1.
  Caso 3. El antiguo no existe y el nuevo está libre
          -> No se crea el usuario: queda PENDIENTE para avisar a quien corresponda y que
             se registre o se cargue aparte (hace falta saber qué licencias cargarle).
  Caso 4. El antiguo no existe y el nuevo ya está registrado
          -> No se hace nada.

A un usuario que ya existe solo se le cambia el correo de acceso. Nombre, apellidos,
rol y el resto de datos no se tocan. Única excepción: el campo "email" de contacto,
que se actualiza también si era idéntico al correo de acceso antiguo (ver
ClienteEDP.cambiar_username en edp_api.py). El programa nunca crea ni borra usuarios.

ENTRADA
-------
Un Excel .xlsx (ver plantilla_entrada.xlsx) con las columnas:
  Correo antiguo   si está vacío, se trata como "el antiguo no existe" (caso 3 o 4)
  Correo nuevo     obligatorio
  Nombre           opcional: se copia al informe en las filas del caso 3, para saber a quién avisar
  Apellidos        opcional, igual que Nombre
  Rol              opcional, igual que Nombre

CÓMO LO HACE (Y POR QUÉ ES RÁPIDO)
----------------------------------
1. Inicia sesión en la API del backoffice (edp_api.py).
2. Busca TODOS los correos del Excel de una vez (antiguos, nuevos y los posibles
   <correo nuevo>_old): la API acepta hasta 150 correos por petición, así que
   1.000 filas se comprueban con unas 20 peticiones. Con eso se construye un
   "índice" en memoria {correo: usuario}.
3. Recorre las filas en orden. Cada fila se clasifica con el índice (sin peticiones
   extra) y, en modo ejecución, se aplica la acción de su caso. El índice se va
   actualizando con cada cambio, así que una fila ve el resultado de las anteriores.
   En el caso 2 se leen además las licencias del usuario que tiene el correo nuevo,
   también en simulación (solo es una lectura).
4. Red de seguridad: si en el caso 1 la API responde que el correo nuevo ya existe
   (alguien lo registró mientras tanto), la fila pasa automáticamente al caso 2,
   con su comprobación de licencias.
5. Escribe el informe Excel (informe_excel.py), incluso si la ejecución se interrumpe.

SALIDA
------
informe_cambio_correos_<fecha>_<hora>.xlsx, junto al Excel de entrada, con las hojas
Resumen, Resultados, Registro API y Leyenda. La hoja Resultados sirve directamente
como entrada de reiniciar_password.py.

USO
---
  python3 cambiar_correos.py usuarios.xlsx                        # simulación: no cambia nada
  python3 cambiar_correos.py usuarios.xlsx --ejecutar --limite 1  # aplica solo la primera fila
  python3 cambiar_correos.py usuarios.xlsx --ejecutar             # aplica todo

Relanzar el mismo Excel es seguro: las filas que ya se hicieron salen como caso 4.
"""
from __future__ import annotations

import argparse
import sys
from collections import Counter
from dataclasses import asdict, dataclass
from datetime import datetime
from pathlib import Path

from playwright.sync_api import sync_playwright

from edp_api import (NOMBRES_ROL, ErrorAPI, SesionRevocada, agregar_args_conexion, cargar_env, conectar,
                     normalizar_correo)
from informe_excel import Columna, formatear_duracion, guardar_informe, leer_excel, seccion_peticiones

# Nombres de columna aceptados en el Excel (ya normalizados: minúsculas, sin tildes, "_").
ALIAS_COLUMNAS = {
    "correo_antiguo": ["correo_antiguo", "email_antiguo", "antiguo", "old_email"],
    "correo_nuevo": ["correo_nuevo", "email_nuevo", "nuevo", "new_email"],
}

# Cada caso: (situación detectada, acción que corresponde).
CASOS = {
    1: ("El antiguo existe y el nuevo está libre", "Se pone el correo nuevo al usuario antiguo."),
    2: ("El antiguo existe y el nuevo ya está registrado",
        "Si el usuario que tiene el correo nuevo no tiene licencias activas, pasa a <correo nuevo>_old y se "
        "pone el correo nuevo al usuario antiguo. Si tiene alguna, no se toca a ninguno de los dos."),
    3: ("El antiguo no existe y el nuevo está libre",
        "No se crea el usuario: hay que avisar para que se registre o se cargue aparte con sus licencias."),
    4: ("El antiguo no existe y el nuevo ya está registrado", "No se hace nada."),
    "-": ("Fila no válida", "No se procesa."),
}

ESTADOS = [
    ("OK", "La acción se ha realizado correctamente."),
    ("SIMULADO", "Modo simulación: es lo que se haría con --ejecutar. No se ha cambiado nada."),
    ("OMITIDO", "No se hace nada: caso 4, caso 2 en el que el usuario que tiene el correo nuevo tiene licencias "
                "activas, o el correo antiguo y el nuevo son iguales."),
    ("PENDIENTE", "Caso 3: no existe ningún usuario con estos correos y el programa no lo crea. Hay que avisar "
                  "para que se registre o se cargue aparte con sus licencias."),
    ("ERROR", "La acción ha fallado; la columna Detalle explica el motivo. La fila se puede relanzar."),
    ("NO PROCESADO", "La ejecución se interrumpió antes de llegar a esta fila."),
]

# Columnas de la hoja Resultados del informe.
COLUMNAS = [
    Columna("fila", "Fila Excel", 8, "Número de fila en el Excel de entrada."),
    Columna("correo_antiguo", "Correo antiguo", 34, "Tal como venía en el Excel de entrada."),
    Columna("correo_nuevo", "Correo nuevo", 34, "Tal como venía en el Excel de entrada (se aplica en minúsculas)."),
    Columna("caso", "Caso", 7, "1, 2, 3 o 4 (ver sección Casos). '-' si la fila no es válida."),
    Columna("situacion", "Situación detectada", 34, "Qué se encontró en la plataforma para esta fila."),
    Columna("estado", "Estado", 14, "Resultado de la fila (ver sección Estados)."),
    Columna("detalle", "Detalle", 70, "Qué se hizo exactamente, o el motivo del error."),
    Columna("guid", "GUID usuario", 38, "Identificador del usuario afectado: el del correo antiguo (casos 1 y 2) o, en el caso 4, el que ya tiene el correo nuevo. Vacío en el caso 3."),
    Columna("nombre", "Nombre", 16, "Nombre de ese usuario (en el caso 3, el del Excel)."),
    Columna("apellidos", "Apellidos", 20, "Apellidos de ese usuario (en el caso 3, los del Excel)."),
    Columna("rol", "Rol", 13, "Rol de ese usuario (en el caso 3, el del Excel)."),
    Columna("email_contacto", "Email de contacto también cambiado", 14, "'Sí' si el campo email de contacto era igual al correo antiguo y se ha cambiado también."),
    Columna("licencias_nuevo", "Licencias activas del usuario con el correo nuevo (caso 2)", 14, "Caso 2: licencias sin caducar del usuario que ya tenía el correo nuevo. Si es 0 se renombra a _old; si es más de 0 no se toca nada."),
    Columna("correo_renombrado", "Correo renombrado (caso 2)", 38, "Correo que recibe el usuario que ya tenía el correo nuevo (correo nuevo + _old). Vacío si no se renombra."),
    Columna("guid_correo_nuevo", "GUID usuario con el correo nuevo (caso 2)", 38, "Identificador del usuario que ya tenía el correo nuevo (renombrado o, si tiene licencias activas, sin tocar)."),
    Columna("hora", "Hora", 10, "Hora a la que se procesó la fila."),
]


@dataclass
class ResultadoFila:
    """Resultado de procesar una fila. Cada campo es una columna del informe."""
    caso: object = "-"
    estado: str = ""
    detalle: str = ""
    guid: str = ""
    nombre: str = ""
    apellidos: str = ""
    rol: str = ""
    email_contacto: str = ""
    licencias_nuevo: object = ""
    correo_renombrado: str = ""
    guid_correo_nuevo: str = ""
    hora: str = ""

    def con_usuario(self, usuario: dict) -> None:
        """Copia guid, nombre, apellidos y rol de un usuario de la API."""
        self.guid = usuario.get("guid", "")
        self.nombre = (usuario.get("name") or "").strip()
        self.apellidos = (usuario.get("lastname") or "").strip()
        self.rol = NOMBRES_ROL.get(usuario.get("role_guid"), usuario.get("role_guid") or "")


class ErrorFila(Exception):
    """Error propio del proceso de una fila (no viene de la API)."""


def si_no(valor: bool) -> str:
    return "Sí" if valor else "No"


# --- Lectura y validación de la entrada ----------------------------------------------

def resolver_columnas(filas: list) -> None:
    """Comprueba que el Excel tiene las columnas de correo y renombra los alias
    (p. ej. "email_antiguo") a los nombres que usa el programa. Termina si faltan."""
    if not filas:
        sys.exit("El Excel no tiene filas con datos.")
    columnas = set().union(*filas)
    for canonica, alias in ALIAS_COLUMNAS.items():
        encontrada = next((a for a in alias if a in columnas), None)
        if not encontrada:
            sys.exit(f"Falta la columna {canonica} en el Excel (columnas leídas: "
                     f"{', '.join(sorted(c for c in columnas if c != '_fila'))})")
        for fila in filas:
            fila[canonica] = fila.pop(encontrada, "")


def avisos_duplicados(filas: list) -> list:
    """Avisos de correos repetidos en el Excel. No es un error: las filas se procesan en
    orden y cada una ve el resultado de la anterior, pero conviene revisarlo."""
    avisos = []
    for columna, nombre in (("correo_antiguo", "Correo antiguo"), ("correo_nuevo", "Correo nuevo")):
        cuenta = Counter(normalizar_correo(f[columna]) for f in filas)
        for correo, veces in cuenta.items():
            if correo and veces > 1:
                numeros = [str(f["_fila"]) for f in filas if normalizar_correo(f[columna]) == correo]
                avisos.append(f"{correo} aparece {veces} veces en la columna {nombre} (filas {', '.join(numeros)})")
    return avisos


# --- Lógica de los casos ----------------------------------------------------------

def clasificar(existe_antiguo: bool, existe_nuevo: bool) -> int:
    """Caso que corresponde según exista o no cada correo en la plataforma."""
    if existe_antiguo:
        return 2 if existe_nuevo else 1
    return 4 if existe_nuevo else 3


def candidatos_renombre(correo: str, sufijo: str):
    """Correos posibles para el usuario que ya tiene el correo nuevo:
    correo_old, correo_old2, ..., correo_old10 (por si quedan de ejecuciones anteriores)."""
    yield f"{correo}{sufijo}"
    for n in range(2, 11):
        yield f"{correo}{sufijo}{n}"


def cambiar_correo(cliente, indice: dict, usuario: dict, correo_actual: str, correo_nuevo: str) -> bool:
    """Cambia el correo de acceso en la plataforma y en el índice. Devuelve True si
    también se cambió el email de contacto."""
    email_cambiado = cliente.cambiar_username(usuario["guid"], correo_nuevo)
    indice.pop(correo_actual, None)
    usuario["username"] = correo_nuevo
    indice[correo_nuevo] = usuario
    return email_cambiado


def liberar_correo(cliente, indice: dict, usuario: dict, correo: str, sufijo: str) -> tuple:
    """Caso 2: renombra al usuario que ocupa `correo` para dejarlo libre.

    Prueba correo+sufijo, correo+sufijo2... Salta los que ya están en el índice y,
    si la API dice que uno ya existe, pasa al siguiente.
    Devuelve (correo asignado, True si también se cambió su email de contacto).
    """
    for candidato in candidatos_renombre(correo, sufijo):
        if candidato in indice:
            continue
        try:
            return candidato, cambiar_correo(cliente, indice, usuario, correo, candidato)
        except ErrorAPI as e:
            if not e.usuario_ya_existe:
                raise
    raise ErrorFila(f"No se encontró un correo libre para renombrar {correo} con el sufijo {sufijo}")


def nuevo_con_licencias(cliente, res: ResultadoFila) -> bool:
    """Caso 2: mira si el usuario que ya tiene el correo nuevo (res.guid_correo_nuevo) tiene
    licencias activas. Si tiene alguna, no se toca a ninguno de los dos usuarios: la fila
    queda OMITIDO y devuelve True."""
    res.licencias_nuevo = len(cliente.licencias_activas(res.guid_correo_nuevo))
    if not res.licencias_nuevo:
        return False
    res.estado = "OMITIDO"
    res.detalle = (f"El usuario que ya tiene el correo nuevo tiene {res.licencias_nuevo} licencia(s) activa(s): "
                   "no se cambia nada en ninguno de los dos usuarios")
    return True


def procesar_fila(cliente, indice: dict, fila: dict, args) -> ResultadoFila:
    """Clasifica una fila y, en modo ejecución, aplica la acción de su caso."""
    antiguo = normalizar_correo(fila.get("correo_antiguo"))
    nuevo = normalizar_correo(fila.get("correo_nuevo"))
    if not nuevo:
        return ResultadoFila(estado="ERROR", detalle="Fila incompleta: falta el correo nuevo")
    if antiguo == nuevo:
        return ResultadoFila(estado="OMITIDO", detalle="El correo antiguo y el nuevo son iguales")

    # Sin correo antiguo (p. ej. un docente que se incorpora) = "el antiguo no existe": caso 3 o 4.
    usuario_antiguo = indice.get(antiguo) if antiguo else None
    usuario_nuevo = indice.get(nuevo)
    res = ResultadoFila(caso=clasificar(usuario_antiguo is not None, usuario_nuevo is not None))
    nota_sin_antiguo = " (la fila no tiene correo antiguo)" if not antiguo else ""

    if res.caso == 4:
        res.con_usuario(usuario_nuevo)
        res.estado = "OMITIDO"
        res.detalle = ("No se hace nada: el correo nuevo ya pertenece a un usuario y el antiguo no existe"
                       + nota_sin_antiguo)
        return res
    if res.caso == 3:
        # No se crea: hay que saber qué licencias cargarle, así que se avisa para que se haga aparte.
        # Nombre, Apellidos y Rol del Excel (si vienen) van al informe para saber a quién avisar.
        res.nombre, res.apellidos, res.rol = (fila.get(c, "").strip() for c in ("nombre", "apellidos", "rol"))
        res.estado = "PENDIENTE"
        res.detalle = ("No existe ningún usuario con estos correos y no se crea: hay que avisar para que se "
                       "registre o se cargue aparte con sus licencias" + nota_sin_antiguo)
        return res

    res.con_usuario(usuario_antiguo)  # casos 1 y 2
    try:
        if res.caso == 2:
            res.guid_correo_nuevo = usuario_nuevo["guid"]
            if nuevo_con_licencias(cliente, res):
                return res
            # Previsión del correo que recibirá (en ejecución se confirma con la API).
            res.correo_renombrado = next((c for c in candidatos_renombre(nuevo, args.sufijo) if c not in indice), "")
        if not args.ejecutar:
            res.estado = "SIMULADO"
            res.detalle = "Se haría: " + (CASOS[1][1] if res.caso == 1 else
                                          f"el usuario que tiene {nuevo} (sin licencias activas) pasa a "
                                          f"{res.correo_renombrado} y se pone el correo nuevo al usuario antiguo.")
            return res
        ejecutar_caso(cliente, indice, res, antiguo, nuevo, args.sufijo)
    except SesionRevocada:
        raise  # no tiene sentido seguir: lo gestiona main()
    except Exception as e:  # un fallo en una fila no debe parar el resto
        res.estado, res.detalle = "ERROR", str(e)
    return res


def ejecutar_caso(cliente, indice: dict, res: ResultadoFila, antiguo: str, nuevo: str, sufijo: str) -> None:
    """Aplica en la plataforma la acción del caso de la fila (1 o 2) y rellena `res`."""
    usuario_antiguo = indice[antiguo]
    nota = ""
    if res.caso == 1:
        try:
            res.email_contacto = si_no(cambiar_correo(cliente, indice, usuario_antiguo, antiguo, nuevo))
            res.estado, res.detalle = "OK", "Correo de acceso cambiado"
            return
        except ErrorAPI as e:
            if not e.usuario_ya_existe:
                raise
            # El correo nuevo se registró después de la búsqueda inicial: pasamos al caso 2.
            usuario_nuevo = cliente.buscar_por_correos([nuevo]).get(nuevo)
            if not usuario_nuevo:
                raise
            indice[nuevo] = usuario_nuevo
            res.caso, res.guid_correo_nuevo = 2, usuario_nuevo["guid"]
            nota = " (el correo nuevo se registró mientras se ejecutaba el programa; se detectó al intentar el cambio)"
            if nuevo_con_licencias(cliente, res):
                res.detalle += nota
                return

    # Caso 2 (el usuario del correo nuevo no tiene licencias activas): primero se deja libre
    # el correo nuevo y después se aplica el caso 1.
    res.correo_renombrado, email_renombrado = liberar_correo(cliente, indice, indice[nuevo], nuevo, sufijo)
    detalle_renombre = f"El usuario que ya tenía {nuevo} (sin licencias activas) pasa a {res.correo_renombrado}" + (
        " (también su email de contacto)" if email_renombrado else "")
    try:
        res.email_contacto = si_no(cambiar_correo(cliente, indice, usuario_antiguo, antiguo, nuevo))
    except ErrorAPI as e:
        res.estado = "ERROR"
        res.detalle = (f"{detalle_renombre}, pero falló el cambio del correo de acceso ({e}). "
                       "Al relanzar la fila será un caso 1.")
        return
    res.estado = "OK"
    res.detalle = f"{detalle_renombre}; después se cambió el correo de acceso{nota}"


# --- Programa principal ---------------------------------------------------------

def buscar_todos(cliente, filas: list, sufijo: str) -> dict:
    """Paso 2: busca de una vez todos los correos del Excel y devuelve el índice.

    También busca los <correo nuevo>_old, para que el caso 2 sepa qué sufijos están ya
    usados sin gastar intentos.
    """
    antiguos = {normalizar_correo(f["correo_antiguo"]) for f in filas} - {""}
    nuevos = {normalizar_correo(f["correo_nuevo"]) for f in filas} - {""}
    indice = cliente.buscar_por_correos(antiguos | nuevos | {c + sufijo for c in nuevos})
    print(f"Ya existen {len(antiguos & set(indice))} de {len(antiguos)} correos antiguos "
          f"y {len(nuevos & set(indice))} de {len(nuevos)} correos nuevos")
    return indice


def fila_informe(fila: dict, res: ResultadoFila) -> dict:
    """Une los datos de la fila de entrada con su resultado (una fila de la hoja Resultados)."""
    return {"fila": fila["_fila"], "correo_antiguo": fila["correo_antiguo"], "correo_nuevo": fila["correo_nuevo"],
            **asdict(res), "situacion": CASOS[res.caso][0]}


def escribir_informe(args, cliente, hoja: str, leidas: int, resultados: list, avisos: list,
                     fin_ejecucion: str, inicio: datetime) -> Path:
    """Paso 5: compone las hojas del informe y lo guarda junto al Excel de entrada."""
    fin = datetime.now()
    por_estado = Counter(r["estado"] for r in resultados)
    por_caso = Counter(r["caso"] for r in resultados)
    resumen = [
        ("Ejecución", [
            ("Programa", "cambiar_correos.py: cambio del correo de acceso"),
            ("Modo", "EJECUCIÓN: los cambios se han aplicado en la plataforma" if args.ejecutar
             else "SIMULACIÓN: no se ha cambiado nada (para aplicar, añade --ejecutar)"),
            ("Cómo terminó", fin_ejecucion),
            ("Inicio", inicio.strftime("%d/%m/%Y %H:%M:%S")),
            ("Fin", fin.strftime("%d/%m/%Y %H:%M:%S")),
            ("Duración", formatear_duracion((fin - inicio).total_seconds())),
            ("Backoffice", cliente.url_base),
            ("Usuario del backoffice", cliente.usuario_sesion),
            ("Fichero de entrada", str(Path(args.entrada).resolve())),
            ("Hoja leída", hoja),
            ("Filas con datos en la entrada", leidas),
            ("Filas en este informe", len(resultados)),
            ("Límite de filas (--limite)", args.limite or "Sin límite"),
            ("Sufijo para el usuario que ya tenía el correo nuevo (caso 2)", args.sufijo),
        ]),
        ("Totales por estado", [(estado, por_estado[estado]) for estado, _ in ESTADOS if por_estado[estado]]),
        ("Totales por caso", [(f"Caso {caso}: {CASOS[caso][0]}", por_caso[caso])
                              for caso in CASOS if por_caso[caso]]),
        seccion_peticiones(cliente.registro),
        ("Avisos", [(f"Aviso {i}", texto) for i, texto in enumerate(avisos, start=1)] or [("Avisos", "Ninguno")]),
    ]
    leyenda = [
        ("Estados", ESTADOS),
        ("Casos", [(f"Caso {caso}", f"{situacion}. {accion}") for caso, (situacion, accion) in CASOS.items()]),
    ]
    ruta = Path(args.entrada).resolve().parent / f"informe_cambio_correos_{inicio:%Y%m%d_%H%M%S}.xlsx"
    return guardar_informe(ruta, "Informe de cambio de correos de acceso", resumen, COLUMNAS,
                           resultados, cliente.registro, leyenda)


def main() -> None:
    cargar_env()
    parser = argparse.ArgumentParser(description="Cambia el correo de acceso de usuarios de Edelvives Digital Plus.")
    parser.add_argument("entrada", help="Excel .xlsx con las columnas Correo antiguo y Correo nuevo")
    parser.add_argument("--ejecutar", action="store_true", help="aplicar los cambios (sin esto solo simula)")
    parser.add_argument("--sufijo", default="_old", help="sufijo para el usuario que ya tenía el correo nuevo (por defecto _old)")
    parser.add_argument("--limite", type=int, help="procesar solo las N primeras filas (para pruebas)")
    agregar_args_conexion(parser)
    args = parser.parse_args()

    inicio = datetime.now()
    filas, hoja = leer_excel(args.entrada)
    resolver_columnas(filas)
    leidas = len(filas)
    if args.limite:
        filas = filas[:args.limite]
    avisos = avisos_duplicados(filas)
    for aviso in avisos:
        print("AVISO:", aviso)
    print(f"{len(filas)} filas · modo {'EJECUCIÓN' if args.ejecutar else 'SIMULACIÓN (añade --ejecutar para aplicar)'}")

    resultados = []
    fin_ejecucion = "Completa"
    en_curso = None  # fila que se está procesando (para el informe, si se interrumpe a mitad)
    with sync_playwright() as pw:
        cliente = conectar(pw, args)  # paso 1
        try:
            indice = buscar_todos(cliente, filas, args.sufijo)  # paso 2
            for i, fila in enumerate(filas, start=1):  # paso 3
                en_curso = fila
                res = procesar_fila(cliente, indice, fila, args)
                res.hora = datetime.now().strftime("%H:%M:%S")
                resultados.append(fila_informe(fila, res))
                en_curso = None
                print(f"[{i}/{len(filas)}] {fila['correo_antiguo']} -> {fila['correo_nuevo']}: "
                      f"caso {res.caso} · {res.estado} · {res.detalle}")
        except SesionRevocada:
            fin_ejecucion = ("Interrumpida: se cerró la sesión porque alguien entró con la misma cuenta "
                             "en otro sitio. Las filas NO PROCESADO pueden relanzarse.")
        except KeyboardInterrupt:
            fin_ejecucion = "Interrumpida por el usuario (Ctrl+C). Las filas NO PROCESADO pueden relanzarse."
        except Exception as e:
            fin_ejecucion = f"Interrumpida por un error inesperado: {e}"
        finally:
            cliente.cerrar()
            for fila in filas[len(resultados):]:
                detalle = ("La ejecución se interrumpió mientras se procesaba esta fila; en el caso 2 puede haberse "
                           "hecho ya el renombrado _old. Al relanzar se completa sola" if fila is en_curso
                           else "La ejecución se interrumpió antes de llegar a esta fila")
                resultados.append(fila_informe(fila, ResultadoFila(estado="NO PROCESADO", detalle=detalle)))
            ruta = escribir_informe(args, cliente, hoja, leidas, resultados, avisos, fin_ejecucion, inicio)  # paso 5
            resumen = Counter(f"caso {r['caso']} {r['estado']}" for r in resultados)
            print(f"\nEjecución: {fin_ejecucion}")
            print("Resumen:", " · ".join(f"{k}: {v}" for k, v in sorted(resumen.items())))
            print("Informe:", ruta)


if __name__ == "__main__":
    main()
