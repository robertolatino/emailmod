"""
reiniciar_password.py — Reinicia la contraseña de usuarios en el backoffice de Edelvives Digital Plus.

QUÉ HACE
--------
Es un módulo independiente de cambiar_correos.py. Hace lo mismo que la pestaña
"Generar nueva contraseña" de la ficha del usuario en el backoffice:
  - Pone la misma contraseña a todos los usuarios: @Edelvives por defecto, u otra
    indicada con --password.
  - Salvo que se use --sin-forzar-cambio, marca la casilla "Notificar al usuario
    para generar una nueva contraseña al iniciar sesión". El usuario entra con esa
    contraseña y la plataforma le pide crear una propia.
Al usuario no se le cambia ningún otro dato.

ENTRADA
-------
Un Excel .xlsx con una columna de correos. Se usa la primera columna que exista
de: "Correo nuevo", "Correo", "Email" o "Username" (o la indicada con --columna).

Se puede usar directamente el informe de cambiar_correos.py
(informe_cambio_correos_*.xlsx). En ese caso se lee su hoja Resultados y solo se
usan las filas en estado OK (correos cambiados). Con --todas se usan todas.

CÓMO LO HACE
------------
1. Lee el Excel y decide qué filas se procesan. Las filas sin correo, las repetidas,
   las que no están en OK (si es un informe) y las que superan --limite se anotan
   en el informe con su motivo.
2. Inicia sesión y busca todos los correos de una vez, en lotes de 150.
3. Para cada usuario encontrado, guarda su ficha con la contraseña nueva.
4. Escribe el informe Excel, incluso si la ejecución se interrumpe.

SALIDA
------
informe_passwords_<fecha>_<hora>.xlsx, junto al Excel de entrada, con las hojas
Resumen, Resultados, Registro API y Leyenda.

USO
---
  python3 reiniciar_password.py informe_cambio_correos_XXXX.xlsx                # simulación
  python3 reiniciar_password.py informe_cambio_correos_XXXX.xlsx --ejecutar     # @Edelvives
  python3 reiniciar_password.py correos.xlsx --password Edelvives2027 --ejecutar
"""
from __future__ import annotations

import argparse
import sys
from collections import Counter
from datetime import datetime
from pathlib import Path

from playwright.sync_api import sync_playwright

from edp_api import (NOMBRES_ROL, SesionRevocada, agregar_args_conexion, cargar_env, conectar,
                     normalizar_correo)
from informe_excel import Columna, formatear_duracion, guardar_informe, leer_excel, seccion_peticiones

PASSWORD_DEFECTO = "@Edelvives"
# Columnas (normalizadas) donde se buscan los correos, por orden de preferencia.
COLUMNAS_CORREO = ["correo_nuevo", "correo", "email", "username"]

ESTADOS = [
    ("OK", "Contraseña reiniciada."),
    ("SIMULADO", "Modo simulación: se reiniciaría con --ejecutar. No se ha cambiado nada."),
    ("NO ENCONTRADO", "No hay ningún usuario con ese correo de acceso."),
    ("OMITIDO", "Fila no procesada a propósito: correo repetido, fila no OK en el informe de entrada o fuera del --limite."),
    ("ERROR", "No se pudo reiniciar; la columna Detalle explica el motivo. Se puede relanzar."),
    ("NO PROCESADO", "La ejecución se interrumpió antes de llegar a esta fila."),
]

COLUMNAS = [
    Columna("fila", "Fila Excel", 8, "Número de fila en el Excel de entrada."),
    Columna("correo", "Correo", 38, "Correo de acceso del usuario."),
    Columna("estado", "Estado", 15, "Resultado de la fila (ver sección Estados)."),
    Columna("detalle", "Detalle", 70, "Qué se hizo, o por qué no."),
    Columna("guid", "GUID usuario", 38, "Identificador del usuario en la plataforma."),
    Columna("nombre", "Nombre", 16, "Nombre del usuario."),
    Columna("apellidos", "Apellidos", 20, "Apellidos del usuario."),
    Columna("rol", "Rol", 13, "Rol del usuario."),
    Columna("hora", "Hora", 10, "Hora a la que se procesó la fila."),
]


def preparar_filas(filas: list, args) -> tuple:
    """Paso 1: decide qué filas se procesan.

    Devuelve (columna de correos usada, descripción del filtro aplicado, plan). El plan
    tiene una entrada por fila del Excel: {"fila", "correo", "estado", "detalle"}. Las
    filas que se van a procesar llevan el estado vacío; las demás, ya su estado y motivo.
    """
    if not filas:
        sys.exit("El Excel no tiene filas con datos.")
    columnas = set().union(*filas)
    columna = args.columna or next((c for c in COLUMNAS_CORREO if c in columnas), None)
    if not columna or columna not in columnas:
        sys.exit(f"No encuentro la columna de correos (columnas leídas: "
                 f"{', '.join(sorted(c for c in columnas if c != '_fila'))}). Usa --columna.")
    es_informe = "estado" in columnas
    if not es_informe:
        filtro = "No aplica: el Excel no es un informe de cambiar_correos.py"
    elif args.todas:
        filtro = "Todas las filas del informe de entrada (--todas)"
    else:
        filtro = "Solo las filas en estado OK del informe de entrada"

    plan, vistos, procesables = [], {}, 0
    for fila in filas:
        correo = normalizar_correo(fila.get(columna))
        entrada = {"fila": fila["_fila"], "correo": correo, "estado": "", "detalle": ""}
        if not correo:
            entrada.update(estado="ERROR", detalle="Fila sin correo")
        elif es_informe and not args.todas and fila.get("estado") != "OK":
            entrada.update(estado="OMITIDO", detalle=f"En el informe de entrada su estado es "
                           f"{fila.get('estado') or '(vacío)'}: solo se usan las filas en OK (--todas para incluirla)")
        elif correo in vistos:
            entrada.update(estado="OMITIDO", detalle=f"Correo repetido: ya se procesa en la fila {vistos[correo]}")
        elif args.limite and procesables >= args.limite:
            entrada.update(estado="OMITIDO", detalle=f"Fuera del límite de {args.limite} usuarios (--limite)")
        else:
            vistos[correo] = fila["_fila"]
            procesables += 1
        plan.append(entrada)
    return columna, filtro, plan


def escribir_informe(args, cliente, hoja: str, columna: str, filtro: str, resultados: list,
                     fin_ejecucion: str, inicio: datetime) -> Path:
    """Paso 4: compone las hojas del informe y lo guarda junto al Excel de entrada."""
    fin = datetime.now()
    por_estado = Counter(r["estado"] for r in resultados)
    resumen = [
        ("Ejecución", [
            ("Programa", "reiniciar_password.py: reinicio de contraseñas"),
            ("Modo", "EJECUCIÓN: las contraseñas se han cambiado en la plataforma" if args.ejecutar
             else "SIMULACIÓN: no se ha cambiado nada (para aplicar, añade --ejecutar)"),
            ("Cómo terminó", fin_ejecucion),
            ("Inicio", inicio.strftime("%d/%m/%Y %H:%M:%S")),
            ("Fin", fin.strftime("%d/%m/%Y %H:%M:%S")),
            ("Duración", formatear_duracion((fin - inicio).total_seconds())),
            ("Backoffice", cliente.url_base),
            ("Usuario del backoffice", cliente.usuario_sesion),
            ("Fichero de entrada", str(Path(args.entrada).resolve())),
            ("Hoja leída", hoja),
            ("Columna de correos", columna),
            ("Filtro de filas", filtro),
            ("Filas en este informe", len(resultados)),
            ("Límite de usuarios (--limite)", args.limite or "Sin límite"),
        ]),
        ("Contraseña", [
            ("Contraseña aplicada", args.password),
            ("Obligar a cambiarla al iniciar sesión", "No (--sin-forzar-cambio)" if args.sin_forzar_cambio else "Sí"),
        ]),
        ("Totales por estado", [(estado, por_estado[estado]) for estado, _ in ESTADOS if por_estado[estado]]),
        seccion_peticiones(cliente.registro),
    ]
    ruta = Path(args.entrada).resolve().parent / f"informe_passwords_{inicio:%Y%m%d_%H%M%S}.xlsx"
    return guardar_informe(ruta, "Informe de reinicio de contraseñas", resumen, COLUMNAS, resultados,
                           cliente.registro, [("Estados", ESTADOS)])


def main() -> None:
    cargar_env()
    parser = argparse.ArgumentParser(description="Reinicia la contraseña de usuarios de Edelvives Digital Plus.")
    parser.add_argument("entrada", help="Excel .xlsx con una columna de correos (o el informe de cambiar_correos.py)")
    parser.add_argument("--password", default=PASSWORD_DEFECTO,
                        help=f"contraseña para todos (por defecto {PASSWORD_DEFECTO})")
    parser.add_argument("--sin-forzar-cambio", action="store_true",
                        help="no obligar al usuario a cambiar la contraseña al iniciar sesión")
    parser.add_argument("--columna", help="columna con los correos, normalizada: p. ej. correo_nuevo (por defecto se detecta)")
    parser.add_argument("--todas", action="store_true", help="con un informe de cambiar_correos.py, usar también las filas que no están en OK")
    parser.add_argument("--ejecutar", action="store_true", help="aplicar los cambios (sin esto solo simula)")
    parser.add_argument("--limite", type=int, help="procesar solo los N primeros usuarios (para pruebas)")
    agregar_args_conexion(parser)
    args = parser.parse_args()

    inicio = datetime.now()
    filas, hoja = leer_excel(args.entrada)
    columna, filtro, plan = preparar_filas(filas, args)
    pendientes = [p for p in plan if not p["estado"]]
    print(f"{len(pendientes)} usuarios a procesar (filtro: {filtro}) · contraseña {args.password} · "
          f"{'sin forzar' if args.sin_forzar_cambio else 'forzando'} el cambio al iniciar sesión · "
          f"modo {'EJECUCIÓN' if args.ejecutar else 'SIMULACIÓN (añade --ejecutar para aplicar)'}")

    cambios = {"password": args.password}
    if not args.sin_forzar_cambio:
        cambios["change_password"] = 1
    fin_ejecucion = "Completa"
    with sync_playwright() as pw:
        cliente = conectar(pw, args)
        try:
            encontrados = cliente.buscar_por_correos(p["correo"] for p in pendientes)  # paso 2
            for i, entrada in enumerate(pendientes, start=1):  # paso 3
                usuario = encontrados.get(entrada["correo"])
                if usuario:
                    entrada.update(guid=usuario["guid"], nombre=(usuario.get("name") or "").strip(),
                                   apellidos=(usuario.get("lastname") or "").strip(),
                                   rol=NOMBRES_ROL.get(usuario.get("role_guid"), usuario.get("role_guid") or ""))
                if not usuario:
                    entrada.update(estado="NO ENCONTRADO", detalle="No hay ningún usuario con este correo de acceso")
                elif not args.ejecutar:
                    entrada.update(estado="SIMULADO", detalle="Se reiniciaría la contraseña")
                else:
                    try:
                        cliente.actualizar_usuario(usuario["guid"], cambios, "Reiniciar contraseña")
                        entrada.update(estado="OK", detalle="Contraseña reiniciada" + (
                            "" if args.sin_forzar_cambio else "; deberá cambiarla al iniciar sesión"))
                    except SesionRevocada:
                        raise
                    except Exception as e:  # un fallo en un usuario no debe parar el resto
                        entrada.update(estado="ERROR", detalle=str(e))
                entrada["hora"] = datetime.now().strftime("%H:%M:%S")
                print(f"[{i}/{len(pendientes)}] {entrada['correo']}: {entrada['estado']} · {entrada['detalle']}")
        except SesionRevocada:
            fin_ejecucion = ("Interrumpida: se cerró la sesión porque alguien entró con la misma cuenta "
                             "en otro sitio. Las filas NO PROCESADO pueden relanzarse.")
        except KeyboardInterrupt:
            fin_ejecucion = "Interrumpida por el usuario (Ctrl+C). Las filas NO PROCESADO pueden relanzarse."
        except Exception as e:
            fin_ejecucion = f"Interrumpida por un error inesperado: {e}"
        finally:
            cliente.cerrar()
            for entrada in pendientes:
                if not entrada["estado"]:
                    # Solo la fila en curso tiene ya el guid rellenado.
                    entrada.update(estado="NO PROCESADO", detalle=(
                        "La ejecución se interrumpió mientras se procesaba este usuario; relánzalo para asegurarte"
                        if entrada.get("guid") else "La ejecución se interrumpió antes de llegar a esta fila"))
            ruta = escribir_informe(args, cliente, hoja, columna, filtro, plan, fin_ejecucion, inicio)  # paso 4
            resumen = Counter(p["estado"] for p in plan)
            print(f"\nEjecución: {fin_ejecucion}")
            print("Resumen:", " · ".join(f"{k}: {v}" for k, v in sorted(resumen.items())))
            print("Informe:", ruta)


if __name__ == "__main__":
    main()
