# Cambio de correos de acceso y reinicio de contraseñas — Edelvives Digital Plus

Programas en Python + Playwright para hacer en bloque, en el backoffice de
Edelvives Digital Plus (https://publisher.edelvivesdigitalplus.com), dos tareas
que a mano se hacen usuario por usuario:

1. **Cambiar el correo de acceso** de una lista de usuarios (`cambiar_correos.py`).
   Resuelve solo los cuatro casos posibles: el correo nuevo libre u ocupado, y el
   usuario antiguo existente o no.
2. **Reiniciar la contraseña** de una lista de usuarios (`reiniciar_password.py`).
   Pone `@Edelvives` u otra contraseña común y obliga a cambiarla al entrar.

Los dos leen un **Excel (.xlsx)** y generan un **informe Excel** con todo lo
ocurrido en la ejecución.

> Este documento explica **cómo usar** los programas. Cómo funcionan por dentro
> (la API del backoffice, el algoritmo, la gestión de errores) está en
> [DOCUMENTACION_TECNICA.md](DOCUMENTACION_TECNICA.md).

---

## Índice

1. [Contenido de la carpeta](#1-contenido-de-la-carpeta)
2. [Requisitos e instalación](#2-requisitos-e-instalación)
3. [Configurar las credenciales](#3-configurar-las-credenciales)
4. [Flujo de trabajo recomendado](#4-flujo-de-trabajo-recomendado)
5. [El Excel de entrada](#5-el-excel-de-entrada)
6. [Cambio de correos: los cuatro casos](#6-cambio-de-correos-los-cuatro-casos)
7. [Reinicio de contraseñas](#7-reinicio-de-contraseñas)
8. [El informe Excel](#8-el-informe-excel)
9. [Opciones de línea de comandos](#9-opciones-de-línea-de-comandos)
10. [Seguridad y buenas prácticas](#10-seguridad-y-buenas-prácticas)
11. [Problemas frecuentes](#11-problemas-frecuentes)

---

## 1. Contenido de la carpeta

| Fichero | Para qué sirve |
|---|---|
| `cambiar_correos.py` | Programa 1: cambia el correo de acceso de los usuarios del Excel. |
| `reiniciar_password.py` | Programa 2: reinicia la contraseña de los usuarios del Excel. |
| `edp_api.py` | Módulo común: conexión con el backoffice (login, búsqueda, licencias y edición de usuarios). Ejecutado solo, comprueba la conexión. |
| `informe_excel.py` | Módulo común: lectura del Excel de entrada, informe Excel y plantilla. |
| `plantilla_entrada.xlsx` | Plantilla vacía del Excel de entrada, con instrucciones. |
| `.env.example` | Ejemplo del fichero de credenciales. |
| `.env` | Tus credenciales (lo creas tú; **no lo compartas**). |
| `README.md` | Este documento: guía de uso. |
| `DOCUMENTACION_TECNICA.md` | Cómo funciona el programa por dentro. |

## 2. Requisitos e instalación

- **Python 3.9 o superior** (en este Mac: `python3`).
- Las librerías **playwright** y **openpyxl**. Ya están instaladas en este equipo.
  En otro equipo:

  ```bash
  pip3 install playwright openpyxl
  python3 -m playwright install chromium   # solo hace falta para --login-navegador
  ```

- Una cuenta del backoffice con permisos para editar usuarios.

Todos los comandos se lanzan desde la carpeta del programa:

```bash
cd ~/Documents/Scripts/edelvives_usuarios
```

## 3. Configurar las credenciales

Copia `.env.example` como `.env` y escribe tu usuario y contraseña del backoffice:

```
EDP_EMAIL=tu.usuario@edelvives.es
EDP_PASSWORD=tu_contraseña
```

Hay dos alternativas al fichero `.env`:

- **Login manual en el navegador** (`--login-navegador`). Úsalo si entras con
  Microsoft o Google (SSO) o si no quieres guardar la contraseña en un fichero.
  Se abre una ventana de Chromium, inicias sesión normalmente y el programa
  continúa solo.
- **Una sesión por cuenta.** El backoffice solo permite una sesión abierta por
  cuenta. Si tienes el backoffice abierto en el navegador y el programa avisa de
  que "ya hay una sesión abierta", relánzalo con `--forzar-login`. Ojo: eso cierra
  la sesión del navegador. Al revés también pasa: si entras en el navegador
  forzando mientras el programa trabaja, el programa se detiene y guarda el
  informe hasta ese punto.

Para comprobar que la conexión funciona (solo lectura, no cambia nada):

```bash
python3 edp_api.py un.correo@colegio.es
```

## 4. Flujo de trabajo recomendado

1. **Rellena el Excel de entrada.** Copia `plantilla_entrada.xlsx` y rellena la
   hoja *Usuarios* (ver [sección 5](#5-el-excel-de-entrada)).

2. **Simula** (no cambia nada en la plataforma):

   ```bash
   python3 cambiar_correos.py usuarios.xlsx
   ```

   Abre el informe que se genera junto a tu Excel
   (`informe_cambio_correos_<fecha>_<hora>.xlsx`) y revisa la hoja *Resultados*.
   En la columna *Caso* verás qué se haría con cada fila, y en *Estado* las filas
   PENDIENTE (no existe ningún usuario con esos correos: hay que avisar para que se
   registre) o ERROR (fila incompleta).

3. **Prueba con una sola fila** (aplica de verdad solo la primera):

   ```bash
   python3 cambiar_correos.py usuarios.xlsx --ejecutar --limite 1
   ```

4. **Ejecuta todo:**

   ```bash
   python3 cambiar_correos.py usuarios.xlsx --ejecutar
   ```

   Si algo se interrumpe o alguna fila da ERROR, puedes **relanzar el mismo
   Excel** sin miedo: las filas ya hechas salen como caso 4 ("no se hace nada").

5. **Reinicia las contraseñas** usando como entrada el informe del paso 4. Solo se
   usan las filas que salieron OK:

   ```bash
   python3 reiniciar_password.py informe_cambio_correos_<fecha>_<hora>.xlsx             # simulación
   python3 reiniciar_password.py informe_cambio_correos_<fecha>_<hora>.xlsx --ejecutar  # @Edelvives
   ```

6. **Revisa el informe** `informe_passwords_<fecha>_<hora>.xlsx`.

Un cambio de correo tarda alrededor de **medio segundo por usuario**: unos
8–9 minutos para 1.000 usuarios. El reinicio de contraseña, lo mismo.

## 5. El Excel de entrada

Usa `plantilla_entrada.xlsx`. Tiene una hoja **Usuarios** para los datos y una
hoja **Instrucciones**.

| Columna | ¿Obligatoria? | Contenido |
|---|---|---|
| **Correo antiguo** | No | Correo de acceso actual del usuario. Si no tiene (alguien que se incorpora), déjalo vacío: se trata como "el antiguo no existe" (caso 3, o caso 4 si el correo nuevo ya existe). |
| **Correo nuevo** | Sí | Correo de acceso que debe tener. |
| **Nombre** | No | Solo informativo: se copia al informe en las filas del caso 3, para saber a quién hay que avisar. |
| **Apellidos** | No | Igual que Nombre. |
| **Rol** | No | Igual que Nombre: `alumno`, `profesor` o `admin colegio` (desplegable). |

Reglas:

- **Solo se admite `.xlsx`.** Si tienes un CSV, ábrelo en Excel y guárdalo como `.xlsx`.
- Se lee la hoja llamada *Usuarios*. Si no existe, la primera hoja.
- La primera fila con contenido es la cabecera. Los títulos pueden ir en
  mayúsculas o minúsculas y con o sin tildes. También se aceptan `email_antiguo`
  y `email_nuevo`.
- Las filas vacías se ignoran.
- Mayúsculas y espacios en los correos no importan: `  Ana@Colegio.ES ` es lo
  mismo que `ana@colegio.es`. El correo nuevo se guarda en minúsculas.
- **Nombre, Apellidos y Rol son solo informativos**: se copian al informe en las
  filas del caso 3. El programa nunca crea usuarios.
- Si un mismo correo aparece en varias filas, el programa avisa (en pantalla y en
  el informe). Las filas se procesan en orden y cada una ve el resultado de las
  anteriores.

## 6. Cambio de correos: los cuatro casos

Para cada fila, el programa mira si el **correo antiguo** y el **correo nuevo**
existen ya en la plataforma, y actúa así:

| Caso | Correo antiguo | Correo nuevo | Qué hace |
|---|---|---|---|
| **1** | existe | libre | Pone el correo nuevo al usuario antiguo. |
| **2** | existe | ya registrado | Mira las licencias del usuario que tiene el correo nuevo. **Sin licencias activas:** ese usuario pasa a `<correo nuevo>_old` y después se pone el correo nuevo al usuario antiguo. **Con alguna licencia activa:** nada, no se toca a ninguno de los dos. |
| **3** | no existe | libre | Nada: **no se crea el usuario**. La fila queda PENDIENTE para avisar a quien corresponda. |
| **4** | no existe | ya registrado | Nada. |

**Ejemplos**, con la fila `ana@antiguo.es → ana@nuevo.es`:

- **Caso 1.** Existe una usuaria con `ana@antiguo.es` y nadie usa `ana@nuevo.es`.
  Su correo de acceso pasa a ser `ana@nuevo.es`.
- **Caso 2.** Existe `ana@antiguo.es`, pero Ana también se registró por su cuenta
  con `ana@nuevo.es`. Se miran las licencias de esa cuenta `ana@nuevo.es` (las de la
  pestaña *Licencias* de su ficha que no han caducado):
  - **Ninguna activa:** esa cuenta pasa a `ana@nuevo.es_old` y deja el correo
    libre. Después la cuenta original (con sus licencias, grupos, etc.) recibe
    `ana@nuevo.es`. Si `ana@nuevo.es_old` ya existe (de una ejecución anterior), se
    usa `ana@nuevo.es_old2`, luego `_old3`… hasta `_old10`.
  - **Alguna activa:** no se toca nada, ni la cuenta antigua ni la nueva. La fila
    sale OMITIDO y el informe indica cuántas licencias activas tiene.
  - La comprobación se hace también en simulación (solo es una lectura), así que la
    simulación ya muestra qué filas del caso 2 se van a quedar sin tocar.
- **Caso 3.** No existe ninguna de las dos. **No se crea ninguna cuenta**, porque
  hay que saber qué licencias cargarle. La fila queda PENDIENTE: avisa a quien
  corresponda para que se registre o se cargue aparte. Si el Excel trae Nombre,
  Apellidos y Rol, aparecen en el informe para saber de quién se trata.
- **Caso 4.** `ana@antiguo.es` no existe y `ana@nuevo.es` ya existe. Normalmente
  significa que el cambio ya se hizo. No se toca nada.

**Qué se cambia y qué no.** A un usuario existente solo se le cambia el **correo de
acceso**. Nombre, apellidos, rol, colegios, grupos, licencias y contraseña quedan
igual. Esto se ha comprobado comparando la ficha completa antes y después.

**Única excepción: el email de contacto.** Además del correo de acceso, cada
usuario tiene un campo *email* de contacto. A veces coincide con el de acceso y a
veces está vacío. Si coincidía con el correo antiguo, se cambia también, para que
no siga apuntando al antiguo. El informe lo indica en la columna *Email de
contacto también cambiado*.

**Seguridad ante cambios simultáneos.** Si mientras el programa trabaja alguien se
registra con un correo nuevo de la lista, el caso 1 fallaría con "el usuario ya
existe". El programa lo detecta y pasa esa fila al caso 2 automáticamente, con su
comprobación de licencias.

## 7. Reinicio de contraseñas

Hace lo mismo que la pestaña *Generar nueva contraseña* de la ficha del usuario:

- Pone la **misma contraseña a todos**: `@Edelvives` por defecto, u otra con
  `--password`.
- Marca *"Notificar al usuario para generar una nueva contraseña al iniciar
  sesión"*. El usuario entra con esa contraseña y la plataforma le pide crear una
  propia. Con `--sin-forzar-cambio` no se marca.
- No cambia ningún otro dato del usuario.

```bash
python3 reiniciar_password.py informe_cambio_correos_XXXX.xlsx --ejecutar                         # @Edelvives
python3 reiniciar_password.py informe_cambio_correos_XXXX.xlsx --password Edelvives2027 --ejecutar
```

Si la contraseña lleva caracteres especiales para la terminal (`!`, `$`, `&`,
espacios…), escríbela entre comillas simples: `--password 'Edel$vives!'`.

**Entrada.** Sirven dos tipos de Excel:

- **El informe de `cambiar_correos.py`** (lo normal). Se lee su hoja *Resultados*
  y solo se usan las filas en estado **OK**: correos cambiados.
  Las demás aparecen en el informe como OMITIDO, con el motivo. Con `--todas` se
  usan todas.
- **Cualquier Excel con una columna de correos.** Se usa la primera columna que
  exista de: *Correo nuevo*, *Correo*, *Email* o *Username*. Con
  `--columna correo_de_acceso` se elige otra (el nombre va en minúsculas, sin
  tildes y con `_` en lugar de espacios).

Los correos repetidos se procesan una sola vez.

## 8. El informe Excel

Cada ejecución, **también en simulación** y también si se interrumpe, genera un
Excel junto al Excel de entrada:

- `informe_cambio_correos_<fecha>_<hora>.xlsx` (programa 1)
- `informe_passwords_<fecha>_<hora>.xlsx` (programa 2)

Hojas:

| Hoja | Contenido |
|---|---|
| **Resumen** | Programa, modo (simulación o ejecución), cómo terminó, inicio, fin, duración, usuario del backoffice, fichero de entrada, opciones usadas, totales por estado y por caso, número de peticiones a la API, errores, reintentos y avisos (p. ej. correos repetidos). En el de contraseñas, también la contraseña aplicada y si se obliga a cambiarla. |
| **Resultados** | Una fila por cada fila del Excel de entrada, con su resultado, coloreada según el estado. Tiene filtros en la cabecera. |
| **Registro API** | Cada petición hecha a la plataforma, en orden: acción en lenguaje llano (con el valor anterior y el nuevo en los cambios), ruta, código HTTP, intentos, duración y resultado. **Nunca incluye contraseñas.** |
| **Leyenda** | Qué significa cada estado, cada caso y cada columna. |

**Estados** (colores de la hoja Resultados):

| Estado | Color | Significado |
|---|---|---|
| OK | verde | Hecho correctamente. |
| SIMULADO | azul | Modo simulación: es lo que se haría con `--ejecutar`. |
| OMITIDO | gris | No se hace nada (caso 4, caso 2 con licencias activas en el usuario del correo nuevo, correos iguales, fila filtrada…). |
| PENDIENTE | amarillo | Caso 3: no existe ningún usuario con esos correos y no se crea. Hay que avisar para que se registre o se cargue aparte. |
| NO ENCONTRADO | amarillo | (Contraseñas) No existe ningún usuario con ese correo. |
| ERROR | rojo | Ha fallado; la columna *Detalle* dice por qué. Se puede relanzar. |
| NO PROCESADO | naranja | La ejecución se interrumpió antes de llegar a esa fila. |

La hoja *Resultados* del informe de `cambiar_correos.py` es la entrada de
`reiniciar_password.py` (paso 5 del flujo).

## 9. Opciones de línea de comandos

`python3 cambiar_correos.py <excel.xlsx> [opciones]`

| Opción | Efecto |
|---|---|
| *(ninguna)* | Simulación: no cambia nada y genera el informe. |
| `--ejecutar` | Aplica los cambios. |
| `--limite N` | Procesa solo las N primeras filas (para pruebas). |
| `--sufijo TEXTO` | Sufijo del caso 2 (por defecto `_old`). |

`python3 reiniciar_password.py <excel.xlsx> [opciones]`

| Opción | Efecto |
|---|---|
| *(ninguna)* | Simulación con la contraseña por defecto. |
| `--ejecutar` | Aplica los cambios. |
| `--password TEXTO` | Contraseña para todos (por defecto `@Edelvives`). |
| `--sin-forzar-cambio` | No obliga a cambiar la contraseña al iniciar sesión. |
| `--todas` | Con un informe de entrada, usa también las filas que no están en OK. |
| `--columna NOMBRE` | Columna de correos a usar (por defecto se detecta). |
| `--limite N` | Procesa solo los N primeros usuarios. |

Opciones de conexión, comunes a los dos programas (y a `edp_api.py`):

| Opción | Efecto |
|---|---|
| `--login-navegador` | Iniciar sesión a mano en una ventana de Chromium (SSO o sin `.env`). |
| `--forzar-login` | Entrar aunque haya otra sesión abierta con la cuenta (la cierra). |
| `--url URL` | Otra dirección del backoffice (por defecto la de producción). |
| `--pausa SEGUNDOS` | Espera tras cada escritura (por defecto 0.2) para no saturar el servidor. |

`python3 <programa> --help` muestra esta ayuda en la terminal.

## 10. Seguridad y buenas prácticas

- **Simula siempre primero** y revisa el informe antes de usar `--ejecutar`.
- **Prueba con `--limite 1`** la primera vez que uses un Excel nuevo.
- **Relanzar es seguro**: las filas ya hechas se detectan como caso 4.
- **Guarda los informes**: son el registro de qué se cambió, cuándo y quién.
- El fichero **`.env` contiene tu contraseña**: no lo compartas ni lo subas a ningún sitio.
- El informe de contraseñas incluye en el *Resumen* la contraseña común aplicada.
  Trátalo con el mismo cuidado que cualquier comunicación de contraseñas.
- El programa nunca crea ni borra usuarios.

## 11. Problemas frecuentes

| Mensaje / situación | Causa y solución |
|---|---|
| *Ya hay una sesión abierta con esta cuenta* | Tienes el backoffice abierto en el navegador. Ciérralo, o relanza con `--forzar-login` (cierra la sesión del navegador). |
| *Login fallido* | Revisa `EDP_EMAIL` y `EDP_PASSWORD` en `.env`. Si entras con Microsoft o Google, usa `--login-navegador`. |
| *Faltan EDP_EMAIL y EDP_PASSWORD* | No existe el `.env` o le faltan datos (ver sección 3). |
| *La entrada debe ser un Excel .xlsx* | Guarda tu fichero como `.xlsx` desde Excel. |
| *Falta la columna correo_antiguo* | La cabecera del Excel no tiene esa columna. Usa la plantilla. |
| *No encuentro la columna de correos* (contraseñas) | Indica la columna con `--columna`. |
| Fila **PENDIENTE** | No existe ningún usuario con esos correos (caso 3). El programa no lo crea: avisa a quien corresponda para que se registre o se cargue aparte con sus licencias. |
| Fila del caso 2 en **OMITIDO** | El usuario que ya tiene el correo nuevo tiene licencias activas, así que no se ha tocado a ninguno de los dos. Revisa las dos cuentas en el backoffice si hay que unificarlas. |
| Fila **ERROR** con *HTTP 409 USERNAME_EXIST* | Ese correo ya lo usa otro usuario. Normalmente se resuelve solo; si persiste, revisa ese correo en el backoffice. |
| *Interrumpida: se cerró la sesión* | Alguien entró con la misma cuenta forzando el acceso. Relanza el mismo Excel: lo ya hecho sale como caso 4. |
| *No se pudo conectar* | Sin conexión a internet o el backoffice no responde. Inténtalo más tarde. |