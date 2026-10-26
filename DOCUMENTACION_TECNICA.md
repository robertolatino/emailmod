# Documentación técnica

Cómo funcionan por dentro `cambiar_correos.py` y `reiniciar_password.py`. La guía
de uso está en [README.md](README.md).

---

## Índice

1. [Visión general](#1-visión-general)
2. [Por qué la API y no "clics" en la web](#2-por-qué-la-api-y-no-clics-en-la-web)
3. [La API del backoffice](#3-la-api-del-backoffice)
4. [Módulos y funciones](#4-módulos-y-funciones)
5. [Algoritmo de cambiar_correos.py](#5-algoritmo-de-cambiar_correospy)
6. [Algoritmo de reiniciar_password.py](#6-algoritmo-de-reiniciar_passwordpy)
7. [Gestión de errores y robustez](#7-gestión-de-errores-y-robustez)
8. [Rendimiento](#8-rendimiento)
9. [Pruebas realizadas](#9-pruebas-realizadas)
10. [Mantenimiento y ampliación](#10-mantenimiento-y-ampliación)
11. [Limitaciones y observaciones](#11-limitaciones-y-observaciones)

---

## 1. Visión general

```
                 ┌────────────────────────┐        ┌────────────────────────────┐
 usuarios.xlsx ─▶│   cambiar_correos.py   │──┐  ┌─▶│  informe_cambio_correos_   │
                 │ (casos 1-4, por fila)  │  │  │  │  <fecha>.xlsx              │
                 └───────────┬────────────┘  │  │  └─────────────┬──────────────┘
                             │               │  │                │ (hoja Resultados,
                             ▼               │  │                │  filas OK)
                 ┌────────────────────────┐  │  │                ▼
                 │       edp_api.py       │  │  │  ┌────────────────────────────┐
                 │  ClienteEDP: login,    │◀─┼──┼──│   reiniciar_password.py    │
                 │  búsqueda, licencias,  │  │  │  └─────────────┬──────────────┘
                 │  edición, registro     │  │  │                ▼
                 └───────────┬────────────┘  │  │  ┌────────────────────────────┐
                             │ HTTPS (JSON)  │  │  │ informe_passwords_<fecha>  │
                             ▼               │  │  │ .xlsx                      │
          publisher.edelvivesdigitalplus.com │  │  └────────────────────────────┘
                         /api/...            │  │
                                             ▼  │
                 ┌────────────────────────────┐ │
                 │      informe_excel.py      │─┘
                 │ leer_excel, guardar_informe│
                 │ crear_plantilla            │
                 └────────────────────────────┘
```

- **`edp_api.py`**: todo lo que habla con la plataforma.
- **`informe_excel.py`**: todo lo que lee o escribe Excel.
- **`cambiar_correos.py`** y **`reiniciar_password.py`**: la lógica de negocio de
  cada tarea. Usan los dos módulos anteriores.

Dependencias: `playwright` (cliente HTTP y, solo para `--login-navegador`,
Chromium) y `openpyxl` (Excel). El resto es librería estándar de Python 3.9.

## 2. Por qué la API y no "clics" en la web

El backoffice es una **SPA** (aplicación de una sola página, en React). El HTML
que sirve el servidor está vacío; todo lo que se ve lo pide el navegador a una
**API JSON** del mismo dominio (`/api/...`). Por ejemplo, al pulsar *Guardar* en
la ficha de un usuario, el navegador envía `PUT /api/users/<guid>` con los datos
del formulario.

Hay dos formas de automatizarlo con Playwright:

| | Automatizar la interfaz (clics) | Llamar a la API (lo que hacemos) |
|---|---|---|
| Velocidad | Segundos por usuario (cargar páginas, esperar) | ~0,5 s por usuario |
| Estabilidad | Se rompe si cambian botones o textos | Solo se rompe si cambia la API |
| Comprobar existencia | Una búsqueda en pantalla por correo | 150 correos por petición |
| Resultado | El mismo | El mismo: son las mismas peticiones que hace la web |

Se usa el cliente HTTP de Playwright (`APIRequestContext`), así que la única
dependencia sigue siendo Playwright.

**Cómo se averiguó la API.** Leyendo el código JavaScript público del backoffice
(`/static/js/*.js`): funciones como `function n(e){return o.A.get("users",{params:e})}`
revelan cada endpoint y sus parámetros. Después se verificó contra la plataforma
real con usuarios de prueba (sección 9). Si la plataforma cambia, se puede volver
a mirar con las herramientas de desarrollo del navegador (sección 10).

## 3. La API del backoffice

### 3.1 Base y formato

- Base: `https://publisher.edelvivesdigitalplus.com/api/`
- Peticiones y respuestas en JSON.
- Respuesta correcta: `{"status": "success", "data": {...}}`.
- Validación de esquemas: la API rechaza con `400 SCHEMA_VALIDATION_FAILED` los
  cuerpos que no cumplen su especificación (por ejemplo, una contraseña de login
  de menos de 3 caracteres).

### 3.2 Autenticación

**Login**

```
POST /api/login
{"email": "...", "password": "...", "source": "backoffice",
 "admitted_roles": ["R04", "R05", "R06", "R07"], "force_login": true (opcional)}

→ 200 {"status": "success", "data": {"token": "<JWT>", ...}}
```

- `admitted_roles`: roles que admite la pantalla de login del backoffice.
- El token se envía después en la cabecera `Authorization: <token>` (tal cual, sin "Bearer").
- **Una sesión por cuenta.** Si la cuenta ya tiene sesión abierta, el login
  responde `errorKey: ACTIVE_SESSION_EXISTS`. Con `force_login: true` se entra y la
  otra sesión se cierra. A esa otra sesión, la API le responde a partir de ese
  momento `403` con `errorKey: SESSION_REVOKED`.
- **Token caducado o ausente:** `403` (no 401) con
  `{"message": "Authorization failed", "status": "fail", "data": {"error": 401}}`.

**Login manual (`--login-navegador`).** Se abre Chromium en `/auth/login`. Tras
entrar, la web hace peticiones a `/api/` con la cabecera `Authorization`; el
programa la captura (`page.on("request")`) y cierra el navegador.

### 3.3 Errores

La API usa varios formatos. `extraer_error()` los unifica:

```json
{"status": "fail", "error": {"code": 409, "errorKey": "USERNAME_EXIST"}}
{"message": "Authorization failed", "status": "fail", "data": {"error": 401}}
{"code": "SCHEMA_VALIDATION_FAILED", "message": "...", "results": {"errors": [{"path": [...], "message": "..."}]}}
```

| Error | Significado | Cómo lo trata el programa |
|---|---|---|
| `409 USERNAME_EXIST` (o `errno 1062`) | El correo ya lo usa otro usuario | Caso 1 → pasa a caso 2; en el renombrado `_old` → prueba el siguiente sufijo |
| `403` + `data.error 401` | Token caducado | Vuelve a iniciar sesión y repite |
| `403 SESSION_REVOKED` | Otra sesión con la misma cuenta | Para y guarda el informe |
| `ACTIVE_SESSION_EXISTS` (login) | Sesión abierta en otro sitio | Mensaje: usar `--forzar-login` |
| `404 USER_NOT_FOUND` | El usuario no existe | Error de la fila |
| `429`, `5xx`, fallo de red | Problema temporal | Hasta 3 intentos, esperando 2 y 4 s |

### 3.4 Endpoints usados

| Método y ruta | Uso | Notas |
|---|---|---|
| `POST /api/login` | Iniciar sesión | Ver 3.2 |
| `GET /api/users/whoami` | Comprobar la sesión y saber con qué cuenta se ha entrado | `data.username` |
| `GET /api/users?usernames[]=a&usernames[]=b&pageSize=156` | Buscar usuarios por correo de acceso | Ver 3.5 |
| `GET /api/users/{guid}` | Ficha completa de un usuario | `data` = usuario |
| `PUT /api/users/{guid}` | Guardar la ficha | Ver 3.6 |
| `GET /api/users/{guid}/license-personals` | Licencias redimidas por el usuario | Ver 3.7 |

Hay otros endpoints que no usa el programa. Por ejemplo, `POST /api/users` crea
usuarios (el programa no crea ninguno: el caso 3 solo avisa) y `DELETE /api/users`
con `{"guid": [...]}` los borra; solo se usó para limpiar los usuarios de prueba.

### 3.5 Búsqueda por correos

```
GET /api/users?usernames[]=ana@x.com&usernames[]=luis@y.com&pageSize=156
→ {"data": {"users": [ {...}, ... ], "total": N, "offset": 0, "pageSize": 156, "left": ...}}
```

- Es la búsqueda que usa la web en el diálogo *"Pega una lista de emails"*, que
  admite hasta **156** correos. El programa manda lotes de **150**.
- **Coincidencia exacta y sin distinguir mayúsculas** (verificado): `ANA@X.COM`
  encuentra `ana@x.com`, y un trozo (`x.com`) no encuentra nada. Aun así, el
  programa filtra por coincidencia exacta.
- Cada usuario trae, entre otros: `guid`, `username` (correo de acceso), `email`
  (contacto), `name`, `lastname`, `role_guid`, `is_demo`, `country_guid`,
  `schools`, `change_password`, `is_active`.
- Los usuarios borrados no aparecen.

### 3.6 Edición (`PUT /api/users/{guid}`)

La ficha del backoffice, al pulsar *Guardar*, envía siempre estos campos:
`username, name, lastname, role_guid, is_demo, country_guid` (más `password` y
`change_password` si se rellena la contraseña).

El programa hace lo mismo:

1. Lee la ficha actual (`GET /api/users/{guid}`).
2. Construye el cuerpo con los valores actuales de esos campos (`datos_edicion()`).
3. Sustituye solo lo que hay que cambiar.
4. Envía el `PUT`.

Así los demás datos se reenvían con su mismo valor y no cambian. **Verificado:**
tras un cambio de correo, comparando la ficha completa antes y después, solo
cambian `username`, `email` (si coincidía), `updated_at` y `searches` (un índice
interno de búsqueda que la plataforma recalcula).

Campos relevantes:

| Campo | Contenido |
|---|---|
| `username` | Correo de acceso. Único en la plataforma. |
| `email` | Email de contacto. **No** es único. A veces igual a `username`, a veces vacío. |
| `password` | Contraseña nueva, en claro (la plataforma la cifra). |
| `change_password` | `1` = la plataforma pide al usuario una contraseña nueva al entrar. |

### 3.7 Licencias de un usuario (`GET /api/users/{guid}/license-personals`)

Es la petición que hace la pestaña *Licencias* de la ficha del usuario. Devuelve
una lista con las licencias redimidas, caducadas incluidas:

```
→ {"status": "success", "data": [
     {"license_guid": "...", "code": "EDP-YjFbgPYg", "linked_at": "2023-09-19T15:21:19.000Z",
      "expired_at": "2027-09-30T00:00:00.000Z", "role_guid": "R02",
      "programs": [...], "contents": [...]},
     ...]}
```

- `linked_at` es la fecha de redención (columna *Fecha inicio* de la pestaña) y
  `expired_at` la de caducidad (*Fecha final*), en UTC.
- **Licencia activa = `expired_at` todavía no ha llegado** (`licencia_activa()`).
  Es el mismo criterio que la columna *Estado* de la pestaña, que la web calcula
  en el navegador: muestra los días que quedan hasta `expired_at` o *Caducada*.
  Sin fecha de caducidad, o con una fecha que no se entiende, se considera activa,
  para no tocar al usuario ante la duda.
- No hay paginación: la API ignora `pageSize`, `offset` y `page`, y devuelve lo
  mismo que muestra la pestaña (verificado con cuentas de 99 y 100 licencias).
- El endpoint se encontró en el código de la web (función que pinta la pestaña
  *Licencias*) y se verificó en solo lectura con usuarios reales.

### 3.8 Roles

| Código | Rol |
|---|---|
| R01 | alumno |
| R02 | profesor |
| R03 | admin colegio |
| R04 | admin (backoffice) |
| R05 | helpdesk (backoffice) |
| R06 | admin fixed layout (backoffice) |
| R07 | publisher (backoffice) |

El informe muestra el nombre del rol (`NOMBRES_ROL` en `edp_api.py`).

## 4. Módulos y funciones

### 4.1 `edp_api.py` — conexión con la plataforma

| Elemento | Qué hace |
|---|---|
| `ClienteEDP` | Cliente de la API. Guarda el token, envía las peticiones y anota cada una en `registro`. |
| `.login(email, password, forzar)` | Login con credenciales (3.2). |
| `.login_navegador()` | Login manual en Chromium capturando el token. |
| `.quien_soy()` | `GET users/whoami`; guarda `usuario_sesion`. |
| `.buscar_por_correos(correos)` | Búsqueda en lotes de 150. Devuelve `{correo_minúsculas: usuario}` solo con los que existen. |
| `.obtener_usuario(guid)` | Ficha completa. |
| `.actualizar_usuario(guid, cambios, accion)` | Lee la ficha, aplica `cambios` y la guarda (3.6). Devuelve la ficha anterior. |
| `.cambiar_username(guid, nuevo)` | Cambia el correo de acceso y, si coincidía, el email de contacto. Devuelve si cambió el email. |
| `.licencias_activas(guid)` | Licencias del usuario que no han caducado (3.7). |
| `._peticion(...)` / `._enviar(...)` | Envío con reintentos, renovación del token y errores (sección 7). |
| `._anotar(...)` | Añade una entrada al registro. |
| `ErrorAPI` | Error de la API, con `status`, `error_key`, `errno` y la propiedad `usuario_ya_existe`. |
| `SesionRevocada` | Subclase de `ErrorAPI` para `SESSION_REVOKED`. |
| `ErrorLogin` | No se pudo iniciar sesión. |
| `datos_edicion(usuario)` | Campos que se reenvían al guardar (3.6). |
| `describir_cambios(actual, cambios)` | Texto "campo: antes → después" para el registro; oculta la contraseña. |
| `licencia_activa(licencia, ahora)` | Si una licencia no ha caducado (3.7). |
| `cargar_env()` | Lee `.env` (sin dependencias externas). |
| `agregar_args_conexion()` / `conectar()` | Opciones de conexión comunes e inicio de sesión. |

### 4.2 `informe_excel.py` — Excel

| Elemento | Qué hace |
|---|---|
| `leer_excel(ruta)` | Lee el `.xlsx` (hoja *Usuarios* → *Resultados* → la primera). Normaliza la cabecera, salta filas vacías y añade `_fila` (nº de fila real). Rechaza lo que no sea `.xlsx`. |
| `normalizar_columna()` | `"Correo antiguo"` → `"correo_antiguo"`. |
| `guardar_informe(...)` | Escribe las 4 hojas (Resumen, Resultados, Registro API, Leyenda) con estilos, filtros y colores. |
| `Columna` | Definición de una columna del informe: clave, título, ancho y descripción. La descripción aparece en la Leyenda. |
| `seccion_peticiones(registro)` | Estadísticas de peticiones para el Resumen. |
| `formatear_duracion()` | `125` → `"2 min 5 s"`. |
| `crear_plantilla(ruta)` | Genera `plantilla_entrada.xlsx` (`python3 informe_excel.py plantilla_entrada.xlsx`). |

### 4.3 `cambiar_correos.py`

| Elemento | Qué hace |
|---|---|
| `main()` | Orquesta los pasos 1–5 (sección 5) y garantiza el informe aunque haya interrupciones. |
| `resolver_columnas()` | Comprueba las columnas obligatorias y acepta alias. |
| `avisos_duplicados()` | Detecta correos repetidos en la entrada. |
| `buscar_todos()` | Búsqueda inicial de todos los correos → índice. |
| `clasificar(existe_antiguo, existe_nuevo)` | Devuelve el caso (1–4). |
| `procesar_fila()` | Valida la fila, la clasifica, comprueba las licencias en el caso 2, rellena los datos del informe y, si se ejecuta, llama a `ejecutar_caso`. |
| `nuevo_con_licencias()` | Caso 2: si el usuario que tiene el correo nuevo tiene licencias activas, deja la fila en OMITIDO. |
| `ejecutar_caso()` | Aplica la acción del caso 1 o 2 en la plataforma. |
| `cambiar_correo()` | Cambia el correo en la plataforma y en el índice. |
| `liberar_correo()` | Caso 2: renombra al usuario que ocupa el correo nuevo (`_old`, `_old2`…). |
| `ResultadoFila` | Resultado de una fila; cada campo es una columna del informe. |
| `CASOS`, `ESTADOS`, `COLUMNAS` | Textos de casos y estados, y columnas del informe (también alimentan la Leyenda). |

### 4.4 `reiniciar_password.py`

| Elemento | Qué hace |
|---|---|
| `main()` | Orquesta los pasos 1–4 (sección 6). |
| `preparar_filas()` | Decide qué filas se procesan: descarta sin correo, repetidas, no OK (si es un informe) y fuera de `--limite`, anotando el motivo. |
| `escribir_informe()` | Compone el Resumen (incluida la contraseña aplicada) y guarda el informe. |

## 5. Algoritmo de `cambiar_correos.py`

**Paso 1. Leer y validar.** `leer_excel()` + `resolver_columnas()` +
`avisos_duplicados()`. Inicio de sesión con `conectar()`.

**Paso 2. Búsqueda masiva → índice.** Se buscan de una vez:

- todos los correos antiguos,
- todos los correos nuevos,
- todos los `<correo nuevo>_old`, para saber de antemano qué sufijos están ocupados.

El resultado es el **índice**, un diccionario en memoria
`{correo en minúsculas: usuario}`. A partir de aquí, clasificar una fila no
cuesta ninguna petición.

**Paso 3. Recorrer las filas en orden.** Para cada fila:

```
correo nuevo vacío ───────────────────────▶ ERROR   "Fila incompleta"
antiguo == nuevo ─────────────────────────▶ OMITIDO "Son iguales"

¿antiguo en el índice?   (correo antiguo vacío = "no")
├── sí ─ ¿nuevo en el índice?
│        ├── no ──▶ CASO 1: cambiar_username(antiguo → nuevo)
│        │            └─ si la API responde 409 (el nuevo se registró mientras tanto):
│        │                 buscar el nuevo y seguir como CASO 2
│        └── sí ──▶ CASO 2: ¿el usuario del nuevo tiene licencias activas?
│                           ├── sí ─▶ OMITIDO (no se toca a ninguno de los dos)
│                           └── no ─▶ liberar_correo(nuevo → nuevo_old | _old2 | …)
│                                     cambiar_username(antiguo → nuevo)
└── no ─ ¿nuevo en el índice?
         ├── no ──▶ CASO 3: PENDIENTE (no se crea; hay que avisar)
         └── sí ──▶ CASO 4: OMITIDO (no se hace nada)
```

En **simulación** se clasifica igual y se rellena el informe con lo que se haría,
incluido el correo `_old` previsto, pero no se llama a ningún endpoint de escritura.
Las licencias del caso 2 sí se leen (es una lectura), así que la simulación ya
distingue los caso 2 que se van a quedar sin tocar.

Tras cada cambio se **actualiza el índice**: el correo antiguo sale y entra el
nuevo. Así una fila posterior con los mismos correos ve el estado real. Por
ejemplo, si dos filas tienen el mismo correo antiguo, la segunda verá que ya no
existe.

**Caso 2, en detalle:**

1. Se leen las licencias del usuario que tiene el correo nuevo
   (`licencias_activas()`, ver 3.7). Si tiene alguna sin caducar, la fila queda
   **OMITIDO**, con el número de licencias activas en el informe, y no se hace
   nada más. Si la lectura falla, la fila queda en ERROR sin haber tocado nada.
2. Se elige el primer correo libre de `nuevo_old`, `nuevo_old2` … `nuevo_old10`
   (los que ya están en el índice se saltan).
3. Se renombra al usuario que tiene el correo nuevo. Si la API responde 409, se
   prueba el siguiente sufijo.
4. Se cambia el usuario antiguo al correo nuevo.
5. Si el paso 4 falla, la fila queda en ERROR y el detalle explica que el
   renombrado sí se hizo. No se deshace: al relanzar, el correo nuevo ya está libre
   y la fila será un caso 1.

**Caso 3.** No se crea el usuario: para darlo de alta hay que saber qué licencias
cargarle, así que se hace aparte. La fila queda **PENDIENTE** y, si el Excel trae
Nombre, Apellidos y Rol, se copian al informe para saber a quién avisar.

**Idempotencia.** Relanzar el mismo Excel es seguro. Tras un caso 1 o 2 el correo
antiguo ya no existe y el nuevo sí, así que la fila es un caso 4. Un caso 2 con
licencias o un caso 3 vuelven a salir igual mientras nadie cambie esas cuentas.

**Paso 4. Interrupciones.** Si la sesión se revoca, se pulsa Ctrl+C o hay un error
inesperado fuera de una fila (por ejemplo, falla la búsqueda inicial), el bucle se
detiene. Las filas que faltan se marcan **NO PROCESADO**.

**Paso 5. Informe.** Siempre se escribe, en el bloque `finally` de `main()`.

## 6. Algoritmo de `reiniciar_password.py`

1. **Preparar filas** (`preparar_filas`). Para cada fila del Excel:
   - sin correo → ERROR;
   - es un informe, la fila no está en OK y no hay `--todas` → OMITIDO;
   - correo ya visto en una fila anterior → OMITIDO;
   - se superó `--limite` → OMITIDO;
   - si no, queda pendiente de procesar.
2. **Búsqueda masiva** de los correos pendientes, en lotes de 150.
3. **Para cada pendiente:**
   - no encontrado → NO ENCONTRADO;
   - simulación → SIMULADO;
   - si no, `actualizar_usuario(guid, {"password": X, "change_password": 1})`
     (sin `change_password` con `--sin-forzar-cambio`) → OK o ERROR.
4. **Informe**, siempre, con todas las filas del Excel de entrada.

**Verificado en la plataforma real:** tras el reinicio, el usuario puede entrar en
el frontoffice con la contraseña nueva (`@Edelvives` y `Edelvives2027`), la
anterior deja de funcionar y `change_password` queda a 1.

## 7. Gestión de errores y robustez

### 7.1 Reintentos (`ClienteEDP._enviar`)

- **Fallo de red, timeout (30 s), 429 o 5xx:** hasta 3 intentos, esperando 2 y 4
  segundos.
- **Token rechazado:** nuevo login con las mismas credenciales y se repite. Si ese
  login falla (por ejemplo, porque hay otra sesión abierta), se trata como sesión
  revocada.
- Las ediciones (`PUT`) se pueden reintentar sin riesgo: repetirlas da el mismo
  resultado.

### 7.2 Errores por fila

Cualquier error al procesar una fila se recoge en esa fila (estado ERROR y el
mensaje en *Detalle*) y el programa sigue con la siguiente. Solo detienen la
ejecución:

- la sesión revocada (seguir fallaría en todas las filas),
- Ctrl+C,
- un fallo en la búsqueda inicial.

En todos los casos se escribe el informe.

### 7.3 Registro de peticiones

Cada petición queda en `ClienteEDP.registro` y en la hoja **Registro API**: fecha
y hora, acción en lenguaje llano, método, ruta, código HTTP, intentos, duración y
resultado. Las acciones de cambio indican el valor anterior y el nuevo
(`username: a@x.com → b@x.com`). **La contraseña nunca se anota**: aparece como
`contraseña: (oculta)`.

### 7.4 Pausa entre escrituras

Tras cada escritura se espera `--pausa` segundos (0,2 por defecto) para no
cargar el servidor en lotes grandes.

## 8. Rendimiento

Tiempos medidos en la plataforma real:

| Petición | Duración media |
|---|---|
| Login | ~360 ms |
| Búsqueda de un lote (hasta 150 correos) | ~95 ms |
| Leer ficha | ~165 ms |
| Guardar ficha | ~145 ms |
| Leer licencias | ~110–360 ms (según cuántas tenga) |

Coste por fila:

| Caso | Peticiones | Tiempo aproximado |
|---|---|---|
| 1 | 1 lectura + 1 escritura | ~0,5 s |
| 2, sin licencias activas | 3 lecturas (licencias + 2 fichas) + 2 escrituras | ~1–1,5 s |
| 2, con licencias activas | 1 lectura (licencias) | ~0,2 s |
| 3, 4, OMITIDO | 0 | — |

Búsqueda inicial: unas 3 direcciones por fila (antiguo, nuevo y nuevo_old) entre
150 por petición, es decir, unas 20 peticiones para 1.000 filas (~2 s).
**Total orientativo: 1.000 filas de caso 1 ≈ 8–9 minutos.**

## 9. Pruebas realizadas

1. **Simulador local.** Un servidor HTTP que imita la API, con fallos provocados:
   un 503 (se reintenta), caducidad del token (nuevo login), un correo registrado
   a mitad de ejecución (paso automático del caso 1 al caso 2), un `_old` ya ocupado
   (se usa `_old2`), filas vacías, incompletas o duplicadas y mayúsculas.
2. **Plataforma real, solo lectura.** Formato de la búsqueda (exacta, sin
   distinguir mayúsculas), campos de la ficha y existencia del campo `email`.
3. **Plataforma real, escritura**, siempre con usuarios de prueba propios
   (`claude.prueba.<código>.*@example.com`; `example.com` es un dominio reservado
   que no entrega correo), borrados al terminar:
   - los cuatro casos de principio a fin, con el Excel de entrada y los informes;
   - comparación completa de la ficha antes y después (solo cambia el correo);
   - error 409 `USERNAME_EXIST` al usar un correo ocupado;
   - reinicio con `@Edelvives` y con `Edelvives2027` y login en el frontoffice con
     cada una;
   - relanzar el mismo Excel (todo sale como caso 4);
   - rechazo de una entrada `.csv`.
4. **Comprobación de licencias en el caso 2 y caso 3 sin alta** (cambio posterior):
   - cliente falso en memoria con todos los caminos: caso 2 con y sin licencias,
     paso del caso 1 al 2 con y sin licencias, caso 3 sin alta, simulación, fallo
     al leer las licencias y fechas de caducidad límite, vacías o raras;
   - plataforma real, **solo lectura**: forma de la respuesta de
     `license-personals`, paginación (no la hay) y una simulación completa con un
     Excel que provoca los cuatro casos (un caso 2 con 36 licencias activas quedó
     OMITIDO y uno sin licencias, SIMULADO).

Ningún usuario real se modificó durante las pruebas.

## 10. Mantenimiento y ampliación

**Si la plataforma cambia y algo deja de funcionar:**

1. Abre el backoffice en Chrome, pulsa F12, ve a la pestaña *Red* (Network) y
   filtra por `api`.
2. Haz a mano la operación (buscar un usuario, guardar su ficha…) y mira la
   petición: método, ruta, cuerpo (*Payload*) y respuesta.
3. Ajusta el método correspondiente de `ClienteEDP` en `edp_api.py`. Toda la
   comunicación con la API está en esa clase.

**Cambios habituales:**

| Quiero… | Dónde |
|---|---|
| Cambiar la contraseña por defecto | `PASSWORD_DEFECTO` en `reiniciar_password.py` |
| Cambiar el sufijo por defecto | Opción `--sufijo` (default en `main()` de `cambiar_correos.py`) |
| Cambiar qué cuenta como licencia activa | `licencia_activa()` en `edp_api.py` |
| Aceptar otro título de columna de correo | `ALIAS_COLUMNAS` en `cambiar_correos.py` / `COLUMNAS_CORREO` en `reiniciar_password.py` |
| Añadir una columna al informe | Añade una `Columna` a `COLUMNAS` y el campo a `ResultadoFila` (o a la fila en `reiniciar_password.py`) |
| Que el email de contacto no se toque nunca | `ClienteEDP.cambiar_username`: quitar las dos líneas que añaden `email` |
| Cambiar los colores de los estados | `COLORES_ESTADO` en `informe_excel.py` |
| Regenerar la plantilla | `python3 informe_excel.py plantilla_entrada.xlsx` |

**Comprobar la conexión** sin cambiar nada: `python3 edp_api.py correo@dominio`.

## 11. Limitaciones y observaciones

- La API no está documentada públicamente: Edelvives (o su proveedor) puede
  cambiarla sin aviso. Ver sección 10.
- El programa no crea usuarios: los del caso 3 hay que darlos de alta aparte.
- En el caso 2 solo cuentan las licencias de la pestaña *Licencias* de la ficha
  (`license-personals`). El acceso MINT del usuario (`has_license_mint`) no se mira.
- El caso 2 prueba hasta `_old10`. Si están todos ocupados, la fila da ERROR.
- **Observación de seguridad, para trasladar al proveedor:** el listado de
  usuarios de la API (`GET /api/users`) devuelve al navegador el hash de la
  contraseña (`password`) y el campo `code_recovery` de cada usuario. Lo
  recomendable es que la API no exponga esos campos.
