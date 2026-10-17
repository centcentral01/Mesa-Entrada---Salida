# Mesa de Entradas (CENT) — contexto para retomar

Registro de notas de **entrada** y **salida** de la CENT (UTN) que lleva Luis Lombardi. Antes era papel:
sello y número a mano. El objetivo es tener cada número registrado y poder responder al instante
"¿qué nota salió con el número X?", "¿se contestó tal oficio?, ¿con qué nota?".

**Al retomar:** leer "Estado" y "Pendientes" al final de este archivo. Al terminar una sesión con
avances, actualizarlos.

## Reglas de trabajo

- **Git:** todo va a `main`. Después de cada cambio con sentido, commit y push a `origin`
  (`https://centcentral01@github.com/centcentral01/Mesa-Entrada---Salida.git`) **sin preguntar**.
  Push: `GCM_INTERACTIVE=never GIT_TERMINAL_PROMPT=0 git push origin main`.
- **Correo:** solo **Outlook clásico de escritorio** (`OUTLOOK.EXE`, el mismo donde está
  viaticoscent@). La UTN tiene deshabilitada la autenticación básica: nada de IMAP/SMTP con clave.
  **Nunca guardar la clave de llombardi@** en archivos ni memoria. Leer sin mover ni marcar nada.
  Outlook web (outlook.cloud.microsoft) no sirve para scripts.
- **Supabase:** no tocar nada fuera del esquema `mesa_entradas` (en el mismo proyecto están
  `public` con viáticos y `convenios`).
- Antes de grabar datos, armar la planilla de revisión (`revision/*.csv`); los scripts graban solo
  con `--grabar` y son idempotentes (upsert o salteo por `correo_id`).

## Dónde vive cada cosa

- **Base:** Supabase, proyecto "Gestión de Viáticos" (ref `ojdwtslgkngfurzkjqzg`), esquema
  **`mesa_entradas`** (expuesto en Data API > Settings > Exposed schemas; solo `service_role` tiene
  permisos). Definición completa en [db/schema.sql](db/schema.sql).
- **`.env` propio** de esta carpeta (no se sube): `SUPABASE_URL`, `SUPABASE_SERVICE_ROLE_KEY`,
  `DB_*`, `POOLER_HOST`. El host directo `db.<ref>.supabase.co` es solo IPv6 y no resuelve: se entra por
  el pooler `aws-1-us-east-1.pooler.supabase.com` con usuario `postgres.<ref>`
  ([scripts/psql_supabase.sh](scripts/psql_supabase.sh) lo arma).
- **Python:** el venv de viáticos, `C:\Proyectos\PythonEnvs\viaticos_casa\Scripts\python.exe`
  (tiene streamlit, supabase, pywin32, pdfplumber).
- **Documentos 2026:** OneDrive de llombardi@,
  `C:\Users\Usuario\Personal CENT\OneDrive - Universidad Tecnologica Nacional\Mesa de Entradas\2026\`
  → `Salidas\npNNNN - ….pdf` y `Entradas\<fecha> - <asunto>\` (correo `.msg` + adjuntos).
  `ruta_archivo` en la base apunta ahí.
- **Fuera del repo** (`.gitignore`): `.env`, `revision/` (planillas de control), `correo/` (adjuntos
  bajados de Outlook; `correo/descartadas` = notas de 2021 y la 0186/25 que vinieron adjuntas en 2026).

## Modelo

- **Dos tablas**, `entradas` y `salidas` (decisión del usuario: registros chicos, numeración propia).
  Catálogos `areas`, `organismos`, `responsables` creados pero sin usar todavía (se usa texto libre:
  `remitente` / `destinatario`, "tal como figura").
- **Numeración:** tabla `numeradores` (tipo, año). Trigger `asignar_numero`: si no viene número asigna
  el siguiente; si viene a mano, lo respeta y adelanta el contador. Reinicia cada año.
- **Entradas:** `medio` = `Papel` (lleva número de mesa) o `Correo` (sin número). OJO: el oficio de
  Valencia Huaracallo dice "Mesa de entrada 001-2026": puede que las de correo también se numeren
  (a confirmar con el usuario).
- **Salidas:** `sufijo` ('' / 'b' / 'c') para notas viejas distintas que salieron con el mismo número;
  la app nunca asigna sufijo. `UNIQUE (anio, numero, sufijo)`.
- **Respuesta:** `salidas.entrada_id` → la entrada que contesta. Trigger `marcar_respondida` pasa la
  entrada a `Respondido`. Lo que no requiere respuesta se carga como `Archivado` (opción elegida en vez
  de un campo `requiere_respuesta`).
- **Vista `v_entradas`:** entradas + `respondida`, `respuestas` ("0016/26, 0043/26") y `vencida`
  (calculada contra hoy; sin fecha de vencimiento nunca vence). Tras agregar columnas a `entradas`
  hay que recrearla (toma `e.*` al crearse) y `NOTIFY pgrst, 'reload schema'`.

## Fuentes de datos

**Salidas** = archivos `npNNNN - <destinatario + asunto>.doc/.pdf` en carpetas "Notas AAAA"
(las mantiene Oscar; son los borradores en Word, sin firma: valen como registro del número, no como
documento). Dentro: "Buenos Aires, 18 de septiembre de 2025", "Nota Nº 0050/25" (también "Nº: 0020/2023"
o "00012/24"), expediente `EX-…` / `Expte. …` / `NO-…`.
- 2023: `Downloads\OneDrive_2026-10-17 (1).zip` · 2024 y 2025: `Downloads\OneDrive_1_16-10-2026.zip`.
- 2026: adjuntos de los correos "Responde oficio.-" de Neifert (parcial); falta la carpeta de Oscar.

**Entradas por correo** (buzón llombardi@): los oficios judiciales no llegan directo.
- Dirección Ejecutiva los reenvía ("RV: OFICIO …") o Luis se los pasa a Neifert ("Oficio Judicial - …").
- Neifert devuelve "Responde oficio.-" con la nota de salida `np00XX` adjunta.
- De `@cent.utn.edu.ar` **solo cuentan** `vneifert@` y `direccionejecutiva@`. Se excluyen: el resto
  del dominio CENT, auditorías `ar010@`…`ar140@`, amartinez, Liliana, Facundo y Nahuel (CNRT), la CCIV,
  las casillas propias de Luis y los auditores que también escriben a viaticoscent@.

**Entradas en papel:** el usuario las escanea (con su número de mesa); todavía no hay importador.

## Scripts (correr desde `C:\MesaEntradas` con el python del venv)

| Script | Qué hace |
|---|---|
| `scripts/importar_salidas.py <zip o carpeta> [--anio N] [--grabar]` | Notas npNNNN → `salidas`. Une archivos de la misma nota (PDF de un Word, "- Copy", versión corregida) y pone sufijo a notas distintas con el mismo número. |
| `scripts/correos_entrantes.py` | Lista del año de posibles notas entrantes en llombardi@ (reglas de exclusión arriba). |
| `scripts/adjuntos_notas.py [--anio N]` | Baja de todas las carpetas del buzón los adjuntos np/"Nota CENT N°" del año a `correo/Notas <año>/`. |
| `scripts/entradas_correo.py [--grabar]` | Arma las entradas desde el correo, une el mismo oficio por carátula y vincula la salida posterior que comparte apellidos. |
| `scripts/archivar_onedrive.py` | Copia salidas y entradas del año a OneDrive y guarda `ruta_archivo`. |
| `app.py` (`iniciar.bat`, puerto 8510) | Streamlit: Notas de salida (año, destinatario, texto, Nº, expediente, fechas, "solo para revisar", detalle, CSV) y Entradas (año, emisor, estado, medio, vencidas). |

**Mañas conocidas:** rutas de Windows > 260 caracteres (los nombres de notas son largos: se copian a
un temporal con nombre corto, y en OneDrive se acortan carpetas/archivos); `psql` desde bash rompe los
acentos en `-c` (usar el cliente de Supabase en Python para textos con tildes); el modo automático
de Claude a veces frena la lectura del `.env` o el push.

## Estado (2026-10-17)

- `salidas`: 2023 = 326, 2024 = 206, 2025 = 172, 2026 = 25 (números 1 a 77, parcial).
  Contadores: 2023→326, 2024→203, 2025→186, 2026→77.
- `entradas`: 13 de 2026 por correo, 9 vinculadas a su respuesta. Pendientes de respuesta:
  León Valverde (22/07), Poder Judicial de Córdoba (28/08, ¿respondido con la 0064/26?),
  Canteros c/ Molina (15/10). Zorrilla: la 0021/26 respondió un primer oficio que no está en el correo.
- Observaciones para revisar en 2023–2025 (columna `observaciones`; en la app, "Solo para revisar"):
  2023 np0024–0035 dicen otro número (renumeradas), np0284 quedó como dos notas (probables versiones),
  2024 np0146 b es un anexo; números sin archivo: 2023 → 29, 84, 220, 309; 2025 → 3, 4, 6, 149–159.
- Destinatario flojo en 2023/2024 (los nombres no siguen "destinatario + verbo"); se podría sacar
  del encabezado del documento si hace falta.
- Se le mandaron a Luis por correo el resumen con enlaces y un recordatorio para el lunes 19/10.

## Pendientes

1. **Lunes 19/10:** importar la carpeta `Notas 2026` de Oscar (`importar_salidas.py`, luego
   `archivar_onedrive.py`) y re-correr `entradas_correo.py` para vincular las entradas pendientes.
2. **Escaneos de lo inventariado:** entradas en papel (número de mesa) y copias selladas de las
   salidas → hacer el importador y vincularlas.
3. Confirmar si las entradas por correo también llevan número de mesa ("Mesa de entrada 001-2026").
4. App: carga de nueva salida (que la app dé el número para el sello) y alta/edición de entradas.
5. Opcional: firmante de las salidas ("emisor") desde el pie del documento.
