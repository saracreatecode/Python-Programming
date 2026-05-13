# outlook-contact-sync

Automatización para limpiar, validar y sincronizar una base de contactos en Excel con Microsoft Outlook Classic en un entorno de red local Windows.

---

## Contexto

La empresa gestiona una base de contactos en un fichero Excel con múltiples hojas (clientes, proveedores, ingenieros, constructoras, arquitectos, oficina e internacionales). El objetivo es que todos los empleados tengan exactamente los mismos contactos actualizados en su Outlook, sin depender de Exchange ni de Microsoft 365 empresarial, y sin que ningún dato salga de la red local.

La solución funciona en dos pasos:

1. **Limpiar y validar** el Excel original → genera un Excel limpio
2. **Sincronizar** el Excel limpio con Outlook en cada equipo

---

## Estructura del proyecto

```
outlook-contact-sync/
├── 01_limpiar_excel.py        # Limpieza, validación y normalización del Excel
├── 02_sincronizar_outlook.py  # Sincronización del Excel limpio con Outlook
├── logs/                      # Logs generados automáticamente (no versionar)
│   ├── limpieza_log.txt
│   └── sincronizacion_log.txt
└── README.md
```

---

## Requisitos

- Windows con Microsoft Outlook Classic instalado
- Python 3.10 o superior con la casilla **"Add Python to PATH"** marcada en la instalación
- Acceso de red a la carpeta compartida del servidor donde vive el Excel

### Instalar dependencias

```bash
pip install pandas openpyxl
pip install pywin32
```

---

## Configuración

Antes de ejecutar, abre cada script y ajusta las rutas de configuración al inicio del fichero:

### `01_limpiar_excel.py`

```python
RUTA_EXCEL_ORIGINAL = r"\\SERVIDOR\Compartido\contactos.xlsx"
RUTA_EXCEL_LIMPIO   = r"\\SERVIDOR\Compartido\contactos_limpio.xlsx"
RUTA_LOG            = r"\\SERVIDOR\Compartido\logs\limpieza_log.txt"
```

### `02_sincronizar_outlook.py`

```python
RUTA_EXCEL_LIMPIO        = r"\\SERVIDOR\Compartido\contactos_limpio.xlsx"
RUTA_LOG                 = r"\\SERVIDOR\Compartido\logs\sincronizacion_log.txt"
CARPETA_RAIZ_OUTLOOK     = "Contactos Empresa"
```

---

## Uso

Ejecutar siempre en este orden:

```bash
# Paso 1 — limpiar el Excel original
python 01_limpiar_excel.py

# Paso 2 — sincronizar con Outlook (con Outlook abierto)
python 02_sincronizar_outlook.py
```

El Script 2 **requiere que Outlook esté abierto** en el equipo donde se ejecuta.

---

## Estructura del Excel de origen

El Excel tiene las siguientes hojas activas (la hoja G no se sincroniza):

| Hoja | Contenido | Carpeta en Outlook |
|------|-----------|-------------------|
| A | Clientes / Actores clave | Clientes_ActoresClave |
| B | Proveedores | Proveedores |
| C | Ingenieros / Consultores / Legal | Ingenieros_Consultores_Legal |
| D | Constructoras | Constructoras |
| E | Arquitectos / Municipales / Permisos | Arquitectos_Municipales |
| F | Oficina / Consumibles / Banco | Oficina_Consumibles |
| G | Acciones / Seguimiento | ⛔ No sincronizada |
| H | Internacionales | Internacionales |

### Columnas del Excel (hojas A–H excepto G)

```
NOMBRE, 1er APELLIDO, 2º APELLIDO, EMPRESA / RAZÓN SOCIAL, TIPO, ÁREA, CARGO,
MAIL, TLF (columna combinada → TLF1 + TLF2), AÑO ÚLTIMO TRABAJO, NOTA,
DIRECCIÓN, WEB, ACCIONES 1–5 (ACCIONES 1–6 en hoja A), ACTIVO
```

La hoja H añade además la columna `PAÍS`.

> **Nota:** La cabecera de la hoja A está en la fila 4. El resto de hojas tienen la cabecera en la fila 3. La columna A está vacía en todas las hojas; los datos empiezan en la columna B.

---

## Lógica de sincronización

| Situación | Comportamiento |
|-----------|---------------|
| Contacto nuevo en Excel | Se crea en Outlook |
| Contacto modificado en Excel | Se actualiza en Outlook |
| Contacto con `ACTIVO=NO` | Se omite (no se crea ni actualiza) |
| Contacto eliminado del Excel | Aviso en el log; no se borra automáticamente de Outlook |
| Email duplicado entre hojas | Se mantiene en ambas carpetas con sus categorías; aviso en log |
| Email con formato inválido | Se importa el contacto sin email; aviso en log |
| Teléfono con texto no numérico | Se limpia automáticamente; si no es posible, se omite ese campo |
| Nombre vacío | El contacto se omite completamente; error en log |

### Campos sin equivalente nativo en Outlook

Los campos `ÁREA`, `TIPO`, `AÑO ÚLTIMO TRABAJO`, `NOTA` y `ACCIONES 1–6` no tienen campo directo en Outlook. Se vuelcan en el campo **Notes** con este formato:

```
Área: Comercial
Tipo empresa: Proveedor
Año últ. trabajo: 2022
──────────────────────────────
Acción 1: Llamada de seguimiento
Acción 2: Presupuesto enviado
──────────────────────────────
Nota: Cliente prioritario renovación anual
```

---

## Gestionar bajas de contactos

Para dar de baja un contacto **no borres la fila del Excel**. En su lugar, pon `NO` en la columna `ACTIVO`. Así:

- El contacto no se sincroniza con Outlook
- El historial se conserva en el Excel
- El log avisará de que existe en Outlook pero ya no está activo en el Excel, para que se borre manualmente si se desea

---

## Automatización diaria (Tarea Programada de Windows)

Para que la sincronización se ejecute automáticamente cada día en cada equipo:

1. Abre el **Programador de tareas** de Windows
2. Crea una tarea nueva con estas opciones:
   - **Desencadenador:** Diariamente a las 07:45
   - **Acción:** Iniciar programa
   - **Programa:** `C:\Python312\python.exe` *(ajusta tu ruta)*
   - **Argumentos:** `\\SERVIDOR\Compartido\scripts\02_sincronizar_outlook.py`
3. En **Configuración:** activa "Ejecutar la tarea lo antes posible si se perdió un inicio programado"

> El Script 1 (`01_limpiar_excel.py`) solo necesita ejecutarse cuando se modifica el Excel original, no necesariamente a diario.

---

## Decisiones de diseño

- **Fuente de verdad única:** el Excel central en el servidor. Los contactos **nunca se editan directamente en Outlook**.
- **Sin dependencia de nube:** todo el proceso es 100% local. Ningún dato sale de la red de la empresa.
- **Sin Exchange ni M365 empresarial:** funciona con cualquier Outlook Classic mediante COM automation nativo de Windows.
- **Privacidad:** no se usa ninguna IA ni servicio externo para procesar los datos.

---

## Notas de mantenimiento

- Revisar el log de sincronización periódicamente para detectar contactos con errores o avisos de borrado pendiente.
- Si se añaden nuevas columnas al Excel, actualizar el diccionario `MAPEO_COLUMNAS` en `01_limpiar_excel.py`.
- Si se añaden nuevas hojas al Excel, añadirlas al diccionario `HOJAS` en `01_limpiar_excel.py` indicando la fila de cabecera correcta.