"""
SCRIPT 2: SINCRONIZACIÓN EXCEL → OUTLOOK
=========================================
Autor: generado para tu empresa
Versión: 1.0

Qué hace:
- Lee el Excel limpio generado por 01_limpiar_excel.py
- Se conecta a Outlook mediante COM (nativo Windows, sin instalar nada)
- Crea/actualiza carpetas de contactos por categoría
- Añade contactos nuevos, actualiza modificados
- Registra contactos eliminados del Excel (con columna ACTIVO=NO) como aviso
- Genera log detallado de la sincronización

Requisitos:
    pip install openpyxl pandas pywin32

    pywin32 es la librería que permite controlar Outlook desde Python.
    Solo funciona en Windows con Outlook Classic instalado.

Uso:
    python 02_sincronizar_outlook.py

Configurar como tarea programada de Windows para ejecución diaria.
Ver instrucciones al final de este fichero.
"""

import pandas as pd
import win32com.client
import os
from datetime import datetime

# ─────────────────────────────────────────────
# CONFIGURACIÓN — ajusta estas rutas
# ─────────────────────────────────────────────

RUTA_EXCEL_LIMPIO = r"\\SERVIDOR\ruta\al\excel\limpio.xlsx"  # Ruta al Excel limpio generado por el paso 1
RUTA_LOG          = r"\\SERVIDOR\ruta\al\log\sincronizacion_log.txt"  # Dónde guardar el log de la sincronización

# Nombre de la carpeta raíz en Outlook donde se crearán las subcarpetas
CARPETA_RAIZ_OUTLOOK = "Contactos Empresa"

# ─────────────────────────────────────────────
# CONSTRUCCIÓN DEL CAMPO NOTES
# ─────────────────────────────────────────────

def construir_notes(fila):
    """
    Construye el contenido del campo Notes de Outlook
    con los campos que no tienen equivalente nativo.
    """
    lineas = []

    # Información de clasificación
    campos_info = [
        ("Área",         fila.get("area", "")),
        ("Tipo empresa", fila.get("tipo", "")),
        ("Año últ. trabajo", fila.get("año_ultimo_trabajo", "")),
    ]
    for etiqueta, valor in campos_info:
        if valor and str(valor).strip():
            lineas.append(f"{etiqueta}: {valor}")

    # Separador antes de acciones si hay info arriba
    acciones = []
    for i in range(1, 7):
        val = fila.get(f"accion{i}", "")
        if val and str(val).strip():
            acciones.append(f"Acción {i}: {val}")

    if acciones:
        if lineas:
            lineas.append("─" * 30)
        lineas.extend(acciones)

    # Nota general al final
    nota = fila.get("nota", "")
    if nota and str(nota).strip():
        if lineas:
            lineas.append("─" * 30)
        lineas.append(f"Nota: {nota}")

    return "\n".join(lineas)


# ─────────────────────────────────────────────
# GESTIÓN DE CARPETAS EN OUTLOOK
# ─────────────────────────────────────────────

def obtener_o_crear_carpeta(carpeta_contactos, nombre_subcarpeta):
    """
    Busca una subcarpeta dentro de la carpeta de contactos de Outlook.
    Si no existe, la crea.
    """
    for carpeta in carpeta_contactos.Folders:
        if carpeta.Name == nombre_subcarpeta:
            return carpeta
    # No existe → crear
    nueva = carpeta_contactos.Folders.Add(nombre_subcarpeta)
    return nueva


def obtener_carpeta_raiz(outlook, nombre_raiz):
    """
    Busca o crea la carpeta raíz 'Contactos Empresa' dentro de
    la carpeta de Contactos predeterminada de Outlook.
    """
    namespace = outlook.GetNamespace("MAPI")
    # Carpeta de Contactos predeterminada (olFolderContacts = 10)
    contactos_default = namespace.GetDefaultFolder(10)

    for carpeta in contactos_default.Folders:
        if carpeta.Name == nombre_raiz:
            return carpeta

    # No existe → crear
    nueva = contactos_default.Folders.Add(nombre_raiz)
    return nueva


# ─────────────────────────────────────────────
# ÍNDICE DE CONTACTOS EXISTENTES
# ─────────────────────────────────────────────

def construir_indice_contactos(carpeta):
    """
    Construye un diccionario {email_lower: ContactItem}
    de todos los contactos en una carpeta de Outlook.
    Para contactos sin email usa 'NOEMAIL_NombreApellido' como clave.
    """
    indice = {}
    items = carpeta.Items
    for i in range(items.Count):
        try:
            item = items.Item(i + 1)
            if item.Class == 40:  # 40 = olContact
                email = (item.Email1Address or "").strip().lower()
                if email:
                    indice[email] = item
                else:
                    # Clave alternativa por nombre
                    clave = f"NOEMAIL_{item.FullName.strip().lower()}"
                    indice[clave] = item
        except Exception:
            continue
    return indice


# ─────────────────────────────────────────────
# CREAR O ACTUALIZAR UN CONTACTO EN OUTLOOK
# ─────────────────────────────────────────────

def aplicar_datos_contacto(contact_item, fila, categoria):
    """
    Rellena o actualiza los campos de un ContactItem de Outlook
    con los datos de una fila del DataFrame.
    """
    # Nombre completo
    nombre    = str(fila.get("nombre", "")).strip()
    apellido1 = str(fila.get("apellido1", "")).strip()
    apellido2 = str(fila.get("apellido2", "")).strip()

    contact_item.FirstName = nombre
    # Apellidos: concatenar apellido1 y apellido2
    apellidos = " ".join(filter(None, [apellido1, apellido2]))
    contact_item.LastName  = apellidos

    # Datos de empresa y cargo
    contact_item.CompanyName = str(fila.get("empresa", "")).strip()
    contact_item.JobTitle    = str(fila.get("cargo",   "")).strip()
    contact_item.Department  = str(fila.get("area",    "")).strip()

    # Email
    email = str(fila.get("email", "")).strip()
    if email:
        contact_item.Email1Address = email

    # Teléfonos
    tlf1 = str(fila.get("TLF1", "")).strip()
    tlf2 = str(fila.get("TLF2", "")).strip()
    if tlf1:
        contact_item.BusinessTelephoneNumber = tlf1
    if tlf2:
        contact_item.MobileTelephoneNumber   = tlf2

    # Dirección
    direccion = str(fila.get("direccion", "")).strip()
    if direccion:
        contact_item.BusinessAddressStreet = direccion

    # Web
    web = str(fila.get("web", "")).strip()
    if web:
        contact_item.WebPage = web

    # País (solo hoja H)
    pais = str(fila.get("pais", "")).strip()
    if pais:
        contact_item.BusinessAddressCountry = pais

    # Categoría (nombre de la hoja/carpeta)
    contact_item.Categories = categoria

    # Notes — campos sin equivalente nativo
    contact_item.Body = construir_notes(fila)

    contact_item.Save()


# ─────────────────────────────────────────────
# SINCRONIZACIÓN PRINCIPAL
# ─────────────────────────────────────────────

def sincronizar():
    log_lineas = []
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    log_lineas.append(f"{'='*60}")
    log_lineas.append(f"LOG SINCRONIZACIÓN OUTLOOK — {timestamp}")
    log_lineas.append(f"{'='*60}\n")

    # Verificar Excel limpio
    if not os.path.exists(RUTA_EXCEL_LIMPIO):
        msg = f"ERROR: No se encuentra el Excel limpio en: {RUTA_EXCEL_LIMPIO}"
        print(msg)
        log_lineas.append(msg)
        guardar_log(log_lineas)
        return

    # Conectar a Outlook
    print("Conectando a Outlook...")
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
    except Exception as e:
        msg = f"ERROR: No se pudo conectar a Outlook. ¿Está instalado y abierto?\n{e}"
        print(msg)
        log_lineas.append(msg)
        guardar_log(log_lineas)
        return

    # Obtener carpeta raíz
    carpeta_raiz = obtener_carpeta_raiz(outlook, CARPETA_RAIZ_OUTLOOK)
    log_lineas.append(f"Carpeta raíz Outlook: '{CARPETA_RAIZ_OUTLOOK}'\n")

    # Leer todas las hojas del Excel limpio
    print("Leyendo Excel limpio...")
    excel_hojas = pd.read_excel(RUTA_EXCEL_LIMPIO, sheet_name=None)

    total_nuevos     = 0
    total_actualizados = 0
    total_omitidos   = 0
    total_avisos_borrado = 0

    for nombre_hoja, df in excel_hojas.items():
        print(f"\nSincronizando hoja '{nombre_hoja}'...")
        log_lineas.append(f"{'─'*50}")
        log_lineas.append(f"HOJA: {nombre_hoja}")
        log_lineas.append(f"{'─'*50}")

        nuevos_hoja     = 0
        actualizados_hoja = 0
        omitidos_hoja   = 0

        # Obtener o crear subcarpeta en Outlook para esta hoja
        subcarpeta = obtener_o_crear_carpeta(carpeta_raiz, nombre_hoja)

        # Construir índice de contactos ya existentes en esa carpeta
        indice_existentes = construir_indice_contactos(subcarpeta)
        emails_en_excel   = set()

        for _, fila in df.iterrows():
            # Saltar inactivos
            activo = str(fila.get("ACTIVO", "SI")).strip().upper()
            if activo != "SI":
                omitidos_hoja += 1
                continue

            nombre = str(fila.get("nombre", "")).strip()
            if not nombre:
                omitidos_hoja += 1
                continue

            email = str(fila.get("email", "")).strip().lower()
            clave = email if email else f"NOEMAIL_{nombre.lower()}"
            if email:
                emails_en_excel.add(email)

            if clave in indice_existentes:
                # ── Actualizar contacto existente ──
                try:
                    aplicar_datos_contacto(indice_existentes[clave], fila, nombre_hoja)
                    actualizados_hoja += 1
                except Exception as e:
                    log_lineas.append(f"  ❌ Error actualizando '{nombre}': {e}")
            else:
                # ── Crear contacto nuevo ──
                try:
                    nuevo_contacto = subcarpeta.Items.Add("IPM.Contact")
                    aplicar_datos_contacto(nuevo_contacto, fila, nombre_hoja)
                    nuevos_hoja += 1
                except Exception as e:
                    log_lineas.append(f"  ❌ Error creando '{nombre}': {e}")

        # ── Detectar contactos en Outlook que ya no están en el Excel ──
        for clave_existente, contact_item in indice_existentes.items():
            if clave_existente.startswith("NOEMAIL_"):
                continue  # No podemos rastrear sin email con fiabilidad
            if clave_existente not in emails_en_excel:
                total_avisos_borrado += 1
                log_lineas.append(
                    f"  ℹ️  AVISO BORRADO: '{contact_item.FullName}' ({clave_existente})\n"
                    f"      Ya no está en el Excel. Si quieres eliminarlo de Outlook,\n"
                    f"      pon ACTIVO=NO en el Excel o bórralo manualmente en Outlook."
                )

        log_lineas.append(f"  ✅ Nuevos: {nuevos_hoja} | 🔄 Actualizados: {actualizados_hoja} | ⏭️  Omitidos: {omitidos_hoja}")
        total_nuevos       += nuevos_hoja
        total_actualizados += actualizados_hoja
        total_omitidos     += omitidos_hoja

    # ── Resumen final ───────────────────────────────────────────
    log_lineas.append(f"\n{'='*60}")
    log_lineas.append("RESUMEN FINAL")
    log_lineas.append(f"{'='*60}")
    log_lineas.append(f"  ✅ Contactos nuevos añadidos:   {total_nuevos}")
    log_lineas.append(f"  🔄 Contactos actualizados:      {total_actualizados}")
    log_lineas.append(f"  ⏭️  Contactos omitidos (ACTIVO=NO o sin nombre): {total_omitidos}")
    log_lineas.append(f"  ℹ️  Avisos de borrado pendiente: {total_avisos_borrado}")
    log_lineas.append(f"\n  Sincronización completada: {timestamp}")

    guardar_log(log_lineas)
    print(f"\n✅ Sincronización completada. Log guardado en: {RUTA_LOG}")


def guardar_log(lineas):
    os.makedirs(os.path.dirname(RUTA_LOG), exist_ok=True) if os.path.dirname(RUTA_LOG) else None
    with open(RUTA_LOG, "w", encoding="utf-8") as f:
        f.write("\n".join(lineas))


if __name__ == "__main__":
    sincronizar()


# ═══════════════════════════════════════════════════════════════
# INSTRUCCIONES: CONFIGURAR TAREA PROGRAMADA EN WINDOWS
# ═══════════════════════════════════════════════════════════════
#
# Para que el script se ejecute automáticamente cada día:
#
# 1. Abre el "Programador de tareas" de Windows
#    (busca "Programador de tareas" en el menú inicio)
#
# 2. Clic en "Crear tarea básica..."
#
# 3. Nombre: "Sincronizar Contactos Outlook"
#
# 4. Desencadenador: "Diariamente" → hora: 07:45 (antes de la jornada)
#
# 5. Acción: "Iniciar un programa"
#    Programa: C:\Python312\python.exe   (ajusta tu ruta de Python)
#    Argumentos: "\\SERVIDOR\Compartido\scripts\02_sincronizar_outlook.py"
#
# 6. En "Condiciones": desmarca "Iniciar la tarea solo si el equipo
#    está conectado a corriente alterna" si son portátiles.
#
# 7. En "Configuración": marca "Ejecutar la tarea lo antes posible
#    si se perdió un inicio programado"
#
# IMPORTANTE: El script debe ejecutarse en CADA equipo de empleado,
# no solo en el servidor. Puedes distribuirlo con una GPO de Windows
# o simplemente configurar la tarea programada en cada PC manualmente.
#
# ALTERNATIVA MÁS SIMPLE: Ejecutar el script solo desde el servidor
# contra el Outlook de cada usuario mediante perfil MAPI remoto.
# Consulta con tu administrador de red si preferís esta opción.
# ═══════════════════════════════════════════════════════════════