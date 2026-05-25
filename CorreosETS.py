from datetime import datetime
from email import encoders
from email.header import Header
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from html import escape
from tkinter import *
from tkinter import filedialog, messagebox, ttk
from tkinter.scrolledtext import ScrolledText
from openpyxl import load_workbook
from urllib.parse import unquote, urlparse
import mimetypes
import os
import queue
import re
import sys
import threading
import time

import boto3

# ========================
# CONFIGURACION DE BLOQUES
# ========================

BLOQUE_TAMANO = 50
ESPERA_ENTRE_BLOQUES = 60  # segundos
SES_MAX_MESSAGE_BYTES = 10 * 1024 * 1024

BENEFICIOS = [
    "1 TRASLADO EN AMBULANCIA GRATUITO AL AÑO EN CASO DE URGENCIA REAL",
    "ASESORIA MEDICA TELEFONICA Y POR VIDEOCONFERENCIA GRATUITA",
    "CONSULTA A DOMICILIO CON MEDICOS GENRALES DE NUESTRA RED A PRECIO PREFERENCIAL. ES INDISPENSABLE AGENDAR PREVIA CITA EN NUESTRO SERVICIO DE CONCIERGE",
    "COBERTURA INDIVIDUAL",
    "AMBULANCIA CON PRECIO PREFERENCIAL",
    "ASESORIA NUTRICIONAL TELEFONICA Y POR VIDEOCONFERENCIA",
    "ASESORIA EMOCIONAL TELEFONICA Y POR VIDEOCONFERENCIA",
    "RED DENTAL CON DESCUENTOS A NIVEL NACIONAL",
    "RED VISUAL CON DESCUENTOS A NIVEL NACIONAL",
    "DESCUENTOS EN ESTABLECIMENTOS COMERCIALES RED TDCONSENTIDO A NIVEL NACIONAL",
    "SERVICIO FUNERARIO. APLICA TIEMPO DE ESPERA DE 90 DIAS A PARTIR DE LA CONTRATACION DE LA MEMBRESIA A EXCEPCION DE MUERTE ACCIDENTAL LA CUAL APLICA DESDE EL DIA 1. INHUMACION O CREMACION. RECOLECCION DEL CUERPO, TRASLADO NIV NACIONAL, ARREGLO ESTETICO, SALA DE VELACION 30 PERSONAS EN CIRCULACION O DOMICILIO SIN COSTO ADICIONAL, SERVICIO DE TANATOLOGIA, ASESORIA JURIDICA TESTAMENTARIA VIA TEL. PARA INHUMACION SERV DE EMBALSAMADO, ATAUD METALICO, TRASLADO EN CARROZA.PARA CREMACION URNA BASICA Y SERV DE CREMACION",
    "CUPONES CON PROMOCIONES ESPECIALES PARA APP MOVIL.",
    "FARMACIA EN LINEA CON ENVIO A DOMICILIO",
    "SEGURO DE ACCIDENTES PERSONALES. MUERTE ACCIDENTAL $200,000 APLICA DE 12 A 70 AÑOS PERDIDA DE MIEMBROS POR ACCIDENTE ESCALA B $30,000 REEMBOLSO DE GASTOS FUNERARIOS POR ACCIDENTE $30,000 REEMBOLSO DE GASTOS MEDICOS POR ACCIDENTE $20,000 APLICAN DE 0 A 70 AÑOS. COBERTURAS A NIVEL NACIONAL. LAS MEMBRESIAS NO SON ACUMULABLES. POLIZA DE SEGURO 2510030074730",
    "PLATAFORMA CON EXPEDIENTE CLINICO ELECTRÓNICO.",
    "1 CHECK UP GRATIS PARA TITULAR. INCLUYE QS6, BIOMETRIA HEMATICA, EXAMEN GENERAL DE ORINA, INTERPRETACION DE RESULTADOS POR NUESTRO CESDI.",
    "1 CONSULTA PRESENCIAL O VIDEOCONSULTA CON MEDICINA GENERAL, NUTRICION, PSICOLOGIA, COORDINAR Y AGENDAR A TRAVES DE NUESTRO CONCIERGE. CONSULTA CON MEDICOS DE LA RED A PRECIO TABULADO.",
    "2 VIDEOCONSULTAS DE PEDIATRIA O GINECOLOGIA O MEDICINA INTERNA",
    "3 CUPONES 2 X 1 PARA CINE, SOLICITAR EL BENEFICIO DEBERA COMUNICARSE A CONCIERGE, SE VERIFICARA QUE EL USUARIO SE ENCUENTRE ACTIVO AL MOMENTO DE LA SOLICITUD Y DE SER EL CASO, SE PROPORCIONARA UN CODIGO 2X1. CONSULTA RESTRICCIONES.",
    "1 VIDEOCONSULTA O CONSULTA PRESENCIAL DE ESPECIALIDAD DENTRO DE NUESTRA RED MEDICA SIN COSTO. PARA UTILIZAR ESTE BENEFICIO, ES INDISPENSABLE COORDINAR Y AGENDAR LA CITA A TRAVES DE NUESTRO CONCIERGE MEDICO.",
    "PROTOCOLO DE TRIAGE MEDICO PARA DETERMINAR NIVEL DE URGENCIA, DIAGNOSTICO Y ATENCION",
    "VIDEOCONSULTAS ILIMITADAS CON MEDICINA GENERAL, NUTRICION Y PSICOLOGIA CON MEDICOS DEL CENTRO DE CONTACTO DE SALUD DIGITAL.",
]

# ========================
# VARIABLES GLOBALES
# ========================

ruta_excel = ""
rutas_pdf = []
ruta_imagen_1 = ""
ruta_imagen_2 = ""
ruta_imagen_3 = ""
cola_envio = queue.Queue()
envio_en_proceso = False

TEXTO_CORREO = (
    "Hola 👋\n\n"
    "Nos emociona darte la bienvenida a tu plataforma de reembolso.\n"
    "Aquí podrás consultar tus beneficios, iniciar tus solicitudes y dar seguimiento a cada paso de forma clara, rápida y segura.\n\n"
    "✨ Descubre cómo funciona y comienza a usarla hoy mismo.\n"
    "Solo elige tu documento y ábrelo"
)


def cargar_env(ruta=".env"):
    rutas = [
        os.path.join(os.getcwd(), ruta),
        os.path.join(os.path.dirname(os.path.abspath(__file__)), ruta),
    ]

    for ruta_env in dict.fromkeys(rutas):
        if not os.path.exists(ruta_env):
            continue

        with open(ruta_env, "r", encoding="utf-8") as archivo:
            for linea in archivo:
                linea = linea.strip()
                if not linea or linea.startswith("#") or "=" not in linea:
                    continue

                clave, valor = linea.split("=", 1)
                clave = clave.strip()
                valor = valor.strip().strip('"').strip("'")
                os.environ.setdefault(clave, valor)


cargar_env()

AWS_REGION = os.getenv("AWS_REGION", "")
AWS_ACCESS_KEY = os.getenv("AWS_ACCESS_KEY", "")
AWS_SECRET_KEY = os.getenv("AWS_SECRET_KEY", "")
SES_SOURCE_EMAIL = os.getenv("SES_SOURCE_EMAIL", "")


def limpiar_valor(valor):
    if valor is None:
        return ""

    texto = str(valor).strip()
    if texto.startswith("'"):
        texto = texto[1:].strip()

    return texto


def normalizar_header(valor):
    return limpiar_valor(valor).lower().replace(" ", "")


def validar_columnas(headers):
    columnas_requeridas = [
        "nombre",
        "correo",
    ]

    faltantes = [col for col in columnas_requeridas if col not in headers]
    if faltantes:
        raise ValueError("Faltan columnas en el Excel: " + ", ".join(faltantes))

    return headers


def construir_nombre_completo(origen, columnas):
    partes = []
    for campo in ("nombre", "paterno", "materno"):
        indice = columnas.get(campo)
        if indice is None:
            continue

        if isinstance(origen, (list, tuple)) and indice < len(origen):
            valor = limpiar_valor(origen[indice])
            if valor:
                partes.append(valor)

    return " ".join(partes).strip()


def obtener_indice_columnas(ws):
    headers = {}

    for i, celda in enumerate(ws[1], start=0):
        header = normalizar_header(celda.value)
        if header:
            headers[header] = i

    return validar_columnas(headers)


def obtener_indice_columnas_xls(header_row):
    headers = {}

    for i, valor in enumerate(header_row):
        header = normalizar_header(valor)
        if header:
            headers[header] = i

    return validar_columnas(headers)


def leer_destinatarios(ruta):
    extension = os.path.splitext(ruta)[1].lower()
    if extension == ".xls":
        return leer_destinatarios_xls(ruta)

    wb = load_workbook(ruta, data_only=True)
    ws = wb.active
    columnas = obtener_indice_columnas(ws)
    destinatarios = []
    descartados_vacios = 0
    descartados_invalidos = 0

    for row in ws.iter_rows(min_row=2, values_only=True):
        nombre = construir_nombre_completo(row, columnas)
        email = limpiar_valor(row[columnas["correo"]]).lower()

        if not email:
            descartados_vacios += 1
            continue

        if "@" not in email:
            descartados_invalidos += 1
            continue

        if email and "@" in email:
            destinatarios.append({
                "nombre": nombre,
                "email": email,
            })

    return destinatarios, descartados_vacios, descartados_invalidos


def leer_destinatarios_xls(ruta):
    import xlrd

    workbook = xlrd.open_workbook(ruta)
    sheet = workbook.sheet_by_index(0)
    columnas = obtener_indice_columnas_xls(sheet.row_values(0))
    destinatarios = []
    descartados_vacios = 0
    descartados_invalidos = 0

    for row_idx in range(1, sheet.nrows):
        row = sheet.row_values(row_idx)
        nombre = construir_nombre_completo(row, columnas)
        email = limpiar_valor(row[columnas["correo"]]).lower()

        if not email:
            descartados_vacios += 1
            continue

        if "@" not in email:
            descartados_invalidos += 1
            continue

        if email and "@" in email:
            destinatarios.append({
                "nombre": nombre,
                "email": email,
            })

    return destinatarios, descartados_vacios, descartados_invalidos


def beneficios_html():
    items = "".join(f"<li>{escape(beneficio)}</li>" for beneficio in BENEFICIOS)
    return (
        '<ul style="margin:0; padding-left:18px; '
        "font-family: Verdana, Tahoma, Arial, sans-serif; "
        'font-size:14px; line-height:1.4;">'
        f"{items}</ul>"
    )


def imagen_frente_html(valor):
    valor = limpiar_valor(valor)
    if not valor:
        return ""

    if valor.startswith("<"):
        return valor

    return f'<img src="{escape(valor, quote=True)}" alt="" style="max-width:180px; height:auto;">'


def personalizar_html(html_base, persona):
    nombre = escape(persona.get("nombre", "").strip())
    saludo = f"<p>{nombre}</p>" if nombre else ""
    html_personalizado = f"{saludo}{html_base}"
    reemplazos = {
        "{{NOMBRE}}": nombre,
        "{$nomconcatenado}": nombre,
    }

    for placeholder, valor in reemplazos.items():
        html_personalizado = html_personalizado.replace(placeholder, valor)

    return html_personalizado


def es_recurso_externo(src):
    src = src.strip()
    esquema = urlparse(src).scheme.lower()
    return esquema in ("http", "https", "cid", "data", "mailto", "tel")


def ruta_imagen_local(carpeta_html, src):
    src_limpio = unquote(src.split("#", 1)[0].split("?", 1)[0]).replace("/", os.sep)
    ruta = os.path.abspath(os.path.join(carpeta_html, src_limpio))
    carpeta_base = os.path.abspath(carpeta_html)

    if not ruta.startswith(carpeta_base):
        return None

    if not os.path.isfile(ruta):
        return None

    return ruta


def embeber_imagenes_html(html, carpeta_html):
    imagenes = {}

    def reemplazar_src(match):
        inicio, src, fin = match.groups()
        if es_recurso_externo(src):
            return match.group(0)

        ruta = ruta_imagen_local(carpeta_html, src)
        if not ruta:
            return match.group(0)

        cid = os.path.basename(ruta)
        imagenes[cid] = ruta
        return f"{inicio}cid:{cid}{fin}"

    html_con_cid = re.sub(
        r'(<img\b[^>]*?\bsrc=["\'])([^"\']+)(["\'])',
        reemplazar_src,
        html,
        flags=re.IGNORECASE,
    )

    return html_con_cid, imagenes


def crear_adjunto_imagen(cid, ruta):
    content_type, _ = mimetypes.guess_type(ruta)
    if not content_type:
        content_type = "application/octet-stream"

    maintype, subtype = content_type.split("/", 1)

    with open(ruta, "rb") as f:
        adjunto = MIMEBase(maintype, subtype)
        adjunto.set_payload(f.read())

    encoders.encode_base64(adjunto)
    adjunto.add_header("Content-ID", f"<{cid}>")
    adjunto.add_header("Content-Disposition", "inline", filename=os.path.basename(ruta))

    return adjunto


def construir_mensaje_raw(
    asunto,
    remitente,
    destinatario,
    nombre_destinatario,
    ruta_img_1,
    ruta_img_2,
    ruta_img_3,
    rutas_adjuntos_pdf,
):
    mensaje = MIMEMultipart("mixed")
    mensaje["Subject"] = Header(asunto, "utf-8")
    mensaje["From"] = remitente
    mensaje["To"] = destinatario

    html_contenido = f"""
<html>
  <body style="margin:0; padding:0; background-color:#ffffff;">
    <table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" style="background-color:#ffffff;">
      <tr>
        <td align="center" style="padding:20px 12px;">
          <table role="presentation" cellpadding="0" cellspacing="0" border="0" width="600" style="width:600px; max-width:600px; background-color:#ffffff; font-family:Arial, sans-serif; color:#111111;">
            <tr>
              <td align="center" style="padding-bottom:18px;">
                <img src="cid:imagen_1" alt="Imagen 1" style="max-width:100%; height:auto; display:block; border:0;">
              </td>
            </tr>
            <tr>
              <td style="padding-bottom:10px;">
                <table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%">
                  <tr>
                    <td valign="top" width="180" style="padding-right:18px;">
                      <img src="cid:imagen_2" alt="Imagen 2" style="max-width:180px; width:100%; height:auto; display:block; border:0;">
                    </td>
                    <td valign="top" style="font-size:16px; line-height:1.35; color:#111111;">
                      <p style="margin:0 0 14px 0;">Hola 👋</p>
                      <p style="margin:0;">Nos emociona darte la bienvenida a tu plataforma de reembolso.</p>
                    </td>
                  </tr>
                </table>
              </td>
            </tr>
            <tr>
              <td style="font-size:16px; line-height:1.35; color:#111111; padding:8px 0 10px 0;">
                <p style="margin:0 0 14px 0;">Aquí podrás consultar tus beneficios, iniciar tus solicitudes y dar seguimiento a cada paso de forma clara, rápida y segura.</p>
                <p style="margin:0;">✨ Descubre cómo funciona y comienza a usarla hoy mismo.<br>Solo elige tu documento y ábrelo.</p>
              </td>
            </tr>
            <tr>
              <td align="center" style="padding-top:18px;">
                <img src="cid:imagen_3" alt="Imagen 3" style="max-width:100%; height:auto; display:block; border:0;">
              </td>
            </tr>
          </table>
        </td>
      </tr>
    </table>
  </body>
</html>
"""

    related = MIMEMultipart("related")
    alternative = MIMEMultipart("alternative")
    alternative.attach(MIMEText(TEXTO_CORREO, "plain", "utf-8"))
    alternative.attach(MIMEText(html_contenido, "html", "utf-8"))
    related.attach(alternative)
    related.attach(crear_adjunto_imagen("imagen_1", ruta_img_1))
    related.attach(crear_adjunto_imagen("imagen_2", ruta_img_2))
    related.attach(crear_adjunto_imagen("imagen_3", ruta_img_3))
    mensaje.attach(related)

    for ruta_adjunto in rutas_adjuntos_pdf:
        content_type, _ = mimetypes.guess_type(ruta_adjunto)
        if not content_type:
            content_type = "application/octet-stream"

        maintype, subtype = content_type.split("/", 1)
        with open(ruta_adjunto, "rb") as f:
            adjunto = MIMEBase(maintype, subtype)
            adjunto.set_payload(f.read())

        encoders.encode_base64(adjunto)
        adjunto.add_header(
            "Content-Disposition",
            "attachment",
            filename=os.path.basename(ruta_adjunto),
        )
        mensaje.attach(adjunto)

    return mensaje


def validar_tamano_mensaje(asunto, remitente, ruta_img_1, ruta_img_2, ruta_img_3, rutas_adjuntos_pdf):
    mensaje_prueba = construir_mensaje_raw(
        asunto,
        remitente,
        "destinatario@ejemplo.com",
        "Destinatario",
        ruta_img_1,
        ruta_img_2,
        ruta_img_3,
        rutas_adjuntos_pdf,
    )
    tamano = len(mensaje_prueba.as_bytes())

    if tamano > SES_MAX_MESSAGE_BYTES:
        raise ValueError(
            "El correo con adjuntos excede el limite de AWS SES (10 MB). "
            f"Tamaño actual: {tamano:,} bytes. "
            "Reduce la cantidad/tamaño de PDFs y vuelve a intentar."
        )

    return tamano


def crear_cliente_ses():
    faltantes = [
        nombre
        for nombre, valor in {
            "AWS_REGION": AWS_REGION,
            "AWS_ACCESS_KEY": AWS_ACCESS_KEY,
            "AWS_SECRET_KEY": AWS_SECRET_KEY,
        }.items()
        if not valor
    ]

    if faltantes:
        raise ValueError("Faltan variables en .env: " + ", ".join(faltantes))

    return boto3.client(
        "ses",
        region_name=AWS_REGION,
        aws_access_key_id=AWS_ACCESS_KEY,
        aws_secret_access_key=AWS_SECRET_KEY,
    )


def dominio_de_correo(correo):
    if "@" not in correo:
        return ""

    return correo.rsplit("@", 1)[1].strip().lower()


def obtener_identidades_verificadas(ses):
    identidades = []

    for tipo in ("EmailAddress", "Domain"):
        token = None
        while True:
            args = {
                "IdentityType": tipo,
                "MaxItems": 1000,
            }
            if token:
                args["NextToken"] = token

            response = ses.list_identities(**args)
            identidades.extend(response.get("Identities", []))
            token = response.get("NextToken")

            if not token:
                break

    verificadas = []
    for inicio in range(0, len(identidades), 100):
        bloque = identidades[inicio:inicio + 100]
        atributos = ses.get_identity_verification_attributes(
            Identities=bloque
        ).get("VerificationAttributes", {})

        for identidad in bloque:
            status = atributos.get(identidad, {}).get("VerificationStatus")
            if status == "Success":
                verificadas.append(identidad.lower())

    return verificadas


def validar_remitente_ses(ses, remitente):
    correo = remitente.strip().lower()
    dominio = dominio_de_correo(correo)
    verificadas = obtener_identidades_verificadas(ses)

    if correo in verificadas or dominio in verificadas:
        return

    opciones = "\n".join(f"- {identidad}" for identidad in verificadas[:20])
    if len(verificadas) > 20:
        opciones += f"\n- ... y {len(verificadas) - 20} mas"

    raise ValueError(
        f"El remitente no esta verificado en SES region {AWS_REGION}: {remitente}\n\n"
        "Usa uno de estos remitentes/dominios verificados:\n"
        f"{opciones or '- No se encontraron identidades verificadas'}"
    )


def escribir_log_ui(texto, tag):
    log_text.configure(state=NORMAL)
    log_text.insert(END, texto + "\n", tag)
    log_text.see(END)
    log_text.configure(state=DISABLED)
    principal.update_idletasks()


def actualizar_ui(actual, total):
    progress["value"] = actual
    porcentaje = int((actual / total) * 100) if total else 0
    label_contador.config(text=f"{actual} / {total} correos")
    label_porcentaje.config(text=f"{porcentaje}%")
    principal.update_idletasks()


def limpiar_log_ui():
    log_text.configure(state=NORMAL)
    log_text.delete("1.0", END)
    log_text.configure(state=DISABLED)


def procesar_cola_envio():
    global envio_en_proceso

    while True:
        try:
            evento = cola_envio.get_nowait()
        except queue.Empty:
            break

        tipo = evento.get("tipo")

        if tipo == "log":
            escribir_log_ui(evento["texto"], evento["tag"])
        elif tipo == "progress":
            actualizar_ui(evento["actual"], evento["total"])
        elif tipo == "error":
            envio_en_proceso = False
            btn_enviar.config(state=NORMAL)
            messagebox.showerror(evento["titulo"], evento["mensaje"])
        elif tipo == "done":
            envio_en_proceso = False
            btn_enviar.config(state=NORMAL)
            messagebox.showinfo("Proceso terminado", evento["mensaje"])

    if envio_en_proceso:
        principal.after(100, procesar_cola_envio)


def enviar_worker(
    remitente,
    asunto,
    ruta_img_1,
    ruta_img_2,
    ruta_img_3,
    rutas_adjuntos_pdf,
    destinatarios,
    descartados_vacios,
    descartados_invalidos,
):
    try:
        ses = crear_cliente_ses()
        validar_remitente_ses(ses, remitente)
    except Exception as e:
        cola_envio.put({
            "tipo": "error",
            "titulo": "AWS SES Error",
            "mensaje": str(e),
        })
        return

    enviados = 0
    total = len(destinatarios)
    nombre_log = f"log_envio_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"

    try:
        with open(nombre_log, "w", encoding="utf-8") as log_file:
            log_file.write("===== LOG DE ENVIO AWS SES =====\n")
            log_file.write(f"Fecha: {datetime.now()}\n")
            log_file.write(f"Remitente: {remitente}\n")
            log_file.write(f"Total destinatarios: {total}\n\n")
            log_file.write(f"Descartados por email vacio: {descartados_vacios}\n")
            log_file.write(f"Descartados por email invalido: {descartados_invalidos}\n\n")

            for i, persona in enumerate(destinatarios, start=1):
                try:
                    mensaje = construir_mensaje_raw(
                        asunto,
                        remitente,
                        persona["email"],
                        persona["nombre"] or "afiliado",
                        ruta_img_1,
                        ruta_img_2,
                        ruta_img_3,
                        rutas_adjuntos_pdf,
                    )

                    response = ses.send_raw_email(
                        Source=remitente,
                        Destinations=[persona["email"]],
                        RawMessage={
                            "Data": mensaje.as_string(),
                        }
                    )

                    enviados += 1
                    message_id = response.get("MessageId", "Sin MessageId")
                    linea = f"[OK] {persona['email']} - {message_id}"
                    log_file.write(linea + "\n")
                    cola_envio.put({"tipo": "log", "texto": linea, "tag": "ok"})

                except Exception as e:
                    linea = f"[ERROR] {persona['email']} -> {e}"
                    log_file.write(linea + "\n")
                    cola_envio.put({"tipo": "log", "texto": linea, "tag": "error"})

                cola_envio.put({"tipo": "progress", "actual": i, "total": total})
                time.sleep(0.1)

                if i % BLOQUE_TAMANO == 0 and i < total:
                    mensaje_espera = f"Esperando {ESPERA_ENTRE_BLOQUES} segundos despues de {i} envios..."
                    log_file.write(f"\n--- {mensaje_espera} ---\n\n")
                    cola_envio.put({"tipo": "log", "texto": mensaje_espera, "tag": "info"})
                    time.sleep(ESPERA_ENTRE_BLOQUES)

            log_file.write("\n===== RESUMEN =====\n")
            log_file.write(f"Enviados correctamente: {enviados}\n")
            log_file.write(f"No enviados: {total - enviados}\n")
            log_file.write(f"Descartados por email vacio: {descartados_vacios}\n")
            log_file.write(f"Descartados por email invalido: {descartados_invalidos}\n")

    except Exception as e:
        cola_envio.put({
            "tipo": "error",
            "titulo": "Error",
            "mensaje": f"Error durante el envio:\n{e}",
        })
        return

    cola_envio.put({
        "tipo": "done",
        "mensaje": (
            f"{enviados} de {total} correos enviados.\n"
            f"Descartados vacios: {descartados_vacios} | invalidos: {descartados_invalidos}\n"
            f"Log generado: {nombre_log}"
        ),
    })


# ========================
# INTERFAZ
# ========================

principal = Tk()
principal.title("Mailing Salud Interactiva")

if hasattr(sys, "_MEIPASS"):
    base_path = sys._MEIPASS
else:
    base_path = os.path.abspath(".")

icono = os.path.join(base_path, "si-logo.ico")
if os.path.exists(icono):
    principal.iconbitmap(icono)

principal.geometry("760x560")
principal.configure(bg="black")
principal.grid_columnconfigure(1, weight=1)

Label(principal, text="Remitente SES:", bg="black", fg="white").grid(row=0, column=0, padx=10, pady=5, sticky="e")
remitente_var = StringVar(value=SES_SOURCE_EMAIL)
Entry(principal, textvariable=remitente_var, width=70).grid(row=0, column=1, padx=5, pady=5, sticky="ew")

Label(principal, text="Asunto:", bg="black", fg="white").grid(row=1, column=0, padx=10, pady=5, sticky="e")
asunto_var = StringVar()
Entry(principal, textvariable=asunto_var, width=70).grid(row=1, column=1, padx=5, pady=5, sticky="ew")


# ========================
# SELECCION DE ARCHIVOS
# ========================

def seleccionar_excel():
    global ruta_excel
    ruta_excel = filedialog.askopenfilename(
        title="Seleccionar Excel",
        filetypes=[("Excel files", "*.xlsx *.xlsm *.xls")],
    )
    if ruta_excel:
        label_excel.config(text=os.path.basename(ruta_excel))


def seleccionar_pdf():
    global rutas_pdf
    rutas_pdf = list(filedialog.askopenfilenames(
        title="Seleccionar PDFs",
        filetypes=[("PDF", "*.pdf")],
    ))
    if rutas_pdf:
        label_pdf.config(text=f"{len(rutas_pdf)} archivo(s) seleccionado(s)")
    else:
        label_pdf.config(text="Ningun archivo seleccionado")


def seleccionar_imagen_1():
    global ruta_imagen_1
    ruta_imagen_1 = filedialog.askopenfilename(
        title="Seleccionar Imagen 1",
        filetypes=[("Imagenes", "*.png *.jpg *.jpeg *.webp")],
    )
    if ruta_imagen_1:
        label_imagen_1.config(text=os.path.basename(ruta_imagen_1))


def seleccionar_imagen_2():
    global ruta_imagen_2
    ruta_imagen_2 = filedialog.askopenfilename(
        title="Seleccionar Imagen 2",
        filetypes=[("Imagenes", "*.png *.jpg *.jpeg *.webp")],
    )
    if ruta_imagen_2:
        label_imagen_2.config(text=os.path.basename(ruta_imagen_2))


def seleccionar_imagen_3():
    global ruta_imagen_3
    ruta_imagen_3 = filedialog.askopenfilename(
        title="Seleccionar Imagen 3",
        filetypes=[("Imagenes", "*.png *.jpg *.jpeg *.webp")],
    )
    if ruta_imagen_3:
        label_imagen_3.config(text=os.path.basename(ruta_imagen_3))


# ========================
# FUNCION PRINCIPAL
# ========================

def enviar():
    global ruta_excel, ruta_imagen_1, ruta_imagen_2, ruta_imagen_3, rutas_pdf, envio_en_proceso

    if envio_en_proceso:
        messagebox.showinfo("Envio en proceso", "Ya hay un envio en curso. Espera a que termine.")
        return

    remitente = remitente_var.get().strip()
    asunto = asunto_var.get().strip()

    if not remitente or "@" not in remitente:
        messagebox.showerror("Error", "Escribe un correo remitente valido")
        return

    if not asunto:
        messagebox.showerror("Error", "Escribe el asunto del correo")
        return

    if not ruta_excel or not ruta_imagen_1 or not ruta_imagen_2 or not ruta_imagen_3:
        messagebox.showerror("Error", "Selecciona Excel e Imagen 1, Imagen 2 e Imagen 3")
        return

    try:
        validar_tamano_mensaje(
            asunto,
            remitente,
            ruta_imagen_1,
            ruta_imagen_2,
            ruta_imagen_3,
            rutas_pdf,
        )
    except Exception as e:
        messagebox.showerror("Error", str(e))
        return

    try:
        destinatarios, descartados_vacios, descartados_invalidos = leer_destinatarios(ruta_excel)
    except Exception as e:
        messagebox.showerror("Error", f"Excel invalido:\n{e}")
        return

    if not destinatarios:
        messagebox.showerror("Error", "No se encontraron correos en la columna correo")
        return

    total = len(destinatarios)

    progress["maximum"] = total
    progress["value"] = 0
    actualizar_ui(0, total)
    limpiar_log_ui()
    envio_en_proceso = True
    btn_enviar.config(state=DISABLED)

    worker = threading.Thread(
        target=enviar_worker,
        args=(
            remitente,
            asunto,
            ruta_imagen_1,
            ruta_imagen_2,
            ruta_imagen_3,
            rutas_pdf,
            destinatarios,
            descartados_vacios,
            descartados_invalidos,
        ),
        daemon=True,
    )
    worker.start()
    principal.after(100, procesar_cola_envio)


# ========================
# BOTONES
# ========================

Button(principal, text="Seleccionar Excel", command=seleccionar_excel).grid(row=2, column=0, padx=10, pady=5, sticky="ew")

label_excel = Label(principal, text="Ningun archivo seleccionado", bg="black", fg="white", anchor="w")
label_excel.grid(row=2, column=1, padx=5, pady=5, sticky="ew")

Button(principal, text="Seleccionar Imagen 1", command=seleccionar_imagen_1).grid(row=3, column=0, padx=10, pady=5, sticky="ew")

label_imagen_1 = Label(principal, text="Ningun archivo seleccionado", bg="black", fg="white", anchor="w")
label_imagen_1.grid(row=3, column=1, padx=5, pady=5, sticky="ew")

Button(principal, text="Seleccionar Imagen 2", command=seleccionar_imagen_2).grid(row=4, column=0, padx=10, pady=5, sticky="ew")

label_imagen_2 = Label(principal, text="Ningun archivo seleccionado", bg="black", fg="white", anchor="w")
label_imagen_2.grid(row=4, column=1, padx=5, pady=5, sticky="ew")

Button(principal, text="Seleccionar Imagen 3", command=seleccionar_imagen_3).grid(row=5, column=0, padx=10, pady=5, sticky="ew")

label_imagen_3 = Label(principal, text="Ningun archivo seleccionado", bg="black", fg="white", anchor="w")
label_imagen_3.grid(row=5, column=1, padx=5, pady=5, sticky="ew")

Button(principal, text="Seleccionar PDFs (Opcional)", command=seleccionar_pdf).grid(row=6, column=0, padx=10, pady=5, sticky="ew")

label_pdf = Label(principal, text="Ningun archivo seleccionado", bg="black", fg="white", anchor="w")
label_pdf.grid(row=6, column=1, padx=5, pady=5, sticky="ew")

progress = ttk.Progressbar(principal, orient="horizontal", length=500, mode="determinate")
progress.grid(row=7, column=1, padx=5, pady=15, sticky="ew")

label_contador = Label(principal, text="0 / 0 correos", bg="black", fg="white")
label_contador.grid(row=8, column=1, padx=5, sticky="w")

label_porcentaje = Label(principal, text="0%", bg="black", fg="white")
label_porcentaje.grid(row=8, column=1, padx=5, sticky="e")

btn_enviar = Button(
    principal,
    text="Enviar Correos",
    command=enviar,
    bg="#215cc1",
    fg="white",
    width=30,
)
btn_enviar.grid(row=9, column=1, padx=5, pady=10)

Label(principal, text="Log en tiempo real:", bg="black", fg="white").grid(row=10, column=0, padx=10, pady=5, sticky="ne")

log_text = ScrolledText(principal, height=14, width=80, state=DISABLED, bg="#111111", fg="white")
log_text.grid(row=10, column=1, padx=5, pady=5, sticky="nsew")
log_text.tag_configure("ok", foreground="#33cc66")
log_text.tag_configure("error", foreground="#ff5555")
log_text.tag_configure("info", foreground="#d7d7d7")

principal.grid_rowconfigure(10, weight=1)

principal.mainloop()
