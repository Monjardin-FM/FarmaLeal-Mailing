# principal.mainloop()
from datetime import datetime
from tkinter import *
from tkinter import filedialog, messagebox, ttk
from openpyxl import load_workbook
import smtplib
import os
import time
import sys
import os
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage


# ========================
# VARIABLES GLOBALES
# ========================

ruta_excel = ""
ruta_html = ""


# ========================
# INTERFAZ
# ========================

principal = Tk()
principal.title("Mailing Salud Interactiva")
if hasattr(sys, '_MEIPASS'):
    base_path = sys._MEIPASS
else:
    base_path = os.path.abspath(".")

principal.iconbitmap(os.path.join(base_path, "si-logo.ico"))
principal.geometry("600x350")
principal.configure(bg="black")

Label(principal, text="Correo:", bg="black", fg="white").grid(row=0, column=0, pady=5, sticky="e")

correo_var = StringVar()
Entry(principal, textvariable=correo_var, width=50).grid(row=0, column=1)


Label(principal, text="Contraseña:", bg="black", fg="white").grid(row=1, column=0, pady=5, sticky="e")

pass_var = StringVar()
Entry(principal, textvariable=pass_var, show="*", width=50).grid(row=1, column=1)


Label(principal, text="Asunto:", bg="black", fg="white").grid(row=2, column=0, pady=5, sticky="e")

asunto_var = StringVar()
Entry(principal, textvariable=asunto_var, width=50).grid(row=2, column=1)


# ========================
# FUNCIONES PARA SELECCIÓN
# ========================

def seleccionar_excel():
    global ruta_excel
    ruta_excel = filedialog.askopenfilename(
        title="Seleccionar Excel",
        filetypes=[("Excel files", "*.xlsx")]
    )

    if ruta_excel:
        label_excel.config(text=os.path.basename(ruta_excel))


def seleccionar_html():
    global ruta_html
    ruta_html = filedialog.askopenfilename(
        title="Seleccionar HTML",
        filetypes=[("HTML files", "*.html")]
    )

    if ruta_html:
        label_html.config(text=os.path.basename(ruta_html))


# ========================
# FUNCION PRINCIPAL
# ========================

def enviar():

    global ruta_excel, ruta_html

    correo = correo_var.get()
    contraseña = pass_var.get()
    asunto = asunto_var.get()

    if not correo or not contraseña or not asunto:
        messagebox.showerror("Error", "Completa todos los campos")
        return

    if not ruta_excel or not ruta_html:
        messagebox.showerror("Error", "Selecciona Excel y HTML")
        return

    carpeta_material = os.path.dirname(ruta_html)

    # Leer HTML
    with open(ruta_html, "r", encoding="utf-8") as f:
        html_base = f.read()

    # Leer Excel
    try:
        wb = load_workbook(ruta_excel)
        ws = wb.active
    except Exception as e:
        messagebox.showerror("Error", f"Excel inválido:\n{e}")
        return

    # Obtener destinatarios
    destinatarios = []

    for row in ws.iter_rows(min_row=2, values_only=True):

        empleado = row[0]
        nombre = row[1]
        vigencia = row[2]
        producto = row[3]
        email = row[4]

        if email and "@" in email:
            destinatarios.append({
                "empleado": empleado,
                "nombre": nombre,
                "vigencia": vigencia,
                "producto": producto,
                "email": email
            })

    if not destinatarios:
        messagebox.showerror("Error", "No se encontraron correos")
        return

    # Conectar SMTP (Rackspace)
    try:
        server = smtplib.SMTP("secure.emailsrvr.com", 587, timeout=30)
        server.ehlo()
        server.starttls()
        server.ehlo()
        server.login(correo.strip(), contraseña.strip())
    except Exception as e:
        messagebox.showerror("SMTP Error", str(e))
        return

    enviados = 0

    # Configurar barra de progreso
    # progress["maximum"] = len(destinatarios)
    # progress["value"] = 0
    # principal.update_idletasks()
    total = len(destinatarios)
    SEGUNDOS_TOTALES = 11 * 60 * 60  # 12 horas
    intervalo_envio = SEGUNDOS_TOTALES / total
    progress["maximum"] = total
    progress["value"] = 0

    label_contador.config(text=f"0 / {total} enviados")
    label_porcentaje.config(text="0%")

    principal.update_idletasks()

    # Crear archivo de log
    fecha_actual = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    nombre_log = f"log_envio_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"

    log_file = open(nombre_log, "w", encoding="utf-8")
    log_file.write("===== LOG DE ENVÍO =====\n")
    log_file.write(f"Fecha: {fecha_actual}\n")
    log_file.write(f"Total destinatarios: {total}\n\n")
    # Enviar correos
    for persona in destinatarios:

        mensaje = MIMEMultipart("related")
        mensaje["Subject"] = asunto
        mensaje["From"] = correo
        mensaje["To"] = persona["email"]

        alt = MIMEMultipart("alternative")
        mensaje.attach(alt)

        html_personalizado = html_base
        html_personalizado = html_personalizado.replace("{{NOMBRE}}", str(persona["nombre"]))
        html_personalizado = html_personalizado.replace("{{EMPLEADO}}", str(persona["empleado"]))
        html_personalizado = html_personalizado.replace("{{VIGENCIA}}", str(persona["vigencia"]))
        html_personalizado = html_personalizado.replace("{{PRODUCTO}}", str(persona["producto"]))

        alt.attach(MIMEText(html_personalizado, "html", "utf-8"))

        # Adjuntar imágenes
        for archivo in os.listdir(carpeta_material):

            if archivo.startswith("._"):
                continue

            if archivo.lower().endswith((".jpg", ".jpeg", ".png")):

                ruta_img = os.path.join(carpeta_material, archivo)

                try:
                    with open(ruta_img, "rb") as f:
                        img = MIMEImage(f.read())
                        img.add_header("Content-ID", f"<{archivo}>")
                        img.add_header("Content-Disposition", "inline", filename=archivo)
                        mensaje.attach(img)
                except Exception as e:
                    print("Error imagen:", archivo, e)

        try:
            server.send_message(mensaje)
            enviados += 1
            log_file.write(f"[OK] {persona['email']}\n")
            progress["value"] += 1

            porcentaje = int((progress["value"] / total) * 100)

            label_contador.config(text=f"{progress['value']} / {total} enviados")
            label_porcentaje.config(text=f"{porcentaje}%")

            principal.update_idletasks()

            time.sleep(intervalo_envio)  # evitar bloqueo

        except Exception as e:
            error_msg = str(e)
            print("Error enviando a:", persona["email"], error_msg)
            log_file.write(f"[ERROR] {persona['email']} -> {error_msg}\n")
    log_file.write("\n===== RESUMEN =====\n")
    log_file.write(f"Enviados correctamente: {enviados}\n")
    log_file.write(f"No enviados: {total - enviados}\n")
    log_file.close()
    server.quit()

    messagebox.showinfo("Éxito", f"{enviados} correos enviados correctamente")


# ========================
# BOTONES Y ELEMENTOS
# ========================

Button(principal, text="Seleccionar Excel", command=seleccionar_excel)\
    .grid(row=3, column=0, pady=5)

label_excel = Label(principal, text="Ningún archivo seleccionado", bg="black", fg="white")
label_excel.grid(row=3, column=1)


Button(principal, text="Seleccionar HTML", command=seleccionar_html)\
    .grid(row=4, column=0, pady=5)

label_html = Label(principal, text="Ningún archivo seleccionado", bg="black", fg="white")
label_html.grid(row=4, column=1)


progress = ttk.Progressbar(principal, orient="horizontal", length=400, mode="determinate")
label_contador = Label(principal, text="0 / 0 enviados", bg="black", fg="white")
label_contador.grid(row=6, column=1)

label_porcentaje = Label(principal, text="0%", bg="black", fg="white")
label_porcentaje.grid(row=7, column=1)
progress.grid(row=5, column=1, pady=15)


Button(
    principal,
    text="Enviar Correos",
    command=enviar,
    bg="#215cc1",
    fg="white",
    width=30
).grid(row=6, column=1, pady=15)


principal.mainloop()