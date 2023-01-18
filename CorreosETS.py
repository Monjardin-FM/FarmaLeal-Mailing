from tkinter import *
from tkinter import ttk
from tkinter import filedialog
from tkinter import messagebox
import xlrd
import smtplib
import ssl
import getpass
import re
import os
from email import encoders
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

principal = Tk()
principal.iconbitmap(
    r'C:\Users\rober\OneDrive\Documentos\Python\Correos Masivos\ets.ico')
principal.geometry('490x200')
principal.configure(bg='black')
principal.title('Mailing')

label1 = Label(principal, text="Usuario: ", bg='black', fg="white", height="3")
label1.grid(row=5, column=1)
correo_string = StringVar()
caja1 = Entry(principal, textvariable=correo_string, width=50)
caja1.grid(row=5, column=2)

label2 = Label(principal, text="Contraseña: ",
               bg='black', fg="white", height="3")
label2.grid(row=7, column=1)
contraseña_string = StringVar()
caja2 = Entry(principal, textvariable=contraseña_string, show="*", width=50)
caja2.grid(row=7, column=2)

label3 = Label(principal, text="Asunto del correo: ",
               bg='black', fg="white", height="3")
label3.grid(row=9, column=1)
asunto_string = StringVar()
caja3 = Entry(principal, textvariable=asunto_string, width=50)
caja3.grid(row=9, column=2)


def clicked():
    ruta = filedialog.askopenfilename()
    correo = caja1.get()
    contraseña = caja2.get()
    asunto = caja3.get()
    print(ruta)
    print(correo)
    print(contraseña)
    print(asunto)

    openFile = xlrd.open_workbook(ruta)
    sheet = openFile.sheet_by_name("Mailing")

    print("N de filas", sheet.nrows)
    print("N de columnas", sheet.ncols)

    correos = []
    destinatario = []
    for i in range(sheet.nrows):
        destinatario.append(sheet.cell_value(i, 0))
        correos.append(sheet.cell_value(i, 1))
        print(sheet.cell_value(i, 0), "     ", sheet.cell_value(i, 1))
        print("**********************")
    contador = 0
    for element in correos:
        contador
        print(element)
        mensaje = MIMEMultipart("alternative")
        mensaje["Subject"] = asunto
        mensaje["From"] = correo
        mensaje["To"] = element

        html = """
                <!DOCTYPE html>
                <html lang="en">
                <head>
                    <meta charset="UTF-8" />
                    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
                    <meta name="viewport" content="width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no">
                    <title>ESCUELA EN TECNOLOGÍAS SUSTENTABLES</title>
                    <style type="text/css">

                        body{
                            background: #000;
                        }
                    </style>
                </head>
                <body background="#000">
                    <p>Dale click a la imagen para comunicarte con nosotros<p/>
                    <a href= "https://wa.link/i0z0v3">
                    <img src="https://i.ibb.co/TK5ZBD7/correo4.png">
                    </a>
                </body>
                </html>
                """
    # El contenido del mensaje como HTML
        parte_html = MIMEText(html, "html")
        mensaje.attach(parte_html)
        email_content = mensaje.as_string()
        if re.search(".@gmail.com", correo):
            server = smtplib.SMTP('smtp.gmail.com:587')
        elif re.search(".@outlook.com", correo):
            server = smtplib.SMTP('smtp-mail.outlook.com:587')
        server.starttls()
        server.login(correo, contraseña)
        print("Iniciando sesión...")
        server.sendmail(correo, element, email_content)
        server.quit()
        print("Correo enviado")
        contador = contador + 1

    messagebox.showinfo('Correos enviados', 'Correos enviados exitosamente')


btn = Button(principal, text="Escoger archivo excel y enviar",
             command=clicked, bg='#9cc121')
btn.grid(column=2, row=13)
principal.mainloop()
