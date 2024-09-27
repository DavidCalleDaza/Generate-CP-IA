import os
import base64
import tkinter as tk
from tkinter import messagebox
from cryptography.hazmat.primitives import serialization
from cryptography.hazmat.primitives.asymmetric import rsa, padding
from cryptography.hazmat.primitives import hashes
from cryptography.hazmat.backends import default_backend
import threading
import time

# Variables globales
private_key = None
public_key = None
password_encrypted = None

def generar_claves():
    global private_key, public_key
    # Generar claves RSA
    private_key = rsa.generate_private_key(public_exponent=65537, key_size=2048)
    public_key = private_key.public_key()

    # Serializar y guardar las claves en archivos
    # Definir una contraseña para proteger la clave privada
    contrasena = b"john1981*"  # Usar bytes
    encrypted_private_key = private_key.private_bytes(
        encoding=serialization.Encoding.PEM,
        format=serialization.PrivateFormat.TraditionalOpenSSL,
        encryption_algorithm=serialization.BestAvailableEncryption(contrasena)
    )

    with open("private_key.pem", "wb") as f:
        f.write(encrypted_private_key)
    
    with open("public_key.pem", "wb") as f:
        f.write(public_key.public_bytes(
            encoding=serialization.Encoding.PEM,
            format=serialization.PublicFormat.SubjectPublicKeyInfo))

    print("Claves generadas y guardadas en archivos.")


def cifrar_contraseña(password):
    # Cifrar la contraseña utilizando la clave pública
    encrypted = public_key.encrypt(
        password.encode(),
        padding.OAEP(
            mgf=padding.MGF1(algorithm=hashes.SHA256()),
            algorithm=hashes.SHA256(),
            label=None
        )
    )
    return base64.b64encode(encrypted).decode()

def actualizar_claves():
    while True:
        generar_claves()
        time.sleep(120)  # Esperar 2 minutos antes de regenerar las claves

def ejecutar_scripts():
    try:
        # Rutas completas a los scripts
        limpiar_hu_path = r"C:\Users\ASUS\Documents\GitHub\Generate-CP-IA\model\limpiar.py"
        cargar_hu_path = r"C:\Users\ASUS\Documents\GitHub\Generate-CP-IA\model\cargar_hu.py"
        hu_path = r"C:\Users\ASUS\Documents\GitHub\Generate-CP-IA\model\hu.py"

        # 0. Ejecutar el script limpiar.py
        subprocess.run(['python', limpiar_hu_path], check=True)
        print("Ejecutado: limpiar.py")

        # 1. Ejecutar el script cargar_hu.py
        subprocess.run(['python', cargar_hu_path], check=True)
        print("Ejecutado: cargar_hu.py")

        # 2. Ejecutar el script hu.py
        subprocess.run(['python', hu_path], check=True)
        print("Ejecutado: hu.py")

        # 3. Abrir el archivo CP-VBA
        file_path = r"C:\Users\ASUS\Documents\GitHub\Generate-CP-IA\Archivos_xpo\VBA\CP-VBA.xlsm"
        os.startfile(file_path)
        print(f"Abrir archivo: {file_path}")

    except subprocess.CalledProcessError as e:
        print(f"Error al ejecutar el script: {e}")
    except Exception as e:
        print(f"Ocurrió un error: {e}")

def verificar_contraseña():
    global password_encrypted
    user_password = entrada_contraseña.get()
    if not password_encrypted:
        # Cifrar la contraseña inicial
        password_encrypted = cifrar_contraseña(user_password)
        messagebox.showinfo("Éxito", "Contraseña guardada correctamente. Puedes ingresar al sistema.")
    else:
        # Verificar la contraseña ingresada
        try:
            decrypted_password = public_key.decrypt(
                base64.b64decode(password_encrypted),
                padding.OAEP(
                    mgf=padding.MGF1(algorithm=hashes.SHA256()),
                    algorithm=hashes.SHA256(),
                    label=None
                )
            ).decode()
            if decrypted_password == user_password:
                messagebox.showinfo("Éxito", "Acceso concedido.")
            else:
                messagebox.showerror("Error", "Contraseña incorrecta.")
        except Exception as e:
            print(f"Error al verificar la contraseña: {e}")

# Crear y mostrar la ventana principal
ventana = tk.Tk()
ventana.title("Ingreso al Sistema")

# Etiqueta y campo de entrada
etiqueta = tk.Label(ventana, text="Ingresa la contraseña:")
etiqueta.pack(pady=10)

entrada_contraseña = tk.Entry(ventana, show="*", width=30)
entrada_contraseña.pack(pady=5)

# Botón para verificar la contraseña
boton_verificar = tk.Button(ventana, text="Ingresar", command=verificar_contraseña)
boton_verificar.pack(pady=10)

# Iniciar la generación de claves en un hilo separado
threading.Thread(target=actualizar_claves, daemon=True).start()

# Iniciar la generación de claves al inicio
generar_claves()

# Iniciar el bucle principal
ventana.mainloop()
