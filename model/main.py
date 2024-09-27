import subprocess
import os

def main():
    try:
        # Rutas completas a los scripts
        cargar_hu_path = r"C:\Users\ASUS\Documents\GitHub\Generate-CP-IA\model\cargar_hu.py"
        hu_path = r"C:\Users\ASUS\Documents\GitHub\Generate-CP-IA\model\hu.py"

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

if __name__ == "__main__":
    main()
