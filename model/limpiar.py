import os
from docx import Document

# Rutas de los archivos
resultado_file_path = r"C:\Users\ASUS\Documents\GitHub\Generate-CP-IA\Archivos_imp\resultado.xlsx"
docx_file_path = r"C:\Users\ASUS\Documents\GitHub\Generate-CP-IA\Archivos_imp\Criterios_de_aceptacion.docx"

def eliminar_archivo_resultado():
    # Eliminar el archivo resultado.xlsx si existe
    if os.path.exists(resultado_file_path):
        os.remove(resultado_file_path)
        print("Archivo resultado.xlsx eliminado.")
    else:
        print("El archivo resultado.xlsx no existe.")

def limpiar_docx_file():
    # Crear un nuevo documento para limpiar el contenido
    doc = Document()
    # Guardar el documento vacío para sobrescribir el existente
    doc.save(docx_file_path)
    print("Contenido de Criterios_de_aceptacion.docx eliminado.")

def main():
    eliminar_archivo_resultado()
    limpiar_docx_file()

if __name__ == "__main__":
    main()
