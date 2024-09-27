import os
from openpyxl import load_workbook
from docx import Document

# Rutas de los archivos
test_file_path = r"C:\Users\d4vid\OneDrive\Escritorio\Generate-CP-IA\Archivos_imp\test.xlsx"
resultado_file_path = r"C:\Users\d4vid\OneDrive\Escritorio\Generate-CP-IA\Archivos_imp\resultado.xlsx"
docx_file_path = r"C:\Users\d4vid\OneDrive\Escritorio\Generate-CP-IA\Archivos_imp\Criterios_de_aceptacion.docx"

def limpiar_test_file():
    # Cargar el libro de Excel
    wb = load_workbook(test_file_path)
    # Seleccionar la hoja HISTORIA DE USUARIO
    if 'HISTORIA DE USUARIO' in wb.sheetnames:
        sheet = wb['HISTORIA DE USUARIO']
        # Eliminar filas desde la 7 en adelante
        sheet.delete_rows(7, sheet.max_row - 6)  # Desde la fila 7 hasta la última fila
        # Guardar los cambios
        wb.save(test_file_path)
        print("Información borrada de test.xlsx desde la fila 7 en adelante.")
    else:
        print("La hoja 'HISTORIA DE USUARIO' no se encuentra en el archivo test.xlsx.")

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
    limpiar_test_file()
    eliminar_archivo_resultado()
    limpiar_docx_file()

if __name__ == "__main__":
    main()
