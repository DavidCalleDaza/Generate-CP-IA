import os
import sys
from docx import Document

def copiar_tabla_hu_y_pegar(nombre_hu):
    # Rutas de los archivos
    ruta_hu = os.path.join(r"C:\Users\d4vid\OneDrive\Escritorio\Generate-CP-IA\Archivos_imp\Historias de usuario", f"{nombre_hu}.docx")
    ruta_test = r"C:\Users\d4vid\OneDrive\Escritorio\Generate-CP-IA\Archivos_imp\Criterios_de_aceptacion.docx"
    
    # Cargar el documento de la historia de usuario
    try:
        doc_hu = Document(ruta_hu)
        print(f"Archivo {nombre_hu}.docx cargado correctamente.")
    except Exception as e:
        print(f"Error al cargar {nombre_hu}.docx: {e}")
        return
    
    # Crear un nuevo documento para el archivo de destino si no existe
    try:
        doc_test = Document(ruta_test)
        print("Archivo Criterios_de_aceptacion.docx cargado correctamente.")
    except Exception as e:
        print(f"Error al cargar Criterios_de_aceptacion.docx: {e}")
        return

    # Buscar la tabla en el archivo de historia de usuario
    tabla_copiar = None
    for table in doc_hu.tables:
        if "Criterios de Aceptación" in table.cell(0, 0).text:
            tabla_copiar = table
            break

    if not tabla_copiar:
        print("No se encontró la tabla 'Criterios de Aceptación'.")
        return

    # Copiar el contenido de la tabla
    print("Copiando la tabla 'Criterios de Aceptación'...")
    for row in tabla_copiar.rows:
        nueva_fila = doc_test.add_table(rows=1, cols=len(row.cells)).rows[-1]
        for idx, cell in enumerate(row.cells):
            nueva_fila.cells[idx].text = cell.text

    # Guardar el archivo Criterios_de_aceptacion.docx con la nueva tabla
    try:
        doc_test.save(ruta_test)
        print("La tabla se ha copiado y Criterios_de_aceptacion.docx ha sido guardado correctamente.")
    except Exception as e:
        print(f"Error al guardar Criterios_de_aceptacion.docx: {e}")

if __name__ == "__main__":
    if len(sys.argv) > 1:
        nombre_hu = sys.argv[1]
        copiar_tabla_hu_y_pegar(nombre_hu)
    else:
        print("Por favor, proporcione el nombre de la historia de usuario.")
