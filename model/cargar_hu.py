import os
from docx import Document

def copiar_tabla_hu_y_pegar():
    # Pedir el nombre de la historia de usuario
    nombre_hu = input("Ingresa el nombre de la historia de usuario (sin extensión, e.g., HU-003): ")
    
    # Rutas de los archivos
    ruta_hu = os.path.join(r"C:\Users\ASUS\Desktop\pruebas\husp4", f"{nombre_hu}.docx")
    ruta_test = r"C:\Users\ASUS\Documents\GitHub\Generate-CP-IA\Archivos_imp\Criterios_de_aceptacion.docx"
    
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
        print("Archivo test.docx cargado correctamente.")
    except Exception as e:
        print(f"Error al cargar test.docx: {e}")
        return

    # Buscar la tabla en HU-003.docx
    tabla_copiar = None
    for table in doc_hu.tables:
        # Suponiendo que la tabla de "Criterios de Aceptación" es la primera
        if "Criterios de Aceptación" in table.cell(0, 0).text:  # Verificamos el encabezado de la tabla
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

    # Guardar el archivo test.docx con la nueva tabla
    try:
        doc_test.save(ruta_test)
        print("La tabla se ha copiado y test.docx ha sido guardado correctamente.")
    except Exception as e:
        print(f"Error al guardar test.docx: {e}")

if __name__ == "__main__":
    copiar_tabla_hu_y_pegar()
