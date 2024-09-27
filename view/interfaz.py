import os
import subprocess
from kivy.app import App
from kivy.uix.button import Button
from kivy.uix.anchorlayout import AnchorLayout
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.textinput import TextInput

class TestCaseGeneratorApp(App):
    def build(self):
        # Layout principal
        layout = AnchorLayout(anchor_x='center', anchor_y='center', padding=20)
        
        # BoxLayout vertical para los elementos
        vertical_layout = BoxLayout(orientation='vertical', spacing=10, size_hint=(None, None), size=(300, 200))
        
        # Crear el botón para generar casos de prueba
        self.generate_button = Button(text='Generar Casos de Prueba', size_hint=(None, None), size=(250, 60))
        self.generate_button.bind(on_press=self.run_script)
        
        # Crear el campo de texto para HU
        self.hu_input = TextInput(hint_text='Ingrese HU', size_hint=(None, None), size=(250, 40))
        
        # Crear el botón para confirmar
        confirm_button = Button(text='Confirmar', size_hint=(None, None), size=(250, 60))
        confirm_button.bind(on_press=self.confirm_action)

        # Agregar widgets al layout vertical
        vertical_layout.add_widget(self.generate_button)
        vertical_layout.add_widget(self.hu_input)
        vertical_layout.add_widget(confirm_button)
        
        # Agregar el layout vertical al layout principal
        layout.add_widget(vertical_layout)

        return layout

    def run_script(self, instance):
        # Obtener el nombre de la historia de usuario desde el campo de texto
        nombre_hu = self.hu_input.text.strip()  # Obtener el texto ingresado y eliminar espacios
        
        if not nombre_hu:
            print("Por favor, ingrese un nombre de historia de usuario.")
            return

        # Ruta del script a ejecutar
        script_path = r"C:\Users\d4vid\OneDrive\Escritorio\Generate-CP-IA\model\main.py"
        try:
            # Ejecutar el script y pasar el nombre_hu como argumento
            subprocess.run(['python', script_path, nombre_hu], check=True)
        except subprocess.CalledProcessError as e:
            print(f"Error al ejecutar el script: {e}")

    def confirm_action(self, instance):
        # Al presionar el botón Confirmar, también ejecutamos el script
        self.run_script(instance)
        print("Acción de Confirmar ejecutada")

if __name__ == '__main__':
    TestCaseGeneratorApp().run()
