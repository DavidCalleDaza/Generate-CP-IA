import os
import subprocess
import threading
from kivy.app import App
from kivy.uix.floatlayout import FloatLayout
from kivy.uix.label import Label
from kivy.uix.textinput import TextInput
from kivy.uix.button import Button

class UserStoryApp(App):
    def build(self):
        # Crear el layout principal
        layout = FloatLayout()

        # Botón para ejecutar el script main.py (más pequeño y centrado)
        execute_button = Button(text="Ejecutar script main.py", size_hint=(0.5, None), height=40)
        execute_button.bind(on_press=self.execute_script)
        execute_button.pos_hint = {'center_x': 0.5, 'top': 0.9}  # Centrando horizontalmente en la parte superior
        layout.add_widget(execute_button)

        # Campo de texto para ingresar el nombre de la historia de usuario (más pequeño y centrado)
        self.story_input = TextInput(hint_text="Ingrese el nombre de la historia de usuario", size_hint=(0.5, None), height=40)
        self.story_input.pos_hint = {'center_x': 0.5, 'top': 0.75}  # Centrando horizontalmente en la parte superior
        layout.add_widget(self.story_input)

        # Botón para confirmar el nombre ingresado (sin espacio entre el campo de texto y este botón)
        confirm_button = Button(text="Confirmar", size_hint=(0.5, None), height=40)
        confirm_button.bind(on_press=self.confirm_story)
        confirm_button.pos_hint = {'center_x': 0.5, 'top': 0.70}  # Un poco más bajo que el campo de texto
        layout.add_widget(confirm_button)

        # Etiqueta para mostrar mensajes, también centrada
        self.message_label = Label(text="", size_hint=(0.5, None), height=40)
        self.message_label.pos_hint = {'center_x': 0.5, 'top': 0.5}  # Centrando horizontalmente en la parte superior
        layout.add_widget(self.message_label)

        return layout

    def confirm_story(self, instance):
        # Obtener el texto ingresado en el campo de texto
        story_name = self.story_input.text
        if story_name:
            # Mostrar un mensaje con el nombre ingresado
            self.message_label.text = f"Historia de Usuario ingresada: {story_name}"
        else:
            self.message_label.text = "Por favor, ingrese un nombre válido."

    def execute_script(self, instance):
        # Ruta del script main.py
        script_path = r'C:\Users\d4vid\OneDrive\Escritorio\Generate-CP-IA\model\main.py'
        
        # Obtener el nombre de la historia de usuario desde el campo de texto
        story_name = self.story_input.text

        # Ejecutar el script en un hilo separado
        threading.Thread(target=self.run_script, args=(script_path, story_name)).start()

    def run_script(self, script_path, story_name):
        try:
            # Ejecutar el script main.py y pasar el nombre de la historia como argumento
            subprocess.run(['python', script_path, story_name], check=True)
            self.update_message("Script main.py ejecutado.")
        except subprocess.CalledProcessError as e:
            self.update_message(f"Error al ejecutar el script: {e}")
        except Exception as e:
            self.update_message(f"Ocurrió un error: {e}")

    def update_message(self, message):
        # Actualiza la etiqueta de mensaje en la interfaz
        self.message_label.text = message

if __name__ == '__main__':
    UserStoryApp().run()
