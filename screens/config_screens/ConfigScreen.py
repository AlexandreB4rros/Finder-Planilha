
from kivymd.uix.screen import MDScreen
from main import LiveApp



class ConfigScreen(MDScreen):
    def aplicar_config(self):
        self.Selecionar_Tema()

    def mudar_tela(self):
        self.manager.transition.direction = 'left'
        self.manager.current = 'HomeScreen'

    def Selecionar_Tema(self):
        if self.ids.selecionar_tema_escuro.active:
            app = LiveApp()
            app.tema_escuro()

        elif self.ids.selecionar_tema_claro.active:
            app = LiveApp()
            app.tema_claro()
        else:
            pass





