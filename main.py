__version__ = ['1.2.0']


from kaki.app import App
from kivy.graphics.svg import Window
from kivymd.app import MDApp
from kivy.factory import Factory




class LiveApp(MDApp, App):
    import os


    #Window.size = (400, 750)
    #DEBUG = 0

    KV_FILES = {

        os.path.join(os.getcwd(), "screens/ScreenManager.kv"),
        os.path.join(os.getcwd(), "screens/login_screens/LoginScreens.kv"),

        os.path.join(os.getcwd(), "screens/config_screens/ConfigScreen.kv"),
        os.path.join(os.getcwd(), "screens/cliente_screens/ClienteScreen.kv"),
        #os.path.join(os.getcwd(), "screens/camera_screens/CameraScreen.kv"),

    }

    CLASSES = {

        "MainScreenManager": "screens.ScreenManager",
        "LoginScreen": "screens.login_screens.LoginScreens",

        "ConfigScreen": "screens.config_screens.ConfigScreen",
        "ClienteScreen": "screens.cliente_screens.ClienteScreen",
        #"CameraScreen": "screens.camera_screens.CameraScreen",

    }

    AUTORELOADER_PATHS = [
        (".", {"recursive": True}),
    ]

    def build_app(self, **kwargs):

        #from android.permissions import request_permissions, Permission  # type: ignore
        #request_permissions([Permission.READ_EXTERNAL_STORAGE, Permission.WRITE_EXTERNAL_STORAGE, Permission.CAMERA])

        self.theme_cls.theme_style = "Dark"
        self.theme_cls.primary_palette = "Orange"
        return Factory.MainScreenManager()

    def tema_escuro(self):
        self.theme_cls.theme_style = "Dark"
        self.theme_cls.primary_palette = "Orange"

    def tema_claro(self):
        self.theme_cls.theme_style = 'Light'
        self.theme_cls.primary_palette = "Orange"


if __name__ == "__main__":
    LiveApp().run()


