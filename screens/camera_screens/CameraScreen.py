import time
from kivy.uix.image import Image
from kivy.uix.screenmanager import Screen
from kivymd.uix.boxlayout import MDBoxLayout
from kivymd.uix.list import TwoLineRightIconListItem, ImageRightWidget, IRightBodyTouch  # type: ignore


class ImageRightWidget(IRightBodyTouch, Image):
    pass


class Content(MDBoxLayout):
    '''Custon content.'''


class CameraScreen(Screen):
    def capture(self):
        import os
        from kivy.storage import FileSystem  # type: ignore
        from android.permissions import request_permissions, Permission  # type: ignore

        request_permissions([Permission.CAMERA, Permission.WRITE_EXTERNAL_STORAGE])

        fs = FileSystem()
        folder_path = fs.join(fs.documents_dir, 'FinderPlanilha', "fotos")

        if not fs.exists(folder_path):
            fs.mkdirs(folder_path)  # Create directories recursively

        camera = self.ids['camera']
        camera.export_to_png(os.path.join(folder_path, f"IMG.png"))



