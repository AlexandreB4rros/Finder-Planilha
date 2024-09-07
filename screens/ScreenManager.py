from kivy.uix.screenmanager import ScreenManager, SlideTransition

sm = ScreenManager()


class MainScreenManager(ScreenManager):
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.transition = SlideTransition()


#    def salvar_dados(self):
#        cto = self.get_screen('ClienteScreenDois').ids.cto.text
#        StsP1 = self.get_screen('PortaUm').ids.StsP1.text
#        SNP1 = self.get_screen('PortaUm').ids.SNP1.text
#        Pwd1 = self.get_screen('PortaUm').ids.Pwd1.text
#        Loid1 = self.get_screen('PortaUm').ids.Loid1.text
#        Obs1 = self.get_screen('PortaUm').ids.Obs1.text

