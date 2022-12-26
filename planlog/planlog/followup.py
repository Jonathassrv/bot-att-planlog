from botcity.core import DesktopBot
from datetime import date
# Uncomment the line below for integrations with BotMaestro
# Using the Maestro SDK
# from botcity.maestro import *


class Bot(DesktopBot):
    def action(self, execution=None):

        print("Follow")

        # FOLLOW UP
        self.browse("https://docs.google.com/spreadsheets/d/1-X2-R4R8IpseM__uS-HJcPij_mJtz0k6eGM0r7Sa0aE/edit#gid=0")
        if not self.find("reconhecimentofollow", matching=0.97, waiting_time=30000):
            self.not_found("reconhecimentofollow")
        if not self.find("arquivo", matching=0.97, waiting_time=20000):
            self.not_found("arquivo")
        self.click()
        if not self.find("download", matching=0.97, waiting_time=10000):
            self.not_found("download")
        self.click()
        if not self.find("excel", matching=0.97, waiting_time=10000):
            self.not_found("excel")
        self.click()
        if not self.find("baixado", matching=0.97, waiting_time=50000):
            self.not_found("baixado")
        self.click()
        if not self.find("habilitaredicao", matching=0.97, waiting_time=10000):
            self.not_found("habilitaredicao")
        self.click()
        if not self.find("validacaoedicao", matching=0.97, waiting_time=10000):
            self.not_found("validacaoedicao")
        if not self.find("arquivoexcel", matching=0.97, waiting_time=10000):
            self.not_found("arquivoexcel")
        self.click()
        if not self.find("salvarcomo", matching=0.97, waiting_time=10000):
            self.not_found("salvarcomo")
        self.click()
        if not self.find("procurar", matching=0.97, waiting_time=10000):
            self.not_found("procurar")
        self.click()
        if not self.find("barraendereco", matching=0.97, waiting_time=10000):
            self.not_found("barraendereco")
        self.click()
        self.paste('U:\GESTAO ABASTECIMENTO\GATE\PCP\Dados\Comodato')
        self.enter()
        if not self.find("followuparq", matching=0.97, waiting_time=10000):
            self.not_found("followuparq")
        self.doubleclick()
        self.kb_type('s')
        self.wait(6000)
        self.alt_f4()
    action()
    
    def not_found(self, label):
        print(f"Element not found: {label}")


if __name__ == '__main__':
    Bot.main()