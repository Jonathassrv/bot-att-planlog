from botcity.core import DesktopBot
from datetime import date

# Uncomment the line below for integrations with BotMaestro
# Using the Maestro SDK
#from botcity.maestro import *


class Bot(DesktopBot):

    def action(self, execution=None):
        sucesso = "Atualização concluída!"
        
    #INSTRUÇÕES INICIAIS
        today = date.today()
        d1 = today.strftime("%d.%m.%Y")
        i = 0
        print("Selecione o que deseja atualizar:")
        print("(1) para atualizar tudo, (2) para atualizar bases separadas em caso de erro")
        escolha = int(input("Selecione: "))
        if escolha == 1:
            i = 0
        else:
            escolha == 2
            print("Escolha o que deseja atualizar: ")
            print("(1) NFs TRANSFERÊNCIA, (2) NFs COLIGADAS, (3) ESTOQUE VESPASIANO, (4) ESTOQUE GERAL, (5) ESTOQUE DIMB, (6) VENDA COMODATO, (7) ATUALIZAR PLANLOG")
            print("*Lembre-se de deixar o SAP aberto e na tela após selecionar a opção.")
            selecao = int(input("Digite o número referente a atualização: "))
            
            if selecao == 1:
                i = 1
            elif selecao == 2:
                i = 2
            elif selecao == 3:
                i = 3
            elif selecao == 4:
                i = 4
            elif selecao == 5:
                i = 5
            elif selecao == 6:
                i = 6
            else:
                i = 7        

    #NF EM TRANSITO
        if i < 2:
            print("*******************")
            print("****ABRINDO SAP****")
            print("*******************")            
            
            self.execute('C:\Program Files (x86)\SAP\FrontEnd\SapGui\saplogon.exe')
            if not self.find( "sap_producao", matching=0.97, waiting_time=10000):
                self.not_found("sap_producao")
            self.doubleclick()
            if not self.find( "reconhecerlogin", matching=0.97, waiting_time=10000):
                self.not_found("reconhecerlogin")      
            print("Insira os seus dados de login")
            self.wait(10000)        
            sucesso = "Concluído!"
            print(sucesso)

            print("*************************")
            print("**ATUALIZANDO TRANSITOS**")
            print("*************************")

            self.paste('ztranicmsnfe')
            self.enter()
            if not self.find( "variante", matching=0.97, waiting_time=10000):
                self.not_found("variante")
            self.click()        
            if not self.find( "criadopor", matching=0.97, waiting_time=10000):
                self.not_found("criadopor")
            self.click()        
            self.tab()
            self.control_a()
            self.paste('victor.p')
            self.key_f8()
            if not self.find( "transitocom", matching=0.97, waiting_time=10000):
                self.not_found("transitocom")             
            self.doubleclick()
            if not self.find( "transferencia", matching=0.97, waiting_time=20000):
                self.not_found("transferencia")
            self.click()
            self.key_f8()
            if not self.find( "simbolotres", matching=0.97, waiting_time=30000):
                self.not_found("simbolotres")               
            self.right_click_at(x=100, y=550)
            if not self.find( "planilha", matching=0.97, waiting_time=10000):
                self.not_found("planilha")
            self.click()        
            if not self.find( "tipogoogle", matching=0.97, waiting_time=10000):
                self.not_found("tipogoogle")
            self.click()        
            self.type_up()
            self.type_up()
            self.enter()               
            self.enter()
            if not self.find( "transferencias", matching=0.97, waiting_time=10000):
                self.not_found("transferencias")
            self.click()
            self.doubleClick()
            self.kb_type('s')
            if not self.find( "reconhecimento_excel", matching=0.97, waiting_time=10000):
                self.not_found("reconhecimento_excel")        
            self.alt_f4()
            self.key_esc()
            print(sucesso)
            i += 1

    #COLIGADAS
        if i < 3:
            print("*************************")
            print("**ATUALIZANDO COLIGADAS**")
            print("*************************")

            if not self.find( "coligadastipo", matching=0.97, waiting_time=10000):
                self.not_found("coligadastipo")
            self.click()
            self.key_f8()
            if not self.find( "simbolotres", matching=0.97, waiting_time=30000):
                self.not_found("simbolotres")
            
            self.right_click_at(x=100, y=550)
            if not self.find( "planilha", matching=0.97, waiting_time=10000):
                self.not_found("planilha")
            self.click()        
            if not self.find( "tipogoogle", matching=0.97, waiting_time=10000):
                self.not_found("tipogoogle")
            self.click()     
            self.type_up()
            self.type_up()
            self.enter()
            self.enter()
            if not self.find( "coligadas", matching=0.97, waiting_time=10000):
                self.not_found("coligadas")                  
            self.doubleclick()
            self.kb_type('s')
            if not self.find( "reconhecimento_excel", matching=0.97, waiting_time=10000):
                self.not_found("reconhecimento_excel")
            self.alt_f4()
            self.key_esc()
            self.key_esc()
            print(sucesso)
            i += 1

    #ESTOQUE VESPASIANO
        if i < 4:
            print("**********************************")
            print("**ATUALIZANDO ESTOQUE VESPASIANO**")
            print("**********************************")

            self.wait(3000)
            self.paste('iq09')
            self.enter()
            if not self.find( "variante", matching=0.97, waiting_time=10000):
                self.not_found("variante")
            self.click()
            if not self.find( "criadopor", matching=0.97, waiting_time=10000):
                self.not_found("criadopor")
            self.click()
            self.tab()
            self.control_a()
            self.paste('victor.p')
            self.key_f8()
            if not self.find( "comodatovesp", matching=0.97, waiting_time=10000):
                self.not_found("comodatovesp")
            self.doubleclick()        
            if not self.find( "periodo", matching=0.97, waiting_time=10000):
                self.not_found("periodo")
            self.click()
            self.tab()
            today = date.today()
            d1 = today.strftime("%d.%m.%Y")
            self.control_a()
            self.paste(d1)
            self.tab()
            self.paste(d1)
            self.key_f8()
            if not self.find( "reconhecer", matching=0.97, waiting_time=30000):
                self.not_found("reconhecer")        
            self.right_click_at(x=100, y=550)
            if not self.find( "planilha", matching=0.97, waiting_time=10000):
                self.not_found("planilha")
            self.click()
            if not self.find( "tipogoogle", matching=0.97, waiting_time=10000):
                self.not_found("tipogoogle")
            self.click()            
            self.type_up()
            self.type_up()
            self.enter()
            self.enter()
            if not self.find( "estvesp", matching=0.97, waiting_time=10000):
                self.not_found("estvesp")        
            self.doubleclick()
            self.kb_type('s')
            if not self.find( "reconhecimento_excel", matching=0.97, waiting_time=10000):
                self.not_found("reconhecimento_excel")
            self.alt_f4()
            self.wait(1500)
            self.key_esc()
            print(sucesso)
            i += 1

    #ESTOQUE GERAL
        if i < 5:
            print("*****************************")
            print("**ATUALIZANDO ESTOQUE GERAL**")
            print("*****************************")

            if not self.find( "variante", matching=0.97, waiting_time=10000):
                self.not_found("variante")
            self.click()
            if not self.find( "criadopor", matching=0.97, waiting_time=10000):
                self.not_found("criadopor")
            self.click()
            self.tab()
            self.control_a()
            self.paste('victor.p')
            self.key_f8()
            if not self.find( "estgeral", matching=0.97, waiting_time=10000):
                self.not_found("estgeral")        
            self.doubleclick()
            if not self.find( "periodo", matching=0.97, waiting_time=10000):
                self.not_found("periodo")
            self.click()
            self.tab()
            self.control_a()
            self.paste(d1)
            self.tab()
            self.paste(d1)
            self.key_f8()
            if not self.find( "reconhecer", matching=0.97, waiting_time=30000):
                self.not_found("reconhecer") 
            self.right_click_at(x=100, y=550)
            if not self.find( "planilha", matching=0.97, waiting_time=10000):
                self.not_found("planilha")
            self.click()       
            if not self.find( "tipogoogle", matching=0.97, waiting_time=10000):
                self.not_found("tipogoogle")
            self.click() 
            self.type_up()
            self.type_up()
            self.enter()
            self.enter()
            if not self.find( "estgeralexcel", matching=0.97, waiting_time=10000):
                self.not_found("estgeralexcel")        
            self.doubleclick()
            self.kb_type('s')
            if not self.find( "reconhecimento_excel", matching=0.97, waiting_time=10000):
                self.not_found("reconhecimento_excel")
            self.alt_f4()
            self.key_esc()
            print(sucesso)
            i += 1

    #ESTOQUE DIMB
        if i < 6:
            print("****************************")
            print("**ATUALIZANDO ESTOQUE DIMB**")
            print("****************************")
            if not self.find( "variante", matching=0.97, waiting_time=10000):
                self.not_found("variante")
            self.click()
            if not self.find( "criadopor", matching=0.97, waiting_time=10000):
                self.not_found("criadopor")
            self.click()
            self.tab()
            self.control_a()
            self.paste('brena.r')
            self.key_f8()
            if not self.find( "periodo", matching=0.97, waiting_time=10000):
                self.not_found("periodo")
            self.click()
            self.tab()
            self.control_a()
            self.paste(d1)
            self.tab()
            self.paste(d1)
            self.key_f8()
            if not self.find( "reconhecer", matching=0.97, waiting_time=30000):
                self.not_found("reconhecer")
            self.right_click_at(x=100, y=550)
            if not self.find( "planilha", matching=0.97, waiting_time=10000):
                self.not_found("planilha")
            self.click()       
            if not self.find( "tipogoogle", matching=0.97, waiting_time=10000):
                self.not_found("tipogoogle")
            self.click()         
            self.type_up()
            self.type_up()
            self.enter()
            self.enter()
            if not self.find( "estdimb", matching=0.97, waiting_time=10000):
                self.not_found("estdimb")        
            self.doubleclick()
            self.kb_type('s')
            if not self.find( "reconhecimento_excel", matching=0.97, waiting_time=10000):
                self.not_found("reconhecimento_excel")
            self.alt_f4()
            self.key_esc()
            self.key_esc()
            print(sucesso)
            i += 1

    #VENDA COMODATO
        if i < 7:
            print("*********************")
            print("**ATUALIZANDO VENDA**")
            print("*********************")
            self.wait(2000)
            self.paste('zsdr315')
            self.enter()
            if not self.find( "variante", matching=0.97, waiting_time=10000):
                self.not_found("variante")
            self.click()
            if not self.find( "variantevendacom", matching=0.97, waiting_time=10000):
                self.not_found("variantevendacom")
            self.doubleclick()
            self.key_f8()
            if not self.find( "reconhecimentovenda", matching=0.97, waiting_time=30000):
                self.not_found("reconhecimentovenda")        
            self.right_click_at(x=100, y=550)
            if not self.find( "planilha", matching=0.97, waiting_time=10000):
                self.not_found("planilha")
            self.click()       
            if not self.find( "tipogoogle", matching=0.97, waiting_time=10000):
                self.not_found("tipogoogle")
            self.click()
            self.type_up()
            self.type_up()
            self.enter()
            self.enter()
            if not self.find( "vendacom", matching=0.97, waiting_time=10000):
                self.not_found("vendacom")        
            self.doubleclick()
            self.kb_type('s')
            if not self.find( "reconhecimento_excel", matching=0.97, waiting_time=10000):
                self.not_found("reconhecimento_excel")
            self.alt_f4()
            print(sucesso)
            i += 1

    #FECHAMENTO SAP
            print("****************")
            print("**FECHANDO SAP**")
            print("****************")

            self.wait(2000)
            self.alt_f4()
            if not self.find( "sairsap", matching=0.97, waiting_time=10000):
                self.not_found("sairsap")
            self.click()

    #ATUALIZAR PLANLOG
        if i < 8:
            print("*******************")
            print("**ABRINDO PLANLOG**")
            print("*******************")

            self.execute('U:\GESTAO ABASTECIMENTO\GATE\PCP\PCP 2022\PLANEJAMENTO COMODATO\Planlog Comodato Dezembro.xlsm')
            if not self.find( "reconhecimento_excel", matching=0.97, waiting_time=30000):
                self.not_found("reconhecimento_excel")                                         
            if not self.find( "att_planlog", matching=0.97, waiting_time=10000):
                self.not_found("att_planlog")            
            self.click()                        
            self.wait(25000)      
            if not self.find( "salvarplanlog", matching=0.97, waiting_time=10000):
                self.not_found("salvarplanlog")            
            self.click()        
            self.wait(5000)
            if not self.find( "arquivoexcel", matching=0.97, waiting_time=10000):
                self.not_found("arquivoexcel")
            self.click()
            if not self.find( "salvarcomo", matching=0.97, waiting_time=10000):
                self.not_found("salvarcomo")
            self.click()
            if not self.find( "procurar", matching=0.97, waiting_time=10000):
                self.not_found("procurar")
            self.click()
            if not self.find( "pastaplanlogdia", matching=0.97, waiting_time=10000):
                self.not_found("pastaplanlogdia")
            self.click()
            if not self.find( "pastadez", matching=0.97, waiting_time=10000):
                self.not_found("pastadez")
            self.doubleclick()        
            self.tab()
            self.tab()
            self.type_right()
            self.kb_type(' ')
            self.paste(d1)
            self.enter()
            self.wait(4000)
            self.alt_f4()
            print("Planlog atualizada com sucesso!")
            i += 1
    
    def not_found(self, label):
        print(f"Elemento não encontrado: {label}")
        print("Execute o código novamente.")
    
            
if __name__ == '__main__':
    Bot.main()











