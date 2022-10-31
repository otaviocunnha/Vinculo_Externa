from botcity.core import DesktopBot
import pandas as pd
import win32com.client as win32
outlook = win32.Dispatch('outlook.application')



email = outlook.CreateItem(0)

email.To = "otavio-cunha@hotmail.com "

email.Subject = 'E-mail do seu amigo Robotnik'



clientes =pd.read_excel(r'C:\Users\Otavio.cunha\Desktop\Projeto Python\Agendamentos\externa\Codigos.xlsx')


class Bot(DesktopBot):
    def action(self, execution=None):

        if not self.find( "Protheus", matching=0.97, waiting_time=10000):
            self.not_found("Protheus")
        self.click()
        for c in clientes['Codigo']:
            if not self.find( "Ferramentas", matching=0.97, waiting_time=90000):
                self.not_found("Ferramentas")
            self.click()
            if not self.find( "vicnular", matching=0.97, waiting_time=90000):
                self.not_found("vicnular")
            self.click()
            if not self.find( "vincularar", matching=0.97, waiting_time=90000):
                self.not_found("vincularar")
            self.click()
            self.wait(8000)
            self.paste('000273')
            self.wait(3000)
            self.tab()
            self.wait(2000)
            self.paste(f'{c}'.zfill(6))
            self.wait(2000)
            self.enter()
            if self.find( "Fec", matching=0.97, waiting_time=20000):
                self.click()
                if not self.find( "can", matching=0.97, waiting_time=20000):
                    self.not_found("can")
                self.click()
            else:
                pass
                self.wait(2000)
                self.tab()
                # self.type_down()
                self.enter()
                self.wait(3000)
                if not self.find( "confirmar", matching=0.97, waiting_time=90000):
                    self.not_found("confirmar")
                self.click()
                if not self.find( "fehcar", matching=0.97, waiting_time=9800000):
                    self.not_found("fehcar")
                self.click()
                self.wait(20000)
            if not self.find( "Desvincular", matching=0.97, waiting_time=10000):
                self.not_found("Desvincular")
            self.click()
            for c in clientes['Codigo']:
                if not self.find( "codigo", matching=0.97, waiting_time=10000):
                    self.not_found("codigo")
                self.click_relative(62, 7)
                self.paste(f'{c}'.zfill(6))
                self.enter()
                if self.find( "fechar", matching=0.97, waiting_time=10000):
                    self.click()
                else:
                    self.tab()
                    self.enter()

              
              
              
              
              
              
              
              
        

    def not_found(self, label):
        print(f"Element not found: {label}")


if __name__ == '__main__':
    Bot.main()


email.HTMLBody = '''<p> Vinculos de externa realizada com sucesso</p>


<p>Atenciosamente,</p>
<p>Seu amigo e companheiro, ROBOTNIK </p>
'''


#email.Send()
print('e-email enviado ')


