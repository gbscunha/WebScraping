from tkinter import *
from tkinter import ttk
from selenium.webdriver.support.ui import Select
import tkinter as tk
from tkinter import Tk, filedialog
from pathlib import Path
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import pandas as pd
from time import sleep
from selenium import webdriver
from PIL import ImageTk, Image
from webdriver_manager.chrome import ChromeDriverManager

# Caminho para a raiz do projeto
ROOT_FOLDER = Path(__file__).parent.parent
# Caminho para a pasta onde o chromedriver está
# driver = webdriver.Chrome(ChromeDriverManager().install())
# Listas
lista_ct = [
    'CT 1', 
    'CT 2',
    'CT 3',
    'CT 4'
]


lista_mes = [
    'Janeiro',
    'Fevereiro',
    'Março',
    'Abril',
    'Maio',
    'Junho',
    'Julho',
    'Agosto',
    'Setembro',
    'Outubro',
    'Novembro',
    'Dezembro'
]
lista_id = [7, 8, 18, 20, 21, 29, 30, 32, 37, 45, 47, 52, 54, 56, 58, 89, 90, 92, 99, 101, 107, 111, 114, 115, 116, 117, 118, 121, 122, 129, 134, 137, 144, 145, 146, 148, 159, 164, 166, 169, 175, 187, 194, 202, 206, 208, 221, 231, 236, 237, 256, 257, 264, 266, 267, 269, 274, 275, 280, 288, 290, 295, 296, 301, 302, 303, 305, 307, 309, 313, 317, 318, 342, 345, 346, 351, 353, 358, 370, 375, 377,
            381, 382, 383, 387, 389, 393, 397, 398, 401, 405, 407, 408, 409, 412, 414, 415, 420, 430, 431, 433, 436, 441, 445, 451, 454, 470, 472, 473, 484, 486, 489, 490, 497, 498, 504, 518, 519, 520, 521, 522, 524, 533, 534, 535, 549, 550, 555, 574, 577, 578, 579, 583, 587, 590, 591, 599, 611, 613, 614, 621, 631, 632, 633, 635, 638, 641, 643, 646, 652, 654, 659, 671, 675, 696, 697, 699, 702, 703, 705]
lista_report = []
lista_nome = []
lista_date = []
lista_describ = []
lista_user = []
lista_classif = []
lista_grave = ["Pauta - Controle Mensal", "Pauta Pendente", "Perda de arquivo", "Preenchimento incorreto da pauta", "Cancelamento de exames", "Candidato não assinou Termo de Orientação", "Fiscal não conferiu documento oficial com foto", "Exame iniciou com atraso", "Exame iniciou antes do horário agendado", "Recurso de questão não enviado no prazo",
               "Fiscal forneceu documento de outro OC", "Não forneceu as instruções iniciais", "Exame não foi realizado", "Fiscal aplicou uma prova cancelada no sistema", "Problemas Técnicos no CT", "Postura inadequada do fiscal", "Liberação do candidato sem finalização da prova", "Cancelamento em cima da hora", "Não cadastro de feriados e bloqueios", "Problema no CT – Data Errada nos Micros"]
lista_media = ["Não enviou os termos de orientação", "Não enviou o relatório do último mês",
               "Arquivo não enviado no final do horário de prova"]
lista_leve = ["Alteração de endereço"]
lista_ano = ["2021", "2022", "2023"]
grave = 10
media = 5
leve = 2
# Browser


class App:
    def make_chrome_browser(self, *options: str) -> webdriver.Chrome:
        chrome_options = webdriver.ChromeOptions()

        # chrome_options.add_argument('--headless')
        if options is not None:
            for option in options:
                chrome_options.add_argument(option)

        chrome_service = Service(
            executable_path="file_path"
        )

        browser = webdriver.Chrome(
            service=chrome_service,
            options=chrome_options

        )
        a.browser = browser
        return browser

    def start_browser(self):
        browser_confirm = tk.Label(
            root, text='Abrindo CertPessoas...', fg='red')
        browser_confirm.place(x=300, y=380, width=200, height=25)
        root.update()
        if __name__ == '__main__':
            a.options = ('--disable-gpu', '--no-sandbox',)
            a.browser = a.make_chrome_browser(*a.options)
            a.browser.get('url')

    def display_selected(self, ct_selected):
        self.ct_selected = variable.get()
        a.ct_selected = str(ct_selected)
        ct_confirm = tk.Label(
            root, text=f'CT selecionado com sucesso!', fg='green')
        ct_confirm.place(x=300, y=220, width=200, height=25)

    def month_selected(self, month_get):
        month_get = variable_month.get()
        a.month_get = str(month_get)
        ct_confirm = tk.Label(
            root, text=f'Mês selecionado com sucesso!', fg='green')
        ct_confirm.place(x=300, y=300, width=200, height=25)

    def ct_check(self):
        if a.ct_selected == 'CT 1':
            id_ct = 7
            a.id_ct = id_ct
        elif a.ct_selected == 'CT 2':
            id_ct = 8
            a.id_ct = id_ct
        elif a.ct_selected == 'CT 3':
            id_ct = 18
            a.id_ct = id_ct
        elif a.ct_selected == 'CT 4':
            id_ct = 20
            a.id_ct = id_ct

    def month_check(self):
        if a.month_get == "Janeiro":
            a.id_month = 2
        elif a.month_get == "Fevereiro":
            a.id_month = 3
        elif a.month_get == "Março":
            a.id_month = 4
        elif a.month_get == "Abril":
            a.id_month = 5
        elif a.month_get == "Maio":
            a.id_month = 6
        elif a.month_get == "Junho":
            a.id_month = 7
        elif a.month_get == "Julho":
            a.id_month = 8
        elif a.month_get == "Agosto":
            a.id_month = 9
        elif a.month_get == "Setembro":
            a.id_month = 10
        elif a.month_get == "Outubro":
            a.id_month = 11
        elif a.month_get == "Novembro":
            a.id_month = 12
        elif a.month_get == "Dezembro":
            a.id_month = 13

    def cert_login(self):
        browser_confirm = tk.Label(
            root, text='Efetuando login em CertPessoas...', fg='red')
        browser_confirm.place(x=300, y=380, width=200, height=25)
        root.update()
        inputlogin = a.browser.find_element(By.NAME, 'UserName')
        inputlogin.send_keys('user')
        inputpasswd = a.browser.find_element(By.NAME, 'Password')
        inputpasswd.send_keys('password')
        sleep(2)
        inputlogin.send_keys(Keys.ENTER)

    def switch_page(self):
        browser_confirm = tk.Label(root, text='Alternando página...', fg='red')
        browser_confirm.place(x=300, y=380, width=200, height=25)
        root.update()
        a.browser.get(
            "url")

    def select_filtro_single(self):
        browser_confirm = tk.Label(
            root, text='Selecionando filtros...', fg='red')
        browser_confirm.place(x=300, y=380, width=200, height=25)
        root.update()
        ct_cadastro_open = a.browser.find_element(
            By.XPATH, '//*[@id="ctl00_ContentMain_IDCentroTeste"]')
        drop = Select(ct_cadastro_open)
        drop.select_by_visible_text(f'{a.ct_selected}')
        mes_cadastro = a.browser.find_element(
            By.XPATH, f'//*[@id="ctl00_ContentMain_dropMes"]/option[{a.id_month}]').click()
        buscar_ocorrencia = a.browser.find_element(
            By.XPATH, '//*[@id="ctl00_ContentMain_btnProcurar"]').click()

    def select_filtro_multi(self):
        mes_cadastro = a.browser.find_element(
            By.XPATH, f'//*[@id="ctl00_ContentMain_dropMes"]/option[{a.id_month}]').click()
        # open_ano = a.browser.find_element(By.XPATH,f'//*[@id="ctl00_ContentMain_dropAno"]').click()
        # ano_cadastro =a.browser.find_element(By.XPATH, f'//*[@id="ctl00_ContentMain_dropAno"]/option[2]').click()
        # ano_cadastro.select_by_visible_text("2021")

    def loop_multi(self):
        browser_confirm = ttk.Progressbar(
            root, orient=HORIZONTAL, length=100, mode='determinate')
        browser_confirm.place(x=300, y=380, width=200, height=25)
        start_value = 100/694
        progress_value = start_value
        for x in range(2, 694, 1):
            browser_confirm['value'] = progress_value
            ct_cadastro = a.browser.find_element(
                By.XPATH, f'//*[@id="ctl00_ContentMain_IDCentroTeste"]/option[{x}]').text
            print(ct_cadastro)
            ct_cadastro = a.browser.find_element(
                By.XPATH, f'//*[@id="ctl00_ContentMain_IDCentroTeste"]/option[{x}]').click()
            mes_cadastro = a.browser.find_element(
                By.XPATH, f'//*[@id="ctl00_ContentMain_dropMes"]/option[{a.id_month}]').click()
            ano_cadastro = a.browser.find_element(
                By.XPATH, f'//*[@id="ctl00_ContentMain_dropAno"]/option[1]').click()
            buscar_ocorrencia = a.browser.find_element(
                By.XPATH, '//*[@id="ctl00_ContentMain_btnProcurar"]').click()
            for i in range(2, 11, 1):
                i = str(i)
                if len(i) < 2:
                    i = i.zfill(2)
                try:
                    date_report = a.browser.find_element(
                        By.XPATH, f'//*[@id="ctl00_ContentMain_gvAcontecimento_ctl{i}_lblData"]').text
                    tipo_report = a.browser.find_element(
                        By.XPATH, f'//*[@id="ctl00_ContentMain_gvAcontecimento_ctl{i}_lblTipo"]').text
                    ct_report = a.browser.find_element(
                        By.XPATH, f'//*[@id="ctl00_ContentMain_gvAcontecimento_ctl{i}_lblCentroTeste"]').text
                    desc_report = a.browser.find_element(
                        By.XPATH, f'//*[@id="ctl00_ContentMain_gvAcontecimento"]/tbody/tr[{i}]/td[5]').text
                    user_report = a.browser.find_element(
                        By.XPATH, f'//*[@id="ctl00_ContentMain_gvAcontecimento_ctl{i}_lblUsuario"]').text
                    tipo_report = str(tipo_report)
                    if tipo_report in lista_leve or tipo_report in lista_media or tipo_report in lista_grave:
                        lista_nome.append(ct_report)
                        lista_date.append(date_report)
                        lista_report.append(tipo_report)
                        lista_describ.append(desc_report)
                        lista_user.append(user_report)
                        if tipo_report in lista_grave:
                            classif = grave
                            lista_classif.append(classif)
                        if tipo_report in lista_media:
                            classif = media
                            lista_classif.append(classif)
                        if tipo_report in lista_leve:
                            classif = leve
                            lista_classif.append(classif)

                except:
                    pass

            progress_value = progress_value + start_value
            root.update()

        browser_confirm = tk.Label(
            root, text='Processo concluido!', fg='green')
        browser_confirm.place(x=300, y=380, width=200, height=25)

    def criar_planilha(self):
        dict = {"CT": lista_nome, "Data": lista_date, "Ocorrência": lista_report,
                "Classificação": lista_classif, "Descrição": lista_describ, "Usuário": lista_user}
        df = pd.DataFrame(dict)
        df.to_excel("Reportdata.xlsx")

    def start_process_single(self):
        a.start_browser()
        a.cert_login()
        a.switch_page()
        a.month_check()
        a.ct_check()
        a.select_filtro_single()

    def start_process_multi(self):
        a.start_browser()
        a.cert_login()
        a.switch_page()
        a.month_check()
        a.select_filtro_multi()
        a.loop_multi()
        a.criar_planilha()
        a.browser.quit()


a = App()
# Interface
root = Tk()
root.title("FGV Cert")
root.iconbitmap('fgv-logo-1.png')
root.geometry('800x600')
# Title
image_logo = Image.open("file_path")
fgv_logo = ImageTk.PhotoImage(image_logo)
label1 = tk.Label(image=fgv_logo)
label1.image = fgv_logo
label1.place(x=50, y=0, width=700, height=200)

# Escolher ct
variable = StringVar(root)
variable.set("Centro de Testes")  # default value
ct_opt = OptionMenu(root, variable, *lista_ct, command=a.display_selected)
ct_opt.place(x=200, y=180, width=400, height=30)

# Escolher mês
variable_month = StringVar(root)
variable_month.set("Mês")  # default value
mes_opt = OptionMenu(root, variable_month, *lista_mes,
                     command=a.month_selected)
mes_opt.place(x=320, y=260, width=150, height=30)


# Começar processo individual
start_btn = Button(root, text='Coletar dado', command=a.start_process_single)
start_btn.place(x=240, y=340, width=150, height=30)

# Começar processo multi
start2_btn = Button(root, text='Coletar todos os dados',
                    command=a.start_process_multi)
start2_btn.place(x=420, y=340, width=150, height=30)

root.mainloop()
