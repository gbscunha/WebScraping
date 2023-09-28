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
    'AGEERP Treinamentos & Estágios - Ribeirão Preto - SP',
    'AGEERP Treinamentos & Estágios II - Ribeirão Preto - SP',
    'Ápice Treinamentos - Pouso Alegre - MG',
    'Aprenda Cursos - Jataí - GO',
    'Aprendiz Projetos e Treinamentos - Goiânia - GO',
    'Be Mersion Cursos - Marília - SP',
    'BIG Cursos - Fernandópolis - SP',
    'BIT - Araxá - MG',
    'Bit Mais - Uberaba - MG',
    'CD6 Centro de Desenvolvimento - Curitiba - PR',
    'CD6 Centro de Desenvolvimento III - Curitiba - PR',
    'CENAIC - Adamantina - SP',
    'CENAIC - Osvaldo Cruz - SP',
    'CENAIC - Tupã - SP',
    'Compuway Ensino Profissional - Santo André - SP',
    'Compuway Formação Profissional - Sobral - CE',
    'Conecta Flex - Dourados - MS',
    'Conexão - Guarulhos - SP',
    'CRF EDUCAÇÃO - Votuporanga - SP',
    'Cruzeiro do Sul virtual - Afogados da Ingazeira - PE',
    'Cyber Informática e Idiomas - Barbacena - MG',
    'Data Byte - Belo Horizonte - MG',
    'Data Byte Filial - Belo Horizonte - MG',
    'Data Byte II - Belo Horizonte - MG',
    'DATACOMPANY - Jundiaí - SP',
    'Datamais Treinamento - Erechim - RS',
    'Decision - Porto Alegre - RS',
    'Digicenter Treinamentos - São Cristóvão - Chapecó - SC',
    'Docvision - Londrina - PR',
    'Easycomp Informática - Bauru - SP',
    'Educatec Brasil - Marabá - PA',
    'ENGAGE EDUC RH - Juazeiro - BA',
    'Escola Ponto Com Informática - Rio Branco - AC',
    'Escola Renascentista - São José do Rio Preto - SP',
    'Escola SAGA - Salvador - BA',
    'Escola Técnica Data Way - Campinas - SP',
    'FAAHF - Luís Eduardo Magalhães - BA',
    'Faculdade de Rolim de Moura - Farol - Rolim de Moura - RO',
    'Faculdade FAEL - Montes Claros - MG',
    'Faculdade Malta - Teresina - PI',
    'FACULDADES DECISION - Florianópolis - SC',
    'Fênix Centro de Ensino - Juiz de Fora - MG',
    'FGV - Brasília - DF',
    'FGV - Treze de Maio - Rio de Janeiro - RJ',
    'Fleek Tech And English - Divinópolis - MG',
    'Flexxo - Caxias do Sul - RS',
    'Grupo Help! Educação Profissional e Inglês - Resende - RJ',
    'IB Consulting / FGV - Palmas - TO',
    'Ideal FGV - Belém - PA',
    'Ideal Qualificação Profissional - S.J. da Boa Vista - SP',
    'Informática Futura - Cachoeiro de Itapemirim - ES',
    'Information - Bacabal - MA',
    'Instituto Máximo - Patos de Minas - MG',
    'Instituto Mix de Profissões - Anápolis - GO',
    'Instituto Mix de Profissões - Blumenau - SC',
    'Instituto Mix de Profissões - Fortaleza - CE',
    'Instituto Mix de Profissões - Porto Velho - RO',
    'Instituto Mix de Profissões - Santa Cruz do Sul - RS',
    'Instituto Tayano de Educação - Tangará da Serra - MT',
    'ITSP / Instituto Tecnológico São Pedro - Bragança Paulista - SP',
    'Jumper - Profissões e Idiomas - Joinville - SC',
    'KNN Idiomas - Varginha - MG',
    'LANGUAGES IDIOMAS - Assis - SP',
    'LOGOS Escola de Informática - Governador Valadares - MG',
    'LOGUS Informática - Araguaína - TO',
    'M.Murad - Vitória - ES',
    'Mais Futuro Central de Cursos - Santa Rosa - RS',
    'Maximize Conhecimento e Tecnologia - Salvador - BA',
    'Mbyte Informática - Santos - SP',
    'Mega Smart - Profissões e Idiomas - Ponta Grossa - PR',
    'Metropolitana Ensino Especializado - Avaré - SP',
    'Micro Fácil - Itumbiara - GO',
    'Microlins - Aracaju - SE',
    'Microlins - Barra do Garças - MT',
    'Microlins - Barreiras - BA',
    'Microlins - Campinas - SP',
    'Microlins - Campos dos Goytacazes - RJ',
    'Microlins - Colatina - ES',
    'Microlins - Itabuna - BA',
    'Microlins - Jequié - BA',
    'Microlins - Juazeiro do Norte - CE',
    'Microlins - Macaé - RJ',
    'Microlins - Manaus - AM',
    'Microlins - Manaus - Centro - AM',
    'Microlins - Natal - RN',
    'Microlins - Nova Friburgo - RJ',
    'Microlins - Parnaíba - PI',
    'Microlins - Piracicaba - SP',
    'Microlins - Poços de Caldas - MG',
    'Microlins - Presidente Prudente - SP',
    'Microlins - Santa Maria da Vitória - BA',
    'Microlins - Santarém - PA',
    'Microlins - São Carlos - SP',
    'Microlins - São José do Rio Preto - SP',
    'Microlins - Sinop - MT',
    'Microlins - Sorocaba - SP',
    'Microlins - Sorriso - MT',
    'Microlins - Vitória da Conquista - BA',
    'Microway Cursos e Treinamento - Balneário Camboriú - SC',
    'Microway Cursos e Treinamento - Barretos - SP',
    'Milenium Informática - Campo Grande - MS',
    'MRH Gestão - Fortaleza - CE',
    'MultiEduc Cursos Profissionalizantes - Caraguatatuba - SP',
    'Naja Informática - Viçosa - MG',
    'Netcom - São Luís - MA',
    'New Life Cursos - Ijuí - RS',
    'Ômega Cursos Profissionalizantes - Francisco Beltrão - PR',
    'On Byte Formação Profissional - São José dos Campos - SP',
    'On Byte Formação Profissional - Taguatinga - DF',
    'People Formação Completa - Petrópolis - RJ',
    'PEOPLE TECH AND ENGLISH - Rio das Ostras - RJ',
    'Play Educação - Arapiraca - AL',
    'Play Educação - Unidade Farol - Maceió - AL',
    'Prepara Cursos Icoaraci - Belém - PA',
    'Prepara Cursos Profissionalizantes - Boa Vista - RR',
    'Prepara Cursos Profissionalizantes - Conselheiro Lafaiete - MG',
    'Prepara Cursos Profissionalizantes - Toledo - PR',
    'Pro Job Treinamentos - Criciúma - SC',
    'Proativa Orientação Profissional - Limeira - SP',
    'Proene - Teofilo Otoni - MG',
    'ProWay IT Training - Blumenau - SC',
    'Qualitec Cursos - Recife - PE',
    'Sebratep Educacional - Joaçaba - SC',
    'SEJA PAIDEIA/ESTÁGIO FÁCIL - Brasnorte - MT',
    'Sematec - Florianópolis - SC',
    'SKILL Idiomas - Franca - SP',
    'Smart Informática - Betim - MG',
    'Soluções Educacionais / FGV - Macapá - AP',
    'Star Cursos e Treinamentos - Lins - SP',
    'Strong - Osasco - SP',
    'Strong - Santo André - SP',
    'Strong - Santos - SP',
    'SUPERAÇÃO - Cruz Alta - RS',
    'SVA Haddock Lobo - São Paulo - SP',
    'Trecsson - Dourados - MS',
    'Trecsson Business - Maringá - PR',
    'UDC Monjolo - Foz do Iguaçu - PR',
    'UNICESUMAR - Rio do Sul - SC',
    'UNICESUMAR - Terra Rica - PR',
    'UNICESUMAR - Vitória - ES',
    'UniDomBosco - Umuarama - PR',
    'UNINASSAU - Uberlândia - MG',
    'Uninter - Altamira - PA',
    'UNINTER - Barracão - PR',
    'UNINTER - Campo Grande - MS',
    'UNINTER - Imperatriz - MA',
    'UNINTER - Niterói - RJ',
    'UNINTER - Presidente Prudente - SP',
    'UNINTER - Santa Cruz do Sul - RS',
    'Uninter / LG Assessoria - São Miguel do Oeste - SC',
    'UNINTER/MCM EDUCACIONAL - Xanxerê - SC',
    'UNIP - Cuiabá - MT',
    'UNIVEL - Cascavel - PR',
    'UNOPAR - Caçador - SC',
    'Visual Midia - Caruaru - PE',
    'VitalNet - Guarapuava - PR',
    'Way Soluções Profissionais - Caratinga - MG',
    'WKS Informática - Ipatinga - MG',
    'Yázigi Tambaú - João Pessoa - PB'

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
        if a.ct_selected == 'AGEERP Treinamentos & Estágios - Ribeirão Preto - SP':
            id_ct = 7
            a.id_ct = id_ct
        elif a.ct_selected == 'AGEERP Treinamentos & Estágios II - Ribeirão Preto - SP':
            id_ct = 8
            a.id_ct = id_ct
        elif a.ct_selected == 'Ápice Treinamentos - Pouso Alegre - MG':
            id_ct = 18
            a.id_ct = id_ct
        elif a.ct_selected == 'Aprenda Cursos - Jataí - GO':
            id_ct = 20
            a.id_ct = id_ct
        elif a.ct_selected == 'Aprendiz Projetos e Treinamentos - Goiânia - GO':
            id_ct = 21
            a.id_ct = id_ct
        elif a.ct_selected == 'Be Mersion Cursos - Marília - SP':
            id_ct = 29
            a.id_ct = id_ct
        elif a.ct_selected == 'BIG Cursos - Fernandópolis - SP':
            id_ct = 30
            a.id_ct = id_ct
        elif a.ct_selected == 'BIT - Araxá - MG':
            id_ct = 32
            a.id_ct = id_ct
        elif a.ct_selected == 'Bit Mais - Uberaba - MG':
            id_ct = 37
            a.id_ct = id_ct
        elif a.ct_selected == 'CD Centro de Desenvolvimento - Curitiba - PR':
            id_ct = 45
            a.id_ct = id_ct
        elif a.ct_selected == 'CD Centro de Desenvolvimento III - Curitiba - PR':
            id_ct = 47
            a.id_ct = id_ct
        elif a.ct_selected == 'CENAIC - Adamantina - SP':
            id_ct = 52
            a.id_ct = id_ct
        elif a.ct_selected == 'CENAIC - Osvaldo Cruz - SP':
            id_ct = 54
            a.id_ct = id_ct
        elif a.ct_selected == 'CENAIC - Taquarituba - SP':
            id_ct = 56
            a.id_ct = id_ct
        elif a.ct_selected == 'CENAIC - Tupã - SP':
            id_ct = 58
            a.id_ct = id_ct
        elif a.ct_selected == 'Compuway Ensino Profissional - Santo André - SP':
            id_ct = 89
            a.id_ct = id_ct
        elif a.ct_selected == 'Compuway Formação Profissional - Sobral - CE':
            id_ct = 90
            a.id_ct = id_ct
        elif a.ct_selected == 'Conecta Flex - Dourados - MS':
            id_ct = 92
            a.id_ct = id_ct
        elif a.ct_selected == 'CRF EDUCAÇÃO - Votuporanga - SP':
            id_ct = 99
            a.id_ct = id_ct
        elif a.ct_selected == 'Cruzeiro do Sul virtual - Afogados da Ingazeira - PE':
            id_ct = 100
            a.id_ct = id_ct
        elif a.ct_selected == 'Cyber Informática e Idiomas - Barbacena - MG':
            id_ct = 107
            a.id_ct = id_ct
        elif a.ct_selected == 'Data Byte - Belo Horizonte - MG':
            id_ct = 111
            a.id_ct = id_ct
        elif a.ct_selected == 'Data Byte Filial - Belo Horizonte - MG':
            id_ct = 114
            a.id_ct = id_ct
        elif a.ct_selected == 'Data Byte II - Belo Horizonte - MG':
            id_ct = 115
            a.id_ct = id_ct
        elif a.ct_selected == 'DATACOMPANY - Jundiaí - SP':
            id_ct = 116
            a.id_ct = id_ct
        elif a.ct_selected == 'Datamais Treinamento - Erechim - RS':
            id_ct = 117
            a.id_ct = id_ct
        elif a.ct_selected == 'Decision - Porto Alegre - RS':
            id_ct = 118
            a.id_ct = id_ct
        elif a.ct_selected == 'Digicenter Treinamentos - São Cristóvão - Chapecó - SC':
            id_ct = 121
            a.id_ct = id_ct
        elif a.ct_selected == 'Docvision - Londrina - PR':
            id_ct = 122
            a.id_ct = id_ct
        elif a.ct_selected == 'Easycomp Informática - Bauru - SP':
            id_ct = 129
            a.id_ct = id_ct
        elif a.ct_selected == 'Educatec Brasil - Marabá - PA':
            id_ct = 134
            a.id_ct = id_ct
        elif a.ct_selected == 'ENGAGE EDUC RH - Juazeiro - BA':
            id_ct = 137
            a.id_ct = id_ct
        elif a.ct_selected == 'Escola Ponto Com Informática - Rio Branco - AC':
            id_ct = 144
            a.id_ct = id_ct
        elif a.ct_selected == 'Escola Renascentista - São José do Rio Preto - SP':
            id_ct = 145
            a.id_ct = id_ct
        elif a.ct_selected == 'Escola SAGA - Salvador - BA':
            id_ct = 146
            a.id_ct = id_ct
        elif a.ct_selected == 'Escola Técnica Data Way - Campinas - SP':
            id_ct = 148
            a.id_ct = id_ct
        elif a.ct_selected == 'FAAHF - Luís Eduardo Magalhães - BA':
            id_ct = 159
            a.id_ct = id_ct
        elif a.ct_selected == 'Faculdade de Rolim de Moura - Farol - Rolim de Moura - RO':
            id_ct = 164
            a.id_ct = id_ct
        elif a.ct_selected == 'Faculdade FAEL - Montes Claros - MG':
            id_ct = 166
            a.id_ct = id_ct
        elif a.ct_selected == 'Faculdade Malta - Teresina - PI':
            id_ct = 169
            a.id_ct = id_ct
        elif a.ct_selected == 'FACULDADES DECISION - Florianópolis - SC':
            id_ct = 175
            a.id_ct = id_ct
        elif a.ct_selected == 'Fênix Centro de Ensino - Juiz de Fora - MG':
            id_ct = 187
            a.id_ct = id_ct
        elif a.ct_selected == 'FGV - Brasília - DF':
            id_ct = 194
            a.id_ct = id_ct
        elif a.ct_selected == 'FGV - Treze de Maio - Rio de Janeiro - RJ':
            id_ct = 202
            a.id_ct = id_ct
        elif a.ct_selected == 'Fleek Tech And English - Divinópolis - MG':
            id_ct = 206
            a.id_ct = id_ct
        elif a.ct_selected == 'Flexxo - Caxias do Sul - RS':
            id_ct = 208
            a.id_ct = id_ct
        elif a.ct_selected == 'Grupo Help! Educação Profissional e Inglês - Resende - RJ':
            id_ct = 221
            a.id_ct = id_ct
        elif a.ct_selected == 'IB Consulting / FGV - Palmas - TO':
            id_ct = 231
            a.id_ct = id_ct
        elif a.ct_selected == 'Ideal FGV - Belém - PA':
            id_ct = 236
            a.id_ct = id_ct
        elif a.ct_selected == 'Ideal Qualificação Profissional - S.J. da Boa Vista - SP':
            id_ct = 237
            a.id_ct = id_ct
        elif a.ct_selected == 'Informática Futura - Cachoeiro de Itapemirim - ES':
            id_ct = 256
            a.id_ct = id_ct
        elif a.ct_selected == 'Information - Bacabal - MA':
            id_ct = 257
            a.id_ct = id_ct
        elif a.ct_selected == 'Instituto Máximo - Patos de Minas - MG':
            id_ct = 264
            a.id_ct = id_ct
        elif a.ct_selected == 'Instituto Mix de Profissões - Anápolis - GO':
            id_ct = 266
            a.id_ct = id_ct
        elif a.ct_selected == 'Instituto Mix de Profissões - Blumenau - SC':
            id_ct = 267
            a.id_ct = id_ct
        elif a.ct_selected == 'Instituto Mix de Profissões - Fortaleza - CE':
            id_ct = 269
            a.id_ct = id_ct
        elif a.ct_selected == 'Instituto Mix de Profissões - Porto Velho - RO':
            id_ct = 274
            a.id_ct = id_ct
        elif a.ct_selected == 'Instituto Mix de Profissões - Santa Cruz do Sul - RS':
            id_ct = 275
            a.id_ct = id_ct
        elif a.ct_selected == 'Instituto Tayano de Educação - Tangará da Serra - MT':
            id_ct = 280
            a.id_ct = id_ct
        elif a.ct_selected == 'ITSP / Instituto Tecnológico São Pedro - Bragança Paulista - SP':
            id_ct = 288
            a.id_ct = id_ct
        elif a.ct_selected == 'Jumper - Profissões e Idiomas - Joinville - SC':
            id_ct = 290
            a.id_ct = id_ct
        elif a.ct_selected == 'KNN Idiomas - Varginha - MG':
            id_ct = 295
            a.id_ct = id_ct
        elif a.ct_selected == 'LANGUAGES IDIOMAS - Assis - SP':
            id_ct = 296
            a.id_ct = id_ct
        elif a.ct_selected == 'LOGOS Escola de Informática - Governador Valadares - MG':
            id_ct = 301
            a.id_ct = id_ct
        elif a.ct_selected == 'LOGUS Informática - Araguaína - TO':
            id_ct = 302
            a.id_ct = id_ct
        elif a.ct_selected == 'M.Murad - Vitória - ES':
            id_ct = 303
            a.id_ct = id_ct
        elif a.ct_selected == 'Mais Futuro Central de Cursos - Santa Rosa - RS':
            id_ct = 305
            a.id_ct = id_ct
        elif a.ct_selected == 'Maximize Conhecimento e Tecnologia - Salvador - BA':
            id_ct = 307
            a.id_ct = id_ct
        elif a.ct_selected == 'Mbyte Informática - Santos - SP':
            id_ct = 309
            a.id_ct = id_ct
        elif a.ct_selected == 'Mega Smart - Profissões e Idiomas - Ponta Grossa - PR':
            id_ct = 313
            a.id_ct = id_ct
        elif a.ct_selected == 'Metropolitana Ensino Especializado - Avaré - SP':
            id_ct = 317
            a.id_ct = id_ct
        elif a.ct_selected == 'Micro Fácil - Itumbiara - GO':
            id_ct = 318
            a.id_ct = id_ct
        elif a.ct_selected == 'Microlins - Aracaju - SE':
            id_ct = 342
            a.id_ct = id_ct
        elif a.ct_selected == 'Microlins - Barra do Garças - MT':
            id_ct = 345
            a.id_ct = id_ct
        elif a.ct_selected == 'Microlins - Barreiras - BA':
            id_ct = 346
            a.id_ct = id_ct
        elif a.ct_selected == 'Microlins - Campinas - SP':
            id_ct = 351
            a.id_ct = id_ct
        elif a.ct_selected == 'Microlins - Campos dos Goytacazes - RJ':
            id_ct = 353
            a.id_ct = id_ct
        elif a.ct_selected == 'Microlins - Colatina - ES':
            id_ct = 358
            a.id_ct = id_ct
        elif a.ct_selected == 'Microlins - Itabuna - BA':
            id_ct = 370
            a.id_ct = id_ct
        elif a.ct_selected == 'Microlins - Jequié - BA':
            id_ct = 375
            a.id_ct = id_ct
        elif a.ct_selected == 'Microlins - Juazeiro do Norte - CE':
            id_ct = 377
            a.id_ct = id_ct
        elif a.ct_selected == 'Microlins - Macaé - RJ':
            id_ct = 381
            a.id_ct = id_ct
        elif a.ct_selected == 'Microlins - Manaus - AM':
            id_ct = 382
            a.id_ct = id_ct
        elif a.ct_selected == 'Microlins - Manaus - Centro - AM':
            id_ct = 383
            a.id_ct = id_ct
        elif a.ct_selected == 'Microlins - Natal - RN':
            id_ct = 387
            a.id_ct = id_ct
        elif a.ct_selected == 'Microlins - Nova Friburgo - RJ':
            id_ct = 389
            a.id_ct = id_ct
        elif a.ct_selected == 'Microlins - Parnaíba - PI':
            id_ct = 393
            a.id_ct = id_ct
        elif a.ct_selected == 'Microlins - Piracicaba - SP':
            id_ct = 397
            a.id_ct = id_ct
        elif a.ct_selected == 'Microlins - Poços de Caldas - MG':
            id_ct = 398
            a.id_ct = id_ct
        elif a.ct_selected == 'Microlins - Presidente Prudente - SP':
            id_ct = 401
            a.id_ct = id_ct
        elif a.ct_selected == 'Microlins - Santa Maria da Vitória - BA':
            id_ct = 405
            a.id_ct = id_ct
        elif a.ct_selected == 'Microlins - Santarém - PA':
            id_ct = 407
            a.id_ct = id_ct
        elif a.ct_selected == 'Microlins - São Carlos - SP':
            id_ct = 408
            a.id_ct = id_ct
        elif a.ct_selected == 'Microlins - São José do Rio Preto - SP':
            id_ct = 409
            a.id_ct = id_ct
        elif a.ct_selected == 'Microlins - Sinop - MT':
            id_ct = 412
            a.id_ct = id_ct
        elif a.ct_selected == 'Microlins - Sorocaba - SP':
            id_ct = 414
            a.id_ct = id_ct
        elif a.ct_selected == 'Microlins - Sorriso - MT':
            id_ct = 415
            a.id_ct = id_ct
        elif a.ct_selected == 'Microlins - Vitória da Conquista - BA':
            id_ct = 420
            a.id_ct = id_ct
        elif a.ct_selected == 'Microway Cursos e Treinamento - Balneário Camboriú - SC':
            id_ct = 430
            a.id_ct = id_ct
        elif a.ct_selected == 'Microway Cursos e Treinamento - Barretos - SP':
            id_ct = 431
            a.id_ct = id_ct
        elif a.ct_selected == 'Milenium Informática - Campo Grande - MS':
            id_ct = 433
            a.id_ct = id_ct
        elif a.ct_selected == 'MRH Gestão - Fortaleza - CE':
            id_ct = 436
            a.id_ct = id_ct
        elif a.ct_selected == 'MultiEduc Cursos Profissionalizantes - Caraguatatuba - SP':
            id_ct = 441
            a.id_ct = id_ct
        elif a.ct_selected == 'Naja Informática - Viçosa - MG':
            id_ct = 445
            a.id_ct = id_ct
        elif a.ct_selected == 'Netcom - São Luís - MA':
            id_ct = 451
            a.id_ct = id_ct
        elif a.ct_selected == 'New Life Cursos - Ijuí - RS':
            id_ct = 454
            a.id_ct = id_ct
        elif a.ct_selected == 'Ômega Cursos Profissionalizantes - Francisco Beltrão - PR':
            id_ct = 470
            a.id_ct = id_ct
        elif a.ct_selected == 'On Byte Formação Profissional - São José dos Campos - SP':
            id_ct = 472
            a.id_ct = id_ct
        elif a.ct_selected == 'On Byte Formação Profissional - Taguatinga - DF':
            id_ct = 473
            a.id_ct = id_ct
        elif a.ct_selected == 'People Formação Completa - Petrópolis - RJ':
            id_ct = 484
            a.id_ct = id_ct
        elif a.ct_selected == 'PEOPLE TECH AND ENGLISH - Rio das Ostras - RJ':
            id_ct = 486
            a.id_ct = id_ct
        elif a.ct_selected == 'Play Educação - Arapiraca - AL':
            id_ct = 489
            a.id_ct = id_ct
        elif a.ct_selected == 'Play Educação - Unidade Farol - Maceió - AL':
            id_ct = 490
            a.id_ct = id_ct
        elif a.ct_selected == 'Prepara Cursos Icoaraci - Belém - PA':
            id_ct = 497
            a.id_ct = id_ct
        elif a.ct_selected == 'Prepara Cursos Profissionalizantes - Boa Vista - RR':
            id_ct = 498
            a.id_ct = id_ct
        elif a.ct_selected == 'Prepara Cursos Profissionalizantes - Conselheiro Lafaiete - MG':
            id_ct = 504
            a.id_ct = id_ct
        elif a.ct_selected == 'Prepara Cursos Profissionalizantes - Toledo - PR':
            id_ct = 518
            a.id_ct = id_ct
        elif a.ct_selected == 'Pro Job Treinamentos - Criciúma - SC':
            id_ct = 519
            a.id_ct = id_ct
        elif a.ct_selected == 'Proativa Orientação Profissional - Limeira - SP':
            id_ct = 520
            a.id_ct = id_ct
        elif a.ct_selected == 'Proene - Teofilo Otoni - MG':
            id_ct = 521
            a.id_ct = id_ct
        elif a.ct_selected == 'ProWay IT Training - Blumenau - SC':
            id_ct = 522
            a.id_ct = id_ct
        elif a.ct_selected == 'Qualitec Cursos - Recife - PE':
            id_ct = 524
            a.id_ct = id_ct
        elif a.ct_selected == 'Sebratep Educacional - Joaçaba - SC':
            id_ct = 533
            a.id_ct = id_ct
        elif a.ct_selected == 'SEJA PAIDEIA/ESTÁGIO FÁCIL - Brasnorte - MT':
            id_ct = 534
            a.id_ct = id_ct
        elif a.ct_selected == 'Sematec - Florianópolis - SC':
            id_ct = 535
            a.id_ct = id_ct
        elif a.ct_selected == 'SKILL Idiomas - Franca - SP':
            id_ct = 549
            a.id_ct = id_ct
        elif a.ct_selected == 'Smart Informática - Betim - MG':
            id_ct = 550
            a.id_ct = id_ct
        elif a.ct_selected == 'Soluções Educacionais / FGV - Macapá - AP':
            id_ct = 555
            a.id_ct = id_ct
        elif a.ct_selected == 'Star Cursos e Treinamentos - Lins - SP':
            id_ct = 574
            a.id_ct = id_ct
        elif a.ct_selected == 'Strong - Osasco - SP':
            id_ct = 577
            a.id_ct = id_ct
        elif a.ct_selected == 'Strong - Santo André - SP':
            id_ct = 578
            a.id_ct = id_ct
        elif a.ct_selected == 'Strong - Santos - SP':
            id_ct = 579
            a.id_ct = id_ct
        elif a.ct_selected == 'SUPERAÇÃO - Cruz Alta - RS':
            id_ct = 583
            a.id_ct = id_ct
        elif a.ct_selected == 'SVA Haddock Lobo - São Paulo - SP':
            id_ct = 587
            a.id_ct = id_ct
        elif a.ct_selected == 'Trecsson - Dourados - MS':
            id_ct = 590
            a.id_ct = id_ct
        elif a.ct_selected == 'Trecsson Business - Maringá - PR':
            id_ct = 591
            a.id_ct = id_ct
        elif a.ct_selected == 'UDC Monjolo - Foz do Iguaçu - PR':
            id_ct = 599
            a.id_ct = id_ct
        elif a.ct_selected == 'UNICESUMAR - Rio do Sul - SC':
            id_ct = 611
            a.id_ct = id_ct
        elif a.ct_selected == 'UNICESUMAR - Terra Rica - PR':
            id_ct = 613
            a.id_ct = id_ct
        elif a.ct_selected == 'UNICESUMAR - Vitória - ES':
            id_ct = 614
            a.id_ct = id_ct
        elif a.ct_selected == 'UniDomBosco - Umuarama - PR':
            id_ct = 621
            a.id_ct = id_ct
        elif a.ct_selected == 'UNINASSAU - Uberlândia - MG':
            id_ct = 631
            a.id_ct = id_ct
        elif a.ct_selected == 'Uninter - Altamira - PA':
            id_ct = 632
            a.id_ct = id_ct
        elif a.ct_selected == 'UNINTER - Barracão - PR':
            id_ct = 633
            a.id_ct = id_ct
        elif a.ct_selected == 'UNINTER - Campo Grande - MS':
            id_ct = 635
            a.id_ct = id_ct
        elif a.ct_selected == 'UNINTER - Imperatriz - MA':
            id_ct = 638
            a.id_ct = id_ct
        elif a.ct_selected == 'UNINTER - Niterói - RJ':
            id_ct = 641
            a.id_ct = id_ct
        elif a.ct_selected == 'UNINTER - Presidente Prudente - SP':
            id_ct = 643
            a.id_ct = id_ct
        elif a.ct_selected == 'UNINTER - Santa Cruz do Sul - RS':
            id_ct = 646
            a.id_ct = id_ct
        elif a.ct_selected == 'Uninter / LG Assessoria - São Miguel do Oeste - SC':
            id_ct = 652
            a.id_ct = id_ct
        elif a.ct_selected == 'UNINTER/MCM EDUCACIONAL - Xanxerê - SC':
            id_ct = 654
            a.id_ct = id_ct
        elif a.ct_selected == 'UNIP - Cuiabá - MT':
            id_ct = 659
            a.id_ct = id_ct
        elif a.ct_selected == 'UNIVEL - Cascavel - PR':
            id_ct = 671
            a.id_ct = id_ct
        elif a.ct_selected == 'UNOPAR - Caçador - SC':
            id_ct = 675
            a.id_ct = id_ct
        elif a.ct_selected == 'Visual Midia - Caruaru - PE':
            id_ct = 696
            a.id_ct = id_ct
        elif a.ct_selected == 'VitalNet - Guarapuava - PR':
            id_ct = 697
            a.id_ct = id_ct
        elif a.ct_selected == 'Way Soluções Profissionais - Caratinga - MG':
            id_ct = 699
            a.id_ct = id_ct
        elif a.ct_selected == 'WKS Informática - Ipatinga - MG':
            id_ct = 702
            a.id_ct = id_ct
        elif a.ct_selected == 'World Informática - União da Vitória - PR':
            id_ct = 703
            a.id_ct = id_ct
        elif a.ct_selected == 'Yázigi Tambaú - João Pessoa - PB':
            id_ct = 705
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
