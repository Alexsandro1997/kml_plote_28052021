#Interface gráfica para gerar arquivo kml

#importando as bibliotecas necessárias
from tkinter import *
import pandas as pd
from tkinter import ttk
from tkinter import filedialog
import os
import keyboard

#Importando as funções para os botões
from functions import *  

global distancia



#Importar dados do arquivo de configuração.
config = pd.read_excel('config.xlsx')


#Lista de itens para o combox  self.combobox_titulo
lista_itens_titulos = []
itens_titulos = (config['titulos'])
itens_titulos = itens_titulos.dropna()

for i in itens_titulos:
    lista_itens_titulos.append(i)
try:
    lista_itens_titulos.sort()
except:
    pass
    print('Os valores da lista "lista_itens_titulos" são números, portanto não podem ser organizados por ordem alfabetica.')


#Lista de itens para o combox  self.combobox_marcadores
lista_itens_marcadores = []
itens_marcadores = (config['marcadores'])
itens_marcadores = itens_marcadores.dropna()

for i in itens_marcadores:
    lista_itens_marcadores.append(i)
try:
    lista_itens_marcadores.sort()
except:
    pass
    print('Os valores da lista "lista_itens_marcadores" são números, portanto não podem ser organizados por ordem alfabética.')



#Lista de itens para o combox  self.combobox_google_earth
lista_google_earth = []
itens_google_earth = (config['google_earth'])
itens_google_earth = itens_google_earth.dropna()

for i in itens_google_earth:
    lista_google_earth.append(i)
try:
    lista_google_earth.sort()
except:
    pass
    print('Os valores da lista "lista_google_earth" são números, portanto não podem ser organizados por ordem alfabética.')
    

lista_treeview_id = []
lista_treeview_titulos = []
lista_treeview_obs = []
lista_treeview_ini = []
lista_treeview_fim = []



#Lista de itens para o combox  self.combobox_utm_zonas
lista_utm_zonas= []
itens_utm_zonas = (config['utm_zona_d'])
itens_utm_zonas = itens_utm_zonas.dropna()

for i in itens_utm_zonas:
    lista_utm_zonas.append(i)
lista_utm_zonas.sort()



#Lista para as cores combobox
lista_cores = ('Amarelo', 'Azul', 'Verde', 'Vermelho', '')
lista_larguras_linha = (2, 4, 6, 8, 10, 12)

#Lista para as coordenadas
lista_coordenadas_lat = []
lista_coordenadas_lon = []



#Valores para controle
i = 0
contar_itens = 0



#Variável para invocar a função kml do simplekml
kml = simplekml.Kml()





class CriarKml:
    def __init__(self, criar_kml):
        self.myparent = criar_kml
        self.myparent.config(bg = 'light gray')
        self.style = ttk.Style(criar_kml)
        self.style.theme_use('clam')
        
        #Frame informativo
        self.frame1 = Frame(criar_kml)
        self.frame1.grid(row = 0, column = 0, columnspan = 2)
        self.frame1.config(bg = 'light gray')
        #Label para informar o que o usuário tem em mãos
        self.label1 = Label(self.frame1, text = 'Gerador de kml, abra o Google Earth antes de usar.', font = 'arial')
        self.label1.grid(row=0, column=0, padx = '1', pady='10')
        self.label1.config(bg = 'light gray')
#          #Informações solicitadas
#         self.label2 = Label(self.frame1, text = 'Preencha os campos abaixo com as informações adequadas.', font = 'arial')
#         self.label2.grid(row=1, column=0, padx = '1', pady='1')
                
        
        
        
        #LabelFrame para informação de configurações iniciais
        self.inf_config= LabelFrame(criar_kml, text = 'Configuração inicial', font = 'calibri')
        self.inf_config.grid(row=1, column=0, columnspan = 1, padx = '1', pady='5')
#         
#         #Informativo campo coordenada inicial Graus Decimais
#         self.label_inf = Label(self.inf_config, text = 'Tipo de coordenada:', font = 'arial')
#         self.label_inf.grid(row=0, column=0)
#         
#         #Checkbutton para a opção de coordenadas UTM
        self.coord_utm = IntVar()
        self.inf_utm = Checkbutton(self.inf_config, text = 'UTM', variable = self.coord_utm, onvalue = 1, offvalue = 0, height = 1, width = 5, font = 'arial', relief = FLAT)
#         self.inf_utm.grid(row=0, column=1)
#         
#         #Checkbutton para a opção de coordenadas Graus Decimais
        self.coord_gd = IntVar()
        self.inf_gd = Checkbutton(self.inf_config, text = 'Graus Decimais', variable = self.coord_gd, onvalue = 1, offvalue = 0, height = 1, width = 15, font = 'arial')
#         self.inf_gd.grid(row=0, column=2)
        
        
        
        #Informativo diretório para salvar os arquivos
        self.label_inf_dir = Label(self.inf_config, text = 'Salvar arquivos em:', font = 'arial')
        self.label_inf_dir.grid(row=0, column=0,  padx = '1', pady='1')

        #Diretório para salvar os arquivos
        self.dir_salvar = StringVar()
        self.dir_salvar = Entry(self.inf_config, textvariable = self.dir_salvar)
        self.dir_salvar.config(width = '50')
        self.dir_salvar.grid(row=0, column=3, padx = '5', pady='1')
        

        #Informativo diretório para o GoogleEarth
        self.label_inf = Label(self.inf_config, text = 'Diretório do GoogleEarth: ', font = 'arial')
        self.label_inf.grid(row=0, column=4,  padx = '5', pady='1')

        #Diretório para o GoogleEarth
        self.combobox_google_earth = StringVar()
        self.combobox_google_earth = ttk.Combobox(self.inf_config, textvariable = self.combobox_google_earth)
        self.combobox_google_earth.config(values = (lista_google_earth))
        self.combobox_google_earth.config(width = '55')
        self.combobox_google_earth.grid(row=0, column=5, padx = '7', pady='1')   
        

        #Informativo nome do arquivo
        self.label_inf_nome = Label(self.inf_config, text = 'Nome do arquivo:   ', font = 'arial')
        self.label_inf_nome.grid(row=1, column=0,  padx = '1', pady='1')

        #Nome do arquivo
        self.nome_arq = StringVar()
        self.nome_arq = Entry(self.inf_config, textvariable = self.nome_arq)
        self.nome_arq.config(width = '50')
        self.nome_arq.grid(row=1, column=1, columnspan = 3, padx = '5', pady='1')
        
        
        # Função para o botão self.button_inf_ini
        def inf_ini():
            print('inf_ini')
            inf_ini_f(self, criar_kml)                
                
        #Botão para informação sobre o campo de configurações iniciais
        self.button_inf_ini = Button(self.inf_config, text = '?', width = '1', height = '1', relief = GROOVE)
        self.button_inf_ini.config(bg = 'gold')
        self.button_inf_ini.config(command = inf_ini)
        self.button_inf_ini.grid(row=0, column = 6, padx = '5', pady = '1')

        def Enter_Infinf(criar_kml):
            print('Configurações iniciais!')
            self.button_inf_ini.config(relief = RAISED)
        self.button_inf_ini.bind('<Enter>', Enter_Infinf)
        
        def Leave_Infinf(criar_kml):
            self.button_inf_ini.config(relief = GROOVE)
        self.button_inf_ini.bind('<Leave>', Leave_Infinf)
            
        
        #Botão para confimar as informações iniciais
        def inf_iniciais():
            utm = self.coord_utm.get()
            gd = self.coord_gd.get()
                        
            if (utm == 1) and (gd == 1):
                print('Selecione apenas um tipo de coordenada.')
                self.inf_gd .deselect()
            elif utm == 1:
                print('self.coord_utm')
                self.inf_gd .deselect()
            elif gd == 1:
                print('self.coord_gd')
                self.inf_utm .deselect()
            else:
                print('Nenhum tipo de coordenada foi selecionado, portanto consideramos UTM.')
                self.inf_utm .select()



      # Funções para os botões relacionados às coordenadas em Graus Decimais       
            #Comando do botão salvar
            def salvar_gd():
                print('salvar_inf')
                salvar_gd_f(self, criar_kml)
            
            
            #Comando do botão exportar
            def exportar_kml():
                print('exportar_kml')
                exportar_gd_f(self, criar_kml)
                
    # Função para o botão 'self.button_inf_gd'
            def inf_gd():
                print('inf_gd')
                inf_gd_f(self, criar_kml)
                
                
     # Funções para os botões relacionados às coordenadas em UTM
            #Comando do botão salvar UTM
            def salvar_item():
                print('salvar_item')
                salvar_item_f(self, criar_kml)

                            
                
            #Comando do botão exportar UTM
            def exportar_kml():
                print('exportar_kml')
                exportar_utm_f(self, criar_kml)
            
    # Função para o botão 'self.button_inf_utm'
            def inf_utm():
                print('inf_utm')
                inf_utm_f(self, criar_kml)


            try:                
                #Variáveis de controle
                utm = self.coord_utm.get()                
                gd = self.coord_gd.get()
                
                #Caso exista a janela de cadastrar itens que ela seja destruída.
                try:
                    self.frame_label_gd.destroy()
                except:
                    pass
                
       
                #LabelFrame para informação de cadastros dos itens    
                self.inf_coord_lf = LabelFrame(criar_kml, text = 'Informações para cadastro', font = 'calibri')
                self.inf_coord_lf.grid(row=3, column=0, columnspan = 1, padx = '1', pady='5')

                #Frame informativo para as informações necessárias para execução do programa
                self.frame_label_inf = LabelFrame(self.inf_coord_lf, text = 'Atributos para o item', font = 'calibri', padx = '1', pady = '5')
                self.frame_label_inf.grid(row=0, column=0, padx = '5', pady='5')
                
                #Informativo campo coordenada inicial Graus Decimais
                self.label3 = Label(self.frame_label_inf, text = 'Início:  ', font = 'arial')
                self.label3.grid(row=0, column=0, padx = '1', pady='1')
                        
                #Informação para o nome        
                self.label7 = Label(self.frame_label_inf, text =  'Título:                       ', font = 'arial')
                self.label7.grid(row=0, column=0, padx='1', pady='1')
                
                #Informações para o título
                self.combobox_titulo = StringVar()
                self.combobox_titulo = ttk.Combobox(self.frame_label_inf, textvariable = self.combobox_titulo)
                self.combobox_titulo.config(values = (lista_itens_titulos))
                self.combobox_titulo.config(width = '45')
                self.combobox_titulo.grid(row=0, column=1, padx = '1', pady='1')
                        
                        
                #Informativo campo coordenada final Graus Decimais
                self.label5 = Label(self.frame_label_inf, text = 'Tipo de marcador:', font = 'arial')
                self.label5.grid(row=1, column=0, padx = '1', pady='1')
                
                #Informações para o marcador
                self.combobox_marcador = StringVar()
                self.combobox_marcador = ttk.Combobox(self.frame_label_inf, textvariable = self.combobox_marcador)
                self.combobox_marcador.config(values = (lista_itens_marcadores))
                self.combobox_marcador.config(width = '45')
                self.combobox_marcador.grid(row=1, column=1, padx = '1', pady='1')
                
                #Informativo para observação
                self.label8 = Label(self.frame_label_inf, text = 'Observação:         ', font = 'arial')
                self.label8.grid(row=2, column=0, padx = '1', pady='1')
                
                #Entrada de informações observação
                self.inf_obs = StringVar()
                self.entry6 = Entry(self.frame_label_inf)
                self.entry6.config(text = self.inf_obs)
                self.entry6.config(width = '48')
                self.entry6.grid(row=2, column=1, padx = '1', pady='1')
                        
                #Informações o combobox de cor
                self.label_utm_zonas = Label(self.frame_label_inf, text = 'Cor:       ', font = 'arial')
                self.label_utm_zonas.grid(row=1, column=3, padx = '1', pady='1')
                
                #Informações para as cores
                self.combobox_cor = StringVar()
                self.combobox_cor = ttk.Combobox(self.frame_label_inf, textvariable = self.combobox_cor)
                self.combobox_cor.config(values = (lista_cores))
                self.combobox_cor.config(width = '10')
                self.combobox_cor.grid(row=1, column=4, padx = '1', pady='1')

                #Informações o combobox de largura da linha
                self.label_utm_zonas = Label(self.frame_label_inf, text = 'Largura:', font = 'arial')
                self.label_utm_zonas.grid(row=2, column=3, padx = '1', pady='1')
                
                #Informações para a largura da linha
                self.combobox_largura= StringVar()
                self.combobox_largura = ttk.Combobox(self.frame_label_inf, textvariable = self.combobox_largura)
                self.combobox_largura.config(values = (lista_larguras_linha))
                self.combobox_largura.config(width = '10')
                self.combobox_largura.grid(row=2, column=4, padx = '1', pady='1')


                def inf_cadastro():
                    print('inf_cadastro')
                    inf_cadastro_f(self, criar_kml)
                    
                #Botão para informação sobre o campo informações para cadastro
                self.button_inf_cadastro = Button(self.frame_label_inf, text = '?', width = '1', height = '1', relief = GROOVE)
                self.button_inf_cadastro.config(bg = 'gold')
                self.button_inf_cadastro.config(command = inf_cadastro)
                self.button_inf_cadastro.grid(row=0, column = 4, padx = '1', sticky=N+S+E)


                def Enter_Infinf(criar_kml):
                    print('Informações para cadastro!')
                    self.button_inf_cadastro.config(relief = RAISED)
                self.button_inf_cadastro.bind('<Enter>', Enter_Infinf)
                
                def Leave_Infinf(criar_kml):
                    self.button_inf_cadastro.config(relief = GROOVE)
                self.button_inf_cadastro.bind('<Leave>', Leave_Infinf)
        

#UTM 
                if (utm == 1) or ((utm == 0) and (gd == 0)):
                    print('utm') 
                                   
                    #Frame informativo para as coordenadas em UTM
                    self.frame_label_utm = LabelFrame(self.inf_coord_lf, text = 'Coordenadas UTM', font = 'calibri', padx = '5', pady = '5')
                    self.frame_label_utm.grid(row=0, column=1, padx = '5', pady='5')
                    #Informativo campo latitude inicial
                    self.label3 = Label(self.frame_label_utm, text = 'Início:  ', font = 'arial')
                    self.label3.grid(row=0, column=0, padx = '1', pady='1')    
                    #Informativo campo longitude inicial
                    self.label4 = Label(self.frame_label_utm, text = 'Fim:   ', font = 'arial')
                    self.label4.grid(row=1, column=0, padx = '1', pady='1')
                    
                    
            ########################### Informações para as coordenadas em UTM ############################
                    
                    #Entrada de informações coordenada inicial
                    self.inf_ini_utm = StringVar()
                    self.inf_ini_utm = Entry(self.frame_label_utm)
                    self.inf_ini_utm.config(text = self.inf_ini_utm)
                    self.inf_ini_utm.config(width = '30')
                    self.inf_ini_utm.grid(row=0, column=1, padx = '10', pady='1')
                    
                    #Entrada de informações coordenada final
                    self.inf_fim_utm = StringVar()
                    self.inf_fim_utm = Entry(self.frame_label_utm)
                    self.inf_fim_utm.config(text = self.inf_fim_utm)
                    self.inf_fim_utm.config(width = '30')
                    self.inf_fim_utm.grid(row=1, column=1, padx = '1', pady='1')
                    
                    
                    
                    #Informações para a coordenada UTM - Designador de zona - Label                                  
                    self.label_utm_zonas = Label(self.frame_label_utm, text = 'Zona:', font = 'arial')
                    self.label_utm_zonas.grid(row=0, column=4, padx = '1', pady='1')       
                    
                    #Informações para a coordenada UTM - Designador de zona
                    self.combobox_utm_zonas = StringVar()
                    self.combobox = ttk.Combobox(self.frame_label_utm, textvariable = self.combobox_utm_zonas)
                    self.combobox.config(values = (lista_utm_zonas))
                    self.combobox.config(width = '3')
                    self.combobox.grid(row=0, column=5, padx = '1', pady='5')
                
                
                    #Botões para as coordenadas UTM
                    #Botão para salvar coordenadas coordenadas UTM
                    self.button_salvar_item = Button(self.frame_label_utm, text = 'Salvar', width = '8', height = '1', relief = GROOVE)
                    self.button_salvar_item.config(bg = 'azure')
                    self.button_salvar_item.config(command = salvar_item)
                    self.button_salvar_item.grid(row=1, column=4, padx = '10', pady='10')

                    def Enter_Salvarutm(criar_kml):
                        
                        #Apagando informação do usuário
                        self.label_inf_user.config (text = '')
                        
                        self.button_salvar_item.config(relief = RAISED)
                    self.button_salvar_item.bind('<Enter>', Enter_Salvarutm)
                    
                    def Leave_Salvarutm(criar_kml):
                        self.button_salvar_item.config(relief = GROOVE)                        
                    self.button_salvar_item.bind('<Leave>', Leave_Salvarutm)
                    
                    
                    #Botão para exportar kml coordenadas UTM
                    self.button_exportar_utm = Button(self.frame_label_utm, text = 'Exportar', width = '8', height = '1', relief = GROOVE)
                    self.button_exportar_utm.config(bg = 'azure')
                    self.button_exportar_utm.config(command = exportar_kml)
                    self.button_exportar_utm.grid(row=1, column=5)

                    def Enter_Exporteutm(criar_kml):
                        
                        #Apagando informação do usuário
                        self.label_inf_user.config (text = '')
                        
                        self.button_exportar_utm.config(relief = RAISED)
                    self.button_exportar_utm.bind('<Enter>', Enter_Exporteutm)
                    
                    def Leave_Exporteutm(criar_kml):
                        self.button_exportar_utm.config(relief = GROOVE)                      
                    self.button_exportar_utm.bind('<Leave>', Leave_Exporteutm)
                                        
                    
                    #Botão para informação coordenada UTM
                    self.button_inf_utm = Button(self.frame_label_utm, text = '?', width = '1', height = '1', relief = GROOVE)
                    self.button_inf_utm.config(bg = 'gold')
                    self.button_inf_utm.config(command = inf_utm)
                    self.button_inf_utm.grid(row=0, column=3)

                    def Enter_Infutm(criar_kml):
                        print('Informação UTM!')
                        self.button_inf_utm.config(relief = RAISED)
                    self.button_inf_utm.bind('<Enter>', Enter_Infutm)
                    
                    def Leave_Infutm(criar_kml):
                        self.button_inf_utm.config(relief = GROOVE)
                    self.button_inf_utm.bind('<Leave>', Leave_Infutm)
                    

#GD
                elif (gd == 1):
                    print('gd')
                    
                    #Caso exista a janela de cadastrar itens que ela seja destruída.
                    try:
                        self.frame_label_utm.destroy()
                    except:
                        pass
                    
                    #Frame informativo para as coordenadas em GD
                    self.frame_label_gd = LabelFrame(self.inf_coord_lf, text = 'Coordenadas Graus decimais', font = 'calibri', padx = '5', pady = '7')
                    self.frame_label_gd.grid(row=0, column=1, padx = '5', pady='5')
                    
                    
                    #Informativo campo coordenada inicial Graus Decimais
                    self.label3 = Label(self.frame_label_gd, text = 'Início:  ', font = 'arial')
                    self.label3.grid(row=0, column=0, padx = '1', pady='1')
                    
                    
                     #Entrada de informações  inicial
                    self.inf_ini_gd = StringVar()
                    self.entry3 = Entry(self.frame_label_gd)
                    self.entry3.config(text = self.inf_ini_gd)
                    self.entry3.config(width = '30')
                    self.entry3.grid(row=0, column=1, padx = '10', pady='1')
                     
                     #Entrada de informações  final
                    self.label3 = Label(self.frame_label_gd, text = 'Fim:   ', font = 'arial')
                    self.label3.grid(row=1, column=0, padx = '1', pady='1')
                    
                    self.inf_fim_gd = StringVar()
                    self.entry4 = Entry(self.frame_label_gd)
                    self.entry4.config(text = self.inf_fim_gd)
                    self.entry4.config(width = '30')
                    self.entry4.grid(row=1, column=1, padx = '1', pady='1')
                    
                    #Informações para a coordenada UTM - Designador de zona #### Apenas para que o valor da zona seja nulo quando a coordenada for em graus décimais
                    self.combobox_utm_zonas = StringVar()
                    self.combobox = ttk.Combobox(self.frame_label_gd, textvariable = self.combobox_utm_zonas)
                    self.combobox.config(values = (lista_utm_zonas))
#                     self.combobox.config(width = '3')
#                     self.combobox.grid(row=0, column=5, padx = '1', pady='5')
            
            
                    #Botões para as coordenadas Graus Decimais      
                    #Botão para salvar coordenadas Graus Decimais
                    self.button_salvar_gd = Button(self.frame_label_gd, text = 'Salvar', width = '8', height = '1', relief = GROOVE)
                    self.button_salvar_gd.config(bg = 'azure')
                    self.button_salvar_gd.config(command = salvar_item)
                    self.button_salvar_gd.grid(row=1, column=4, padx = '10', pady='10')
                    
                    def Enter_Salvargd(criar_kml):
                        #Apagando informação do usuário
                        self.label_inf_user.config (text = '')
                        
                        self.button_salvar_gd.config(relief = RAISED)
                    self.button_salvar_gd.bind('<Enter>', Enter_Salvargd)
                    
                    def Leave_Salvargd(criar_kml):
                        self.button_salvar_gd.config(relief = GROOVE)
                    self.button_salvar_gd.bind('<Leave>', Leave_Salvargd)
                    
                    
                    #Botão para exportar kml Graus Decimais
                    self.button_exportar_gd = Button(self.frame_label_gd, text = 'Exportar', width = '8', height = '1', relief = GROOVE)
                    self.button_exportar_gd.config(bg = 'azure')
                    self.button_exportar_gd.config(command = exportar_kml)
                    self.button_exportar_gd.grid(row=1, column=5)
                    
                    def Enter_Exportegd(criar_kml):
                        #Apagando informação do usuário
                        self.label_inf_user.config (text = '')
                        
                        self.button_exportar_gd.config(relief = RAISED)
                    self.button_exportar_gd.bind('<Enter>', Enter_Exportegd)
                    
                    def Leave_Exportegd(criar_kml):
                        self.button_exportar_gd.config(relief = GROOVE)                                                
                    self.button_exportar_gd.bind('<Leave>', Leave_Exportegd)
                    

                    #Botão para informação coordenada Graus Decimais
                    self.button_inf_gd = Button(self.frame_label_gd, text = '?', width = '1', height = '1', relief = GROOVE)
                    self.button_inf_gd.config(bg = 'gold')
                    self.button_inf_gd.config(command = inf_gd)
                    self.button_inf_gd.grid(row=0, column=3)

                    def Enter_Infgd(criar_kml):
                        print('Informação GD!')
                        self.button_inf_gd.config(relief = RAISED)
                    self.button_inf_gd.bind('<Enter>', Enter_Infgd)
                    
                    def Leave_Infgd(criar_kml):
                        self.button_inf_gd.config(relief = GROOVE)
                    self.button_inf_gd.bind('<Leave>', Leave_Infgd)
                    
                
                else:
                    pass
                    print('Erro, nem é graus decimais nem é UTM.')
                    
            except:
                pass
                print('Não foi possível definir qual é o tipo de coordenada.')
            
            tela1.update()

#################################### Antigo botão para confirmação de coordenadas ######################################
#         self.conf_inf = Button(self.inf_config, text = 'OK', width = '5', height = '1', relief = GROOVE)
#         self.conf_inf.config(bg = 'azure')
#         self.conf_inf.config(command = inf_iniciais)
#         self.conf_inf.grid(row=0, column=3)
        #self.conf_inf
        
#         def Enter_Inf(criar_kml):
#             print('Ok config!')
#             self.conf_inf.config(relief = RAISED)
#         self.conf_inf.bind('<Enter>', Enter_Inf)
#         
#         def Leave_Inf(criar_kml):
#             self.conf_inf.config(relief = GROOVE)
#         self.conf_inf.bind('<Leave>', Leave_Inf)
        
        
        inf_iniciais()
        
        
        
        
        
        
        #LabelFrame para o Treeview
        self.inf_cadastros = LabelFrame(criar_kml , text = 'Itens cadastrados', font = 'calibri')
        self.inf_cadastros.grid(row=4, column=0)    
            

        
        
        #Treeview para as obras
        self.treeview_cadastros = ttk.Treeview(self.inf_cadastros, height = 15, selectmode = 'browse')
        self.treeview_cadastros.pack(side = 'left', expand = False)
        
        
        #Scrollbar
        barra_vertical = ttk.Scrollbar(self.inf_cadastros,
                                       orient = 'vertical',
                                       command = self.treeview_cadastros.yview)
        barra_vertical.pack(side = 'left', fill = 'y')
        self.treeview_cadastros.configure(yscrollcommand = barra_vertical.set)
        
        


        self.treeview_cadastros['columns'] = ('1', '2', '3', '4', '5', '6', '7')
        self.treeview_cadastros['show'] = 'headings'
        
        #ID
        self.treeview_cadastros.column('1', width = 40, anchor = 'c')
        #Título
        self.treeview_cadastros.column('2', width = 250, anchor = 'w')
        #Observação
        self.treeview_cadastros.column('3', width = 300, anchor = 'w')
        #Ponto inicia
        self.treeview_cadastros.column('4', width = 150, anchor = 'c')
        #Ponto final
        self.treeview_cadastros.column('5', width = 150, anchor = 'c')
        #QTD
        self.treeview_cadastros.column('6', width = 55, anchor = 'c')
        #Un
        self.treeview_cadastros.column('7', width = 30, anchor = 'c')



        #Nomes para as colunas
        self.treeview_cadastros.heading('1', text = 'ID')
        self.treeview_cadastros.heading('2', text = 'Título')
        self.treeview_cadastros.heading('3', text = 'Observação')
        self.treeview_cadastros.heading('4', text = 'Ponto inicial')
        self.treeview_cadastros.heading('5', text = 'Ponto final')
        self.treeview_cadastros.heading('6', text = 'QTD')
        self.treeview_cadastros.heading('7', text = 'Un')   
        
        #Tentando alterar as caracteristicas do nome dado para a coluna
#         style = ttk.Style(criar_kml)
#         style.configure("Treeview.Headings", rowheight = 25, font = ('calibri', 14))
        
        #Informações das ações do usuário
        self.label_inf_user = Label(criar_kml, text = 'Informações para o usuário.', font = ('times', 15), background = 'light gray')
        self.label_inf_user.grid(row=5, column=0,  padx = '1', pady='5', sticky=N+S+W)
        
        
  # Funções para os botões relacionados às coordenadas em Graus Decimais       
        #Comando do botão salvar
        def salvar_gd():
            print('salvar_inf')
            salvar_gd_f(self, criar_kml)
        
        
        #Comando do botão exportar
        def exportar_kml():
            print('exportar_kml')
            exportar_gd_f(self, criar_kml)
            
# Função para o botão 'self.button_inf_gd'
        def inf_gd():
            print('inf_gd')
            inf_gd_f(self, criar_kml)
            
            

        #Comando do botão salvar
        def salvar_item():                      
            print('salvar_item.')
            salvar_item_f(self, criar_kml)
            
            
        #Comando do botão exportar UTM
        def exportar_kml():
            print('exportar_kml')
            exportar_utm_f(self, criar_kml)
        
# Função para o botão 'self.button_inf_utm'
        def inf_utm():
            print('inf_utm')
            inf_utm_f(self, criar_kml)
            

        def exportar_kml():
            print('exportar_kml')
            exportar_kml_f(self, criar_kml)
       


############  Botões para controle dos dados cadastrados ###################################
        ############## Funções ################

        def deletar_item():
            print('deletar_item')
            deletar_item_f(self, criar_kml)
        
        
        def limpar_cadastros():
            print('limpar_cadastros')
            verificar_limpar_cadastros(self, criar_kml)
            
        
        #Função do botão  self.button_exportar_csv
        def exporte_csv():
            print('exporte_csv')
            exporte_csv_f(self, criar_kml)
        
############  Botões para controle dos dados cadastrados ###############################        
        

#Botão para exportar arquivo.csv
        #Botão para exportar planilha
        self.button_exportar_csv = Button(self.inf_cadastros, text = 'CSV', width = '5', height = '1', relief = GROOVE)
        self.button_exportar_csv.config(bg = 'azure')
        self.button_exportar_csv.config(command = exporte_csv)
        self.button_exportar_csv.pack(anchor = NE, padx = '5', pady = '5')

        #self.button_exportar_csv
        def Enter_Csv(criar_kml):
            print('Csv!')
            self.button_exportar_csv.config(relief = RAISED  )
        self.button_exportar_csv.bind('<Enter>', Enter_Csv)
        
        def Leave_Csv(criar_kml):
            self.button_exportar_csv.config(relief = GROOVE )
        self.button_exportar_csv.bind('<Leave>', Leave_Csv)


        #Função do botão  self.button_exportar_doc
        def exporte_doc():
            print('exporte_doc')
#             try:
            print('exporte_doc_ok')
            exporte_doc_f(self, criar_kml)
#             except:
#                 print('exporte_doc_falha')
#                 pass

#Botão para exportar arquivo.doc
        #Botão para exportar documento
        self.button_exportar_doc = Button(self.inf_cadastros, text = 'DOC', width = '5', height = '1', relief = GROOVE)
        self.button_exportar_doc.config(bg = 'azure')
        self.button_exportar_doc.config(command = exporte_doc)
        self.button_exportar_doc.pack(anchor = NE, padx = '5', pady = '5')

        #self.button_exportar_doc
        def Enter_Doc(criar_kml):
            print('Doc!')
            self.button_exportar_doc.config(relief = RAISED  )
        self.button_exportar_doc.bind('<Enter>', Enter_Doc)
        
        def Leave_Doc(criar_kml):
            self.button_exportar_doc.config(relief = GROOVE )
        self.button_exportar_doc.bind('<Leave>', Leave_Doc)
        
        
#Botão para exportar arquivo.kml
        #Botão para exportar documento
        self.button_exportar_kml = Button(self.inf_cadastros, text = 'KML', width = '5', height = '1', relief = GROOVE)
        self.button_exportar_kml.config(bg = 'azure')
        self.button_exportar_kml.config(command = exportar_kml)
        self.button_exportar_kml.pack(anchor = NE, padx = '5', pady = '5')

        #self.button_exportar_kml
        def Enter_Kml(criar_kml):
            print('Kml!')
            self.button_exportar_kml.config(relief = RAISED  )
        self.button_exportar_kml.bind('<Enter>', Enter_Kml)
        
        def Leave_Kml(criar_kml):
            self.button_exportar_kml.config(relief = GROOVE )
        self.button_exportar_kml.bind('<Leave>', Leave_Kml)
        
#Botão para excluir item do self.treeview_cadastros
        #Botão para deletar item 
        self.button_deletar = Button(self.inf_cadastros, text = 'Deletar', width = '5', height = '1', relief = GROOVE)
        self.button_deletar.config(bg = 'salmon')
        self.button_deletar.config(command = deletar_item)
        self.button_deletar.pack(anchor = NE, padx = '5', pady = '50')

        #self.button_deletar
        def Enter_Delete(criar_kml):
            print('Delete!')
            self.button_deletar.config(relief = RAISED  )
            #Apagar informações para o usuário
            self.label_inf_user.config (text = '')
        self.button_deletar.bind('<Enter>', Enter_Delete)
        
        def Leave_Delete(criar_kml):
            self.button_deletar.config(relief = GROOVE )
        self.button_deletar.bind('<Leave>', Leave_Delete)
        
        
#Botão para exportar arquivo.doc
        #Botão para exportar documento
        self.button_limpar = Button(self.inf_cadastros, text = 'Limpar', width = '5', height = '1', relief = GROOVE)
        self.button_limpar.config(bg = 'salmon')        
        self.button_limpar.config(command = limpar_cadastros)
        self.button_limpar.pack(anchor = NE, padx = '5', pady = '30')        
        
        #self.button_limpar
        def Enter_Limpar(criar_kml):
            print('Limpar!')
            self.button_limpar.config(relief = RAISED  )
        self.button_limpar.bind('<Enter>', Enter_Limpar)
        
        def Leave_Limpar(criar_kml):
            self.button_limpar.config(relief = GROOVE )
        self.button_limpar.bind('<Leave>', Leave_Limpar)


        


#Função sem definição
#         def donothing():
#             pass
#             print('Nada.')

        #Função Sair do menu
        def Sair():
            print('Sair.')
            inf_sair = messagebox.askquestion(title = 'Verificar', message = 'Você deseja sair?', icon = 'warning')
            if inf_sair == 'yes':
                tela1.destroy()
            else:
                pass
                print('Manter programa aberto.')


        ################################ Fazer depois #################################
        #Função Selecionar UTM do menu
        def Selecionar_Utm():            
            print('Selecionar_Utm')
            
            #Seleciona o item pretendido
            self.inf_gd .deselect()
            self.inf_utm .select()
            
            #Atualiza a escolha
            inf_iniciais()

            
        #Função Selecionar graus decimais do menu
        def Selecionar_Gd():
            print('Selecionar_Gd')

            #Seleciona o item pretendido
            self.inf_utm .deselect()
            self.inf_gd .select()
            
            #Atualiza a escolha
            inf_iniciais()


        #Função Novo do menu
        def Novo_Arquivo_menu():
            print('Novo_Arquivo_menu')
            verificar_limpar_cadastros(self, criar_kml)



        #Função Importar .csv do menu
        def importar_itens_menu():
            print('importar_itens_menu')            
            global data_geral
            importar_itens_f(self, criar_kml)
        
        def exporte_csv_menu():
            print('exporte_csv_menu')
            exporte_csv_f(self, criar_kml)

        def exporte_doc_menu():
            print('exporte_doc_menu')
            exporte_doc_f(self, criar_kml)
            
        def menu_arq_github():
            print('menu_arq_github')
            menu_arq_github_f(self, criar_kml)

        def menu_arq_sobre():
            print('menu_arq_sobre')
            menu_arq_sobre_f(self, criar_kml)
            
            
        self.menubar = Menu(criar_kml)
        self.filemenu = Menu(self.menubar, tearoff=0)
        self.filemenu.config(background='light sky blue', foreground='black', activebackground='blue', activeforeground='white')
        self.filemenu.add_command(label="Novo", command = Novo_Arquivo_menu)
#         self.filemenu.add_command(label="Abrir", command=donothing)
#         self.filemenu.add_command(label="Salvar", command=donothing)
        self.filemenu.add_command(label="Exportar .csv", command =exporte_csv_menu)
        self.filemenu.add_command(label="Exportar .doc", command = exporte_doc_menu)
        self.filemenu.add_command(label="Importar .csv", command = importar_itens_menu)
        self.menubar.add_cascade(label="Arquivo", menu = self.filemenu)
        self.filemenu.add_separator()
        self.filemenu.add_command(label = 'Sair', command = Sair)


        self.editmenu = Menu(self.menubar, tearoff=0)
        self.editmenu.config(background='light sky blue', foreground='black', activebackground='blue', activeforeground='white')
#         editmenu.add_command(label="Limpar itens", command=donothing)
#         editmenu.add_command(label="Deletar item", command=donothing)
        ################################ Fazer depois #################################
        self.editmenu.add_command(label="Selecionar UTM", command = Selecionar_Utm)
        self.editmenu.add_command(label="Selecionar graus decimais", command = Selecionar_Gd)
        self.menubar.add_cascade(label="Editar", menu = self.editmenu)

        self.helpmenu = Menu(self.menubar, tearoff=0)
        self.helpmenu.config(background='light sky blue', foreground='black', activebackground='blue', activeforeground='white')
        self.helpmenu.add_command(label="Arquivo do Github", command = menu_arq_github)
        self.helpmenu.add_command(label="Sobre...", command = menu_arq_sobre)
        self.menubar.add_cascade(label="Ajuda", menu = self.helpmenu)
        
        self.menubar.config( background = 'gray', fg='white')
        criar_kml.config(menu = self.menubar)
        
        
        #Funções das teclas de atalho
        
        #Atalho para exportar arquivo.kml
        keyboard.add_hotkey("ctrl+e", lambda: exportar_kml_f(self, criar_kml))
        
        #Atalho para deletar item
        keyboard.add_hotkey("ctrl+d", lambda: deletar_item_f(self, criar_kml))
        
        #Atalho para salvar item
        keyboard.add_hotkey("ctrl+s", lambda: salvar_item_f(self, criar_kml))


tela1 = Tk()
criar_kml = CriarKml(tela1)





#Função Sair tela1
def Sair():
    print('Sair.')
    inf_sair = messagebox.askquestion(title = 'Verificar', message = 'Você deseja sair?', icon = 'warning')
    if inf_sair == 'yes':
        tela1.destroy()
    else:
        pass
        print('Manter programa aberto.')                
                
#Protocolo para fechar o programa
tela1.protocol("WM_DELETE_WINDOW", Sair)


#Configurando a tela
tela1.title('KmlPlote')
# tela1.geometry('1200x750+100+40')

#Maximização da tela ao iniciar
tela1.wm_state('zoomed')

#Ícone do programa na tela do Tkinter
tela1.iconbitmap('kp.ico')

tela1.mainloop()







