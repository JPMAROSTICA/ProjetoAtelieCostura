import customtkinter as ctk
from CTkListbox import *
from tkinter import END
from CTkMessagebox import CTkMessagebox
import datetime
import pandas as pd

class Ajustes:
    def __init__(self):
        self.__ajustes = {
            #AJUSTE,TEMPO,VALOR
            1: ['AJUSTE DE ALÇA',15,20],
            2: ['AJUSTE DE BOCA',20,20],
            3: ['AJUSTE DE CÓS',20,20],
            4: ['AJUSTE DE OMBRO',15,20],
            5: ['AJUSTE DE OMBRO E BARRA',30,20],
            6: ['AJUSTE DE  REBORDAGEM',40,60],
            7: ['AJUSTE LATERAL',15,20],
            8: ['AJUSTE DE MANGA',20,20],
            9: ['AJUSTE NAS COSTAS',15,20],
            10: ['AUMENTAR DECOTE',15,20],
            11: ['BARRA',15,20],
            12: ['BARRA DE CORTINA',15,20],
            13: ['BARRA DE LENÇO',20,20],
            14: ['BARRA E AFUNILAR PERNA',20,40],
            15: ['CERZIDO',20,20],
            16: ['AUMENTAR CÓS',40,60],
            17: ['NESGA NO CAVALO',20,20],
            18: ['PENCES',15,20],
            19: ['RECOSTURA',10,20],
            20: ['SUBIR PUNHO',20,20],
            21: ['TROCAR ELÁSTICO',20,20],
            22: ['TROCAR ZÍPER',15,20],
            23: ['TROCAR/COLOCAR BOTÕES',10,20],
            24: ['VIRAR COLARINHO',20,20],
            25: ['COLOCAR BOJO',15,20],
            26: ['COLOCAR COLCHETE',10,20],
            27: ['AJUSTE CÓS E BARRA',30,30],
            28: ['AJUSTE LATERAL E BARRA',20,25],
            29: ['AJUSTE LATERAL, BARRA E PUNHO',30,40],
            30: ['AJUSTE CÓS, LATERAL E BARRA',30,40],
            31: ['AJUSTE LATERAL E PENCE',20,20],
            32: ['AJUSTE BOCA E BARRA',20,20],
            33: ['AJUSTE LATERAL E PUNHO',30,40],
            34: ['AJUSTE ALÇA E LATERAL',20,20],
            35: ['AJUSTE DE TAMANHO',40,60],
            36: ['AJUSTES DIVERSOS',20,40],
            37: ['AJUSTE CÓS E LATERAL',30,40],
            38: ['AJUSTE ALÇA E BARRA',30,40],
            39: ['AUMENTAR LATERAIS',40,60],
            40: ['BARRA NA MANGA E COMPRIMENTO',20,20]
        }
        self.__cores = {
            1: 'AMARELO',
            2: 'AZUL',
            3: 'AZUL CLARO',
            4: 'BEJE',
            5: 'BRANCO',
            6: 'CAQUI',
            7: 'CINZA CLARO',
            8: 'JEANS',
            9: 'LARANJA',
            10: 'LILÁS',
            11: 'LIMA',
            12: 'MARINHO',
            13: 'MARROM',
            14: 'PINK',
            15: 'PRETO',
            16: 'ROSA',
            17: 'ROSA CLARO',
            18: 'ROXO',
            19: 'TERRACOTA',
            20: 'VERDE CLARO',
            21: 'VERDE ESCURO',
            22: 'VERMELHO',
            23: 'AZUL ROYAL',
            24: 'CARAMELO',
            25: 'VERDE BANDEIRA',
            26: 'VINHO',
            27: 'PRATA',
            28: 'DOURADO',
            29: 'MARROM CLARO',
            30: 'CINZA ESCURO'
        }

    def get_ajuste(self,key):
        if int (key) >= 1 and int (key) <= len(self.__ajustes): return self.__ajustes[key][0]
        else: return 'NotFound'

    def get_tempo(self,key):
        if int (key) >= 1 and int (key) <= len(self.__ajustes): return self.__ajustes[key][1]
        else: return 'NotFound'

    def get_valor(self,key):
        if int (key) >= 1 and int (key) <= len(self.__ajustes): return self.__ajustes[key][2]
        else: return 'NotFound'

    def get_cor(self,key):
        if int (key) >= 1 and int (key) <= len(self.__cores): return self.__cores[key]
        else: return 'NotFound'

    def get_listaAjustes(self):
        return self.__ajustes

    def get_listaCores(self):
        return self.__cores

class Validadores:
    def verifyIsANumber(self,data):
        try:
            int(data)
            return True
        except ValueError: return False

    def verifyBlankFields(self,data):
        if data == "": return True
        return False

    def generateDayTime(self):
        if datetime.date.today().day >=10: day = str(datetime.date.today().day)
        else: day = "0" + str(datetime.date.today().day)

        if datetime.date.today().month >= 10: month = str(datetime.date.today().day)
        else: month = "0" + str(datetime.date.today().month)

        date = day + "/" + month + "/" + str(datetime.date.today().year)
        return date

    def definirNomeDoAjuste(self,isUrgenciaChecked,name,boleta):
        name = str(boleta) + " " + name + "_URGÊNCIA" if isUrgenciaChecked else str(boleta) + " " + name
        return name

class Operadores:
    def formatarTempoPadraoHorasMinutos(self,listTempo):
        tempoEmMinutos = sum(listTempo)
        hour = ""
        min = ""

        if tempoEmMinutos >= 60:
            resto = tempoEmMinutos%60
            horas = int((tempoEmMinutos - resto)/60)
            minutos = round(resto)

            if horas >= 1 and horas <= 9: hour = "0"+str(horas)
            elif horas >=10: hour = str(horas)

            if minutos >= 0 and minutos <= 9: min = "0"+str(minutos)
            elif minutos >= 10: min = str(minutos)

        else:
            hour = "00"
            if tempoEmMinutos >= 0 and tempoEmMinutos <= 9: min = "0"+str(tempoEmMinutos)
            if tempoEmMinutos >= 10: min = str(tempoEmMinutos)

        return hour + ":" + min

class Excel:
    def gerarPlanilhaDeAjustes(self,dadosAjustes):
        dadosPlanilha = {
            "Cliente": dadosAjustes["Loja"],
            "Serviço": dadosAjustes["Ajuste"],
            "Valor Unitário": dadosAjustes["Valor"],
            "Data": dadosAjustes["Data"]
        }


        dadosEmDataFrame = pd.DataFrame(data=dadosPlanilha)
        dadosEmDataFrame = dadosEmDataFrame.sort_values(by="Cliente")

        dadosEmDataFrame.to_excel("Ajustes.xlsx",index=False)

    def gerarPlanilhaDeHoras(self,dadosAjustes):
        dadosPlanilha = {
            "Cor": dadosAjustes["Cor"],
            "Serviço": dadosAjustes["Ajuste"],
            "Tempo": dadosAjustes["Tempo"]
        }

        dadosEmDataFrame = pd.DataFrame(data=dadosPlanilha)
        dadosEmDataFrame = dadosEmDataFrame.sort_values(by="Cor")

        dadosEmDataFrame.to_excel("Horas.xlsx",index=False)

class GerenteDeForm:
    def __init__(self):
        self.objAjustes = Ajustes()
        self.validadores = Validadores()
        self.operadores = Operadores()
        self.excel = Excel()

    def gerarLabelDeErro(self,typeError):
        if typeError == 'branco':
            texto = "Campos de Ajuste, Cor e Loja são obrigatórios!"
            CTkMessagebox(title="Error", message=texto, icon="cancel")

        if typeError == 'number':
            texto = "Digite somente números nos campos de Ajuste e Cor!"
            CTkMessagebox(title="Error", message=texto, icon="cancel")

        if typeError == 'NotFound':
            texto = "Número do ajuste e/ou número da cor fora dos dados listados!"
            CTkMessagebox(title="Error", message=texto, icon="cancel")

        if typeError == 'gerarPlanilha':
            texto = "Não foi possível gerar a planilha. Verifique se arquivos gerados com o mesmo nome não se encontram abertos. Senão, tente novamente mais tarde."
            CTkMessagebox(title="Error", message=texto, icon="cancel")

    def gerarLabelDeSucesso(self,success):
        if success == "gerarPlanilha":
            texto = "Planilhas geradas com sucesso!"
            CTkMessagebox(title="Sucesso", message=texto, icon="check")

    def gerarMensagemDeConfirmacao(self,typeMsg):
        if typeMsg == "gerarPlanilha":
            texto = "Se houver alguma planilha com o nome Ajustes.xlsx ou Horas.xlsx, ela terá os dados sobrescritos. Deseja prosseguir?"

            msg = CTkMessagebox(title="Confirm", message=texto,
                                icon="question", option_1="Yes", option_2="No")
            response = msg.get()

        if typeMsg == "clean":
            texto = "Esta operação irá excluir todos os serviços lançados. Deseja prosseguir?"

            msg = CTkMessagebox(title="Confirm", message=texto,
                                icon="question", option_1="Yes", option_2="No")
            response = msg.get()

        return response

    def gerarMensagensDeAlertaPlanilha(self,typeMsg):
        if typeMsg == "gerarPlanilha":
            texto = "Planilhas não foram geradas!"
            CTkMessagebox(title="Atenção!", message=texto, icon="warning")

class FormularioLojas(GerenteDeForm):
    def __init__(self):
        #cliente serviço valor_unitario quantidade valor_total data
        self.dadosDosAjustes = {
            "Loja": [],
            "Ajuste": [],
            "Tempo": [],
            "Cor": [],
            "Valor": [],
            "Data": []
        }
        self.entryBoleta = ''
        self.entryAjuste = ''
        self.entryCor = ''
        self.entryLoja = ''
        self.entryData = ''
        self.checkUrgencia = False
        self.janela = ctk.CTk()
        self.listBox = CTkListbox(self.janela,
                             width=310,
                             height=210,
                             hover_color='black',
                             )

        super().__init__()

    def apresentarTextoDosAjustes(self):
        titulo = ctk.CTkLabel(self.janela, text="AJUSTES",
                              text_color='white',
                              fg_color='#01a6f8',
                              font=('Bold',15),
                              width=1130,
                              corner_radius = 5,
                              anchor='w')
        titulo.place(x=10,y=10)

        ajustes = self.objAjustes.get_listaAjustes()

        linha,coluna = 40,10
        for key in ajustes:
            numeroAjuste = ctk.CTkLabel(self.janela, text=str(key), anchor='w', text_color='light blue')
            numeroAjuste.place(x=coluna,y=linha)

            ajuste = ctk.CTkLabel(self.janela, text=str(ajustes[key][0]), anchor='w')
            ajuste.place(x=coluna+20,y=linha)

            linha += 20
            if linha >= 240:
                linha = 40
                coluna += 300

    def apresentarTextoDasCores(self):

        titulo = ctk.CTkLabel(self.janela, text="CORES",
                              text_color='white',
                              fg_color='#01a6f8',
                              font=('Bold',15),
                              width=550,
                              corner_radius = 5,
                              anchor='w')
        titulo.place(x=10,y=270)

        cores = self.objAjustes.get_listaCores()

        linha,coluna = 300,10
        for key in cores:
            numeroCor = ctk.CTkLabel(self.janela, text=str(key), anchor='w', text_color='light blue')
            numeroCor.place(x=coluna,y=linha)

            cor = ctk.CTkLabel(self.janela, text=str(cores[key]), anchor='w')
            cor.place(x=coluna+20,y=linha)


            linha += 20
            if linha == 600:
                linha = 300
                coluna += 300

    def formatarFormulario(self):
        self.entryBoleta = ctk.CTkEntry(self.janela,placeholder_text="Boleta",width=200)
        self.entryBoleta.place(x=590,y=320)

        self.entryAjuste = ctk.CTkEntry(self.janela,placeholder_text="Número Do Ajuste",width=200)
        self.entryAjuste.place(x=590,y=360)

        self.entryCor = ctk.CTkEntry(self.janela,placeholder_text="Número Da Cor",width=200)
        self.entryCor.place(x=590,y=400)

        self.checkUrgencia = False
        urgencia = ctk.StringVar(value="off")
        checkUrgencia = ctk.CTkCheckBox(self.janela,text='URGÊNCIA',hover_color='blue',variable=urgencia,onvalue='on',offvalue='off',command=self.changeCheckBoxUrgencia)
        checkUrgencia.place(x=590,y=440)

    def gerarListBox(self):
        self.listBox = CTkListbox(self.janela,
                             width=310,
                             height=210,
                             hover_color='black',
                             )

        if len(self.dadosDosAjustes["Ajuste"]) > 0:
            for ajustes in self.dadosDosAjustes["Ajuste"]:
                self.listBox.insert(END, ajustes)

        self.listBox.place(x=800,y=320)

    def gerarFormsDeEntradaDeServicos(self):
        titulo = ctk.CTkLabel(self.janela, text="DADOS",
                              text_color='white',
                              fg_color='#01a6f8',
                              font=('Bold',15),
                              width=550,
                              corner_radius = 5,
                              anchor='w')
        titulo.place(x=590,y=270)

        buttonCleanUp = ctk.CTkButton(self.janela,text="Limpar", width=30, height=28, command=self.limparDados)
        buttonCleanUp.place(x=1094,y=270)

        self.entryBoleta = ctk.CTkEntry(self.janela,placeholder_text="Boleta",width=200)
        self.entryBoleta.place(x=590,y=320)

        self.entryAjuste = ctk.CTkEntry(self.janela,placeholder_text="Número Do Ajuste",width=200)
        self.entryAjuste.place(x=590,y=360)

        self.entryCor = ctk.CTkEntry(self.janela,placeholder_text="Número Da Cor",width=200)
        self.entryCor.place(x=590,y=400)

        self.checkUrgencia = False
        urgencia = ctk.StringVar(value="off")
        checkUrgencia = ctk.CTkCheckBox(self.janela,text='URGÊNCIA',hover_color='blue',variable=urgencia,onvalue='on',offvalue='off',command=self.changeCheckBoxUrgencia)
        checkUrgencia.place(x=590,y=440)

        buttonNext = ctk.CTkButton(self.janela,text="Next", width=80, command=self.validateAndTransformData)
        buttonNext.place(x=710,y=440)

        # criamos um objeto listBox Para apresentar os ajustes lançados
        self.gerarListBox()
        self.apresentarONumeroDeAjustesEHoras()

    def gerarFormsDeEntradaLojaEDataButtonGerador(self):
        self.entryLoja = ctk.CTkEntry(self.janela,placeholder_text="LOJA",width=200)
        self.entryLoja.place(x=590,y=480)

        date = self.validadores.generateDayTime()

        self.entryData = ctk.CTkEntry(self.janela,placeholder_text=date,width=200)
        self.entryData.place(x=590,y=520)

        buttonNext = ctk.CTkButton(self.janela,text="Gerar Planilhas", width=200, command=self.gerarPlanilhas)
        buttonNext.place(x=590,y=560)

    def changeCheckBoxUrgencia(self):
        self.checkUrgencia = True if not self.checkUrgencia else False

    def inserirDadosValidadosEmListas(self,ajuste,cor,boleta,loja,date):
        nomeDoAjuste = self.validadores.definirNomeDoAjuste(self.checkUrgencia,self.objAjustes.get_ajuste(ajuste),boleta)

        self.dadosDosAjustes["Loja"].append(loja)
        self.dadosDosAjustes["Ajuste"].append(nomeDoAjuste)
        self.dadosDosAjustes["Tempo"].append(self.objAjustes.get_tempo(ajuste))
        self.dadosDosAjustes["Cor"].append(self.objAjustes.get_cor(cor))
        self.dadosDosAjustes["Valor"].append(self.objAjustes.get_valor(ajuste))
        self.dadosDosAjustes["Data"].append(date)

        return nomeDoAjuste

    def apresentarONumeroDeAjustesEHoras(self):
        tempo = self.operadores.formatarTempoPadraoHorasMinutos(self.dadosDosAjustes["Tempo"])

        texto = "TOTAL DE AJUSTES: " + str(len(self.dadosDosAjustes["Ajuste"])) + "                  " + "TEMPO: " + str(tempo)
        titulo = ctk.CTkLabel(self.janela, text=texto,
                              text_color='white',
                              fg_color='#173f4f',
                              font=('Bold',15),
                              width=342,
                              corner_radius = 5,
                              anchor='w')
        titulo.place(x=800,y=560)

    def validateAndTransformData(self):
        #dados obtidos do formulário
        boleta = str(self.entryBoleta.get())
        ajuste = str(self.entryAjuste.get())
        cor = str(self.entryCor.get())
        loja = str(self.entryLoja.get())
        date = self.validadores.generateDayTime() if str(self.entryData.get()) == "" else self.entryData.get()


        #VALIDAMOS OS DADOS NECESSÁRIOS
        #o método verifyBlankFields retornará True quando o campo estiver em branco
        if self.validadores.verifyBlankFields(ajuste) or self.validadores.verifyBlankFields(cor) or self.validadores.verifyBlankFields(loja):
            self.gerarLabelDeErro('branco')
            return

        if not self.validadores.verifyIsANumber(ajuste) or not self.validadores.verifyIsANumber(cor):
            self.gerarLabelDeErro('number')
            return

        ajuste = int(ajuste)
        cor = int(cor)

        if self.objAjustes.get_ajuste(ajuste) == 'NotFound' or self.objAjustes.get_cor(cor) == 'NotFound':
            self.gerarLabelDeErro('NotFound')
            return

        ajustes = self.inserirDadosValidadosEmListas(ajuste,cor,boleta,loja,date)
        self.listBox.insert(END, ajustes)
        self.apresentarONumeroDeAjustesEHoras()
        self.formatarFormulario()

    def gerarFormulario(self):
        #CRIAMOS A JANELA
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("dark-blue")
        self.janela.title("AJUSTES DE LOJAS")
        self.janela.geometry("1150x610")

        #INSERIMOS OS ELEMENTOS NA JANELA CRIADA
        self.apresentarTextoDosAjustes()
        self.apresentarTextoDasCores()
        self.gerarFormsDeEntradaDeServicos()
        self.gerarFormsDeEntradaLojaEDataButtonGerador()

        self.janela.mainloop()

    def gerarPlanilhas(self):
        resposta = self.gerarMensagemDeConfirmacao("gerarPlanilha")
        if (resposta == 'Yes'):
            try:
                self.excel.gerarPlanilhaDeHoras(self.dadosDosAjustes)
                self.excel.gerarPlanilhaDeAjustes(self.dadosDosAjustes)

                self.gerarLabelDeSucesso("gerarPlanilha")
            except PermissionError:
                self.gerarLabelDeErro("gerarPlanilha")
        else:
            self.gerarMensagensDeAlertaPlanilha("gerarPlanilha")

    def limparDados(self):
        resposta = self.gerarMensagemDeConfirmacao("clean")
        if resposta == 'Yes':
            for key in self.dadosDosAjustes: self.dadosDosAjustes[key].clear()
            self.entryBoleta = ''
            self.entryAjuste = ''
            self.entryCor = ''
            self.entryLoja = ''
            self.entryData = ''
            self.checkUrgencia = False

            self.gerarFormsDeEntradaDeServicos()
            self.gerarFormsDeEntradaLojaEDataButtonGerador()









