import PySimpleGUI as sg
import os
from openpyxl import Workbook
# Define the window's contents
import pandas as pd
import time

class SeparadorPlanilha():
    def __init__(self,):

        self.working_directory = os.getcwd()
        self.layout = [  
        [sg.Text("Coloque o arquivo Excel")],
        [sg.InputText(key='-FILE_PATH-'),
        sg.FileBrowse(initial_folder=self.working_directory, file_types=[("ALL CSV FILES", "*.xlsx *.csv")])],
        [sg.Button('Ok')],
        ]
        
    def Windoww(self):
        self.window = sg.Window('Swan', self.layout)

    def primeiraTela(self):
        while True:   
            event, values = self.window.read()
            self.csv_address = values["-FILE_PATH-"]
            if event in (sg.WIN_CLOSED, 'Exit'):
                break
            elif event == "Ok":
                csv_address = values["-FILE_PATH-"]
                if not csv_address == '':
                    self.window.close()
                    window.segundaTela()
    def segundaTela(self):
        if self.csv_address[-4:] == 'xlsx':
            self.csv = pd.read_excel(self.csv_address)
            self.colunascsv = list(self.csv.columns)
        elif self.csv_address[-3:] == 'csv':
            self.csv = pd.read_csv(self.csv_address, error_bad_lines=False, encoding='ISO-8859-1', sep=';')
            self.colunascsv = list(self.csv.columns)
        self.layout2 = [  
        [sg.Text("Selecione a coluna chave")], 
        [sg.Listbox(values=self.colunascsv, size=(55,5),select_mode='single', key='ColunaEscolhida')],
        [sg.Text("Salvar em:",key='-Salvar_em-')], 
 
        [sg.Input(key='-FOLDER-'), sg.FolderBrowse()],
        [sg.Text("", key='_carregamentoColunaChave_')],
        [sg.Button('Ok'),sg.Button('Sair')]
                                ]
        self.window = sg.Window('Swan', self.layout2)
        while True:
            event, values = self.window.read()
            self.localPasta = values['-FOLDER-']
            if event in (sg.WIN_CLOSED, 'Exit'):
                break
            elif event == 'Ok':
                if not values == {'ColunaEscolhida': []}:
                    ColunaChave =  (values['ColunaEscolhida'])
                    window.SeparandoAsPlanilhas(list(ColunaChave))
            elif event == 'Sair':
                break
    def SeparandoAsPlanilhas(self,ColunaChave):
        ind = {1: 'A1', 2:'B1', 3:'C1', 4:'D1', 5:'E1', 6:'F1', 7:'G1', 8: 'H1', 9:'I1', 10: 'J1',11: 'K1', 12:'L1', 13: 'M1', 14: 'N1', 15: 'O1', 16: 'P1', 17: 'Q1', 18: 'R1', 19:'S1', 20:'T1',21:'U1',22:'V1', 23:'W1', 24: 'X1', 25:'Y1',26:'Z1',27:'AA1',28:'AB1',29:'AC1',30:'AD1',31:'AE1',32:'AF1',33:'AG1',34:'AH1',35:'AI1',36:'AJ1',37:'AK1',38:'AL1',39:'AM1',40:'AN1',41:'AO1',42:'AP1',43:'AQ1',44:'AR1'  }
        nomesUnicos = self.csv[f'{ColunaChave[0]}'].unique()
        concluido = 1
        print(list(nomesUnicos))
        for dadoDaColuna in list(nomesUnicos):
            novaPlanilha = Workbook()
            sheet = novaPlanilha.active
            self.window['_carregamentoColunaChave_'].update(f'{concluido} de {len(nomesUnicos)} ')
            adicionandoColunasNewCSV = 1
            for colunas in self.colunascsv:
                sheet[ind[adicionandoColunasNewCSV]] = colunas
                adicionandoColunasNewCSV += 1
            for i in range (len(self.csv)):
                if dadoDaColuna == self.csv.loc[i][f'{ColunaChave[0]}']:
                    sheet.append(list(self.csv.loc[i]))
            novaPlanilha.save(filename=f'{self.localPasta}' '/' f'{dadoDaColuna}' + '.xlsx')
            concluido += 1
            self.window.refresh()
        self.window['_carregamentoColunaChave_'].update('Concluido')   
        self.window.refresh()

window = SeparadorPlanilha()
a = window.Windoww()
window.primeiraTela()