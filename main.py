from docx2pdf import convert
import PySimpleGUI as sg
import os

def converter_arquivo(caminho):
    try:
        convert(r'{}'.format(caminho), r'{}'.format(caminho).replace('.docx', '.pdf'))
        return True
    except:

        return False

class TelaConversor:
    def __init__(self):

        sg.theme('Dark')

        layout = [
                [sg.Text('Buscar arquivo:')],
                [sg.Text('Upload PDF:', size=(5, 0)), sg.Input(size=(40, 0), key='upload'), sg.FileBrowse('Buscar')],
                [sg.Text('Status:', size=(5, 0)), sg.Text(size=(40, 0), key='status')],
                [sg.Button('Converter arquivo em PDF', size=(49, 0))]
            ]

        self.janela = sg.Window('Conversor de Word para PDF').layout(layout)

    def Iniciar(self):
        while True:
            self.button, self.values = self.janela.Read()
            if self.button == sg.WINDOW_CLOSED:
                break

            if self.button == 'Converter arquivo em PDF':
                filepath = self.values['upload']
                
                diretorio = os.path.dirname(filepath)
                nome_arquivo = os.path.basename(filepath)
                caminho = f'{diretorio}/{nome_arquivo}'                
                status = converter_arquivo(caminho)

                if status:
                    self.janela['status'].update('Arquivo convertido com sucesso!')
                else:
                    self.janela['status'].update('Erro ao converter o arquivo.')

tela = TelaConversor()
tela.Iniciar()




