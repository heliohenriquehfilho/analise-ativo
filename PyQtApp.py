import pandas as pd 
import yfinance as yf
from datetime import datetime, timedelta
import matplotlib.pyplot as plt
from matplotlib.figure import Figure
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from PyQt5.QtWidgets import QApplication, QMessageBox, QMainWindow, QLineEdit, QLabel, QComboBox, QPushButton, QVBoxLayout, QHBoxLayout, QWidget, QTreeWidget, QTreeWidgetItem, QHeaderView
from PyQt5.QtCore import Qt
import matplotlib.pyplot as plt
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from PyQt5.QtCore import Slot
from PyQt5.QtGui import QColor
from pandas.core.arrays.datetimelike import frequencies
from functools import partial
import webbrowser
import plotly.graph_objects as go
from PyQt5.QtWidgets import QComboBox, QPushButton, QVBoxLayout, QHBoxLayout, QWidget, QTreeWidget, QTreeWidgetItem, QHeaderView, QProgressDialog
from PyQt5.QtGui import QColor
from fpdf import FPDF
import pickle
from PyQt5.QtWidgets import QDialog
import mplfinance as mpf
import webview
import os
import firebase_admin
from firebase_admin import credentials
from firebase_admin import db

# Path para o arquivo JSON de credenciais
cred = credentials.Certificate(r'Dados/analise-ativos-firebase-adminsdk-j64bc-81c5088e13.json')
firebase_admin.initialize_app(cred, {
    'databaseURL': 'https://analise-ativos-default-rtdb.firebaseio.com/'
})

# Inicialize o aplicativo Firebase
firebase_admin.initialize_app(cred, name='analise-ativos')

versao = "V0.7"

today = datetime.today()
yesterday = today - timedelta(days=1)

assets = {"Ativo": [], 
          "Cotas": [],
          "Dividendos": [],
          "Volatilidade": [],
          "Preco": [],
          'Quant Min': [],
          'Progresso': [],
          'Data Compra': [],
          "Historico": [],
          "Resultado": [],
          'Total': [],
          'Frequencia': [],
          'Div Total': [],
          'Rentabilidade(%)': [],
          'Data Compra': [],
          "Preco Compra" : [],
          "ROI": []}
days = []
df_ativos = pd.DataFrame(assets)

name = ""
dados = {}
global Users
Users = []

arquivoPickle = r'Dados\dados.pickle'

pickles = {}

try:
    with open(arquivoPickle, "rb") as arquivo:
        pickles = pickle.load(arquivo)
except FileNotFoundError:
    print("Arquivo 'dados.pickle' não encontrado. Criando novo arquivo")

with open(arquivoPickle, "wb") as arquivo:
    pickle.dump(pickles, arquivo)

class PDF(FPDF):
    def inicio_pdf(dataframe, nome):
        pdf = PDF(orientation='L')
        pdf.set_auto_page_break(auto=True, margin=15)
        
        # Adicionar uma página ao PDF
        pdf.add_page()
        
        # Definir a fonte e o tamanho do texto
        pdf.set_font("Arial", size=8)
        
        # Título do PDF
        pdf.cell(0, 10, txt="Relatório de Ativos", ln=True, align="C")
        pdf.cell(0, 5, txt=f"{name}", ln=True, align="C")
        pdf.cell(0, 10, txt=F"Análise de Ativo {versao}", ln=True, align="C")
        pdf.ln(10)

        headers = df_ativos.columns.tolist()
        column_widths = []

        for header in headers:
            header_width = pdf.get_string_width(header) + 2  # Adicionar um espaço extra
            column_widths.append(header_width)

        # Calcular a largura máxima entre o cabeçalho e o valor da coluna
        for _, row in df_ativos.iterrows():
            for i in range(len(row)):
                value = str(row[i])
                value_width = pdf.get_string_width(value) + 3  # Adicionar um espaço extra
                cell_width = max(column_widths[i], value_width)
                column_widths[i] = cell_width  # Atualizar a largura da coluna se necessário

        # Adicionar o cabeçalho com as larguras atualizadas
        for i, header in enumerate(headers):
            pdf.cell(column_widths[i], 10, txt=header, border=1, align="C")

        pdf.ln()
        
        # Adicionar os valores das colunas
        for _, row in df_ativos.iterrows():
            for i in range(len(row)):
                pdf.cell(column_widths[i], 10, txt=str(row[i]), border=1, align="C")
            
            pdf.ln()

        # Gerar e adicionar os gráficos ao PDF

        ativos = df_ativos['Ativo']

        for i in ativos:
            symbol = i
            title = f"Evolução de Preço - {symbol}"
            chart_path = f"chart_{i}.png"

            # Obter os dados históricos da ação usando o yfinance
            df = yf.download(symbol, start=today - timedelta(days=180), end=today)

            # Limpar o gráfico anterior
            plt.clf()

            Open = []
            High = []
            Low = []
            Close = []

            for i in range(len(df)):
                Open.append(df['Open'].iloc[i])
                High.append(df['High'].iloc[i])
                Low.append(df['Low'].iloc[i])
                Close.append(df['Close'].iloc[i])

            data = pd.DataFrame({
                "Open": Open,
                "High": High,
                "Low" : Low,
                "Close": Close
            }, index=df.index)

            mpf.plot(data, type='candle', style='yahoo', ylabel="Data")

            # Salvar o gráfico em um arquivo
            plt.savefig(chart_path)

            # Adicionar o capítulo ao PDF
            pdf.chapter(title, chart_path)

        # Salvar o PDF
        pdf.output(r"C:\Users\HELIO\Documents\relatorio.pdf")

        path = r"C:\Users\HELIO\Documents\relatorio.pdf"

        # Caminho completo para o arquivo PDF
        caminho_pdf = path

        # Abrir o arquivo PDF no navegador padrão
        webbrowser.open(caminho_pdf)
        
        msg_box = QMessageBox()
        msg_box.setText("A função terminou de executar!")
        msg_box.exec_()
    
    def header(self):
        # Configurar o cabeçalho do PDF
        self.set_font('Arial', 'B', 12)
        self.cell(0, 10, 'Relatório de Evolução de Preço', 0, 1, 'C')

    def chapter_title(self, title):
        # Configurar o título do capítulo
        self.set_font('Arial', 'B', 12)
        self.cell(0, 10, title, 0, 1, 'L')
        self.ln(5)

    def chapter_body(self, chart_path):
        # Adicionar o gráfico ao PDF
        self.image(chart_path, x=10, y=None, w=290)

    def chapter(self, title, chart_path):
        # Adicionar um capítulo ao PDF
        self.add_page()
        self.chapter_title(title)
        self.chapter_body(chart_path)

class App(QMainWindow):
    def __init__(self):
        global versao, pickles, dados, name, username
        super().__init__()
        self.setWindowTitle(f"Análise de Ativos - {versao}")
        self.resize(1200, 900)

        #Define a cor de fundo
        cor_fundo = QColor(256, 256, 256)

        #Configura a cor de fundo usando o estilo CSS
        self.setStyleSheet(f"background-color: {cor_fundo.name()}; color: white; font: 17SSpx;")
        
        # Criar o widget principal
        main_widget = QWidget(self)
        self.setCentralWidget(main_widget)

        # Criar os layouts
        main_layout = QVBoxLayout(main_widget)
        uptop_layout = QHBoxLayout()
        semitop_layout = QHBoxLayout()
        top_layout = QHBoxLayout()
        bottom_layout = QHBoxLayout()

        # Criar o dropdown para selecionar o ativo
        self.dropdown_label = QLabel("Selecione um Ativo:")
        self.dropdown = QComboBox()

        #Adicionar um ativo
        self.ativo_label = QLabel("Nome do Ativo: ")
        self.nomeAtivo = QLineEdit()
        self.nomeAtivo.setFixedWidth(80)

        self.cotas_label = QLabel("Quantas cotas: ")
        self.cotas = QLineEdit()
        self.cotas.setFixedWidth(80)

        self.dataCompra_label = QLabel("Data de compra (MM/DD/AAAA): ")
        self.dataCompra = QLineEdit()
        self.dataCompra.setFixedWidth(80)

        self.dividendos_label = QLabel("Valor de dividendos: ")
        self.dividendos = QLineEdit()
        self.dividendos.setFixedWidth(80)

        self.frequencia_label = QLabel("Qual a frequencia de pagamento")
        self.frequencia = QComboBox()
        self.frequencia.addItems(["A","S","T","B","M"])

        self.add_ativo_button = QPushButton("Adicionar Ativo")

        self.nomeUsuarioLabel = QLabel("Nome do Usuário")
        self.nomeUsuario =  QLineEdit()
        self.nomeUsuario.setText(username)
        self.carregarBotao = QPushButton("Carregar Usuário")
        self.carregarBotao.clicked.connect(lambda: self.carregar_usuario(self.nomeUsuario.text()))

        self.salvarProgresso = QPushButton("Salvar Progresso")
        self.salvarProgresso.clicked.connect(partial(self.salvar_progresso))

        #Retirar um ativo
        self.remover_label = QLabel("Selecione um Ativo para remover")
        self.remover_ativo = QComboBox()
        self.remove_ativo_button = QPushButton("Remover Ativo")

        #self.nomeAtivo.textChanged.connect(update_label)

        self.add_ativo_button.clicked.connect(partial(self.addAtivo))
        self.remove_ativo_button.clicked.connect(partial(self.removerAtivo))
        
        # Criar o botão para atualizar o gráfico
        self.atualizar_btn = QPushButton("Atualizar Gráfico")
        self.atualizar_btn.clicked.connect(self.atualizar_grafico)

        uptop_layout.addWidget(self.ativo_label)
        uptop_layout.addWidget(self.nomeAtivo)
        uptop_layout.addWidget(self.cotas_label)
        uptop_layout.addWidget(self.cotas)
        uptop_layout.addWidget(self.dataCompra_label)
        uptop_layout.addWidget(self.dataCompra)
        uptop_layout.addWidget(self.dividendos_label)
        uptop_layout.addWidget(self.dividendos)
        uptop_layout.addWidget(self.frequencia_label)
        uptop_layout.addWidget(self.frequencia)
        uptop_layout.addWidget(self.add_ativo_button)

        semitop_layout.addWidget(self.remover_label)
        semitop_layout.addWidget(self.remover_ativo)
        semitop_layout.addWidget(self.remove_ativo_button)
        semitop_layout.addWidget(self.carregarBotao)
        semitop_layout.addWidget(self.salvarProgresso)

        # Adicionar os widgets ao layout superior
        top_layout.addWidget(self.dropdown_label)
        top_layout.addWidget(self.dropdown)
        top_layout.addWidget(self.atualizar_btn)

        # Criar a tabela com scroll
        self.tabela = QTreeWidget()
        self.tabela.setHeaderLabels(df_ativos.columns.tolist())
        self.tabela.header().setSectionResizeMode(QHeaderView.ResizeToContents)
        self.tabela.setStyleSheet(f"QHeaderView::section {{ background-color: black ; border-radius: 10px; text-align: center;}};")
        self.tabela.header().setDefaultAlignment(Qt.AlignCenter) 

        # Preencher a tabela com os dados do DataFrame
        for _, row in df_ativos.iterrows():
            item = QTreeWidgetItem(self.tabela)
            for i in range(14):
                item.setTextAlignment(i, Qt.AlignCenter)

            for i in range(len(row)):
                item.setText(i, str(row[i]))

        # Criar o widget FigureCanvasQtAgg para exibir o gráfico
        self.fig = Figure()
        self.ax = self.fig.add_subplot(111)
        self.canvas = FigureCanvas(self.fig)

        # Adicionar os widgets ao layout inferior
        bottom_layout.addWidget(self.canvas)

        self.download = QPushButton("Gerar Relatório")
        self.download.clicked.connect(partial(self.gerar_relatorio_pdf))

        # Adicionar os layouts ao layout principal
        main_layout.addLayout(uptop_layout)
        main_layout.addLayout(semitop_layout)
        main_layout.addLayout(top_layout)
        main_layout.addWidget(self.download)
        main_layout.addWidget(self.tabela)
        main_layout.addLayout(bottom_layout)

        # Configurar o widget FigureCanvasQtAgg
        self.ax.set_xlabel('Data')
        self.ax.set_ylabel('Preço de Fechamento')
        self.ax.tick_params(axis='x', rotation=15)

        # Exibir a tabela e o gráfico na janela principal
        self.tabela.show()
        self.canvas.show()

    @Slot()

    def show_dialog(self, message):
        dialog = QMessageBox()
        dialog.setWindowTitle("Mensagem")
        dialog.setText(message)
        dialog.exec_()

    def removerAtivo(self):
        ativo = self.remover_ativo.currentText()
        indices_remover = df_ativos[df_ativos['Ativo'] == ativo].index
        df_ativos.drop(indices_remover, inplace=True)
        
        self.dropdown.removeItem(self.dropdown.findText(ativo))
        self.remover_ativo.removeItem(self.remover_ativo.findText(ativo))
        try:
            self.update_table()
        except:
            pass

        msg_box = QMessageBox()
        msg_box.setText("A função terminou de executar!")
        msg_box.exec_()

    def salvar_progresso(self):
        global df_ativos

        try:
            # Converter a coluna 'Data Compra' para o formato desejado (por exemplo, '%Y-%m-%d')
            df_ativos['Data Compra'] = pd.to_datetime(df_ativos['Data Compra'], format='%Y-%m-%d').dt.strftime('%Y-%m-%d')
            
            # Converta o DataFrame df_ativos para um dicionário
            user_data = df_ativos.to_dict()
            ref = db.reference(f'usuarios/{username}')
            ref.set(user_data)

            print(f"Progresso do usuário '{username}' salvo com sucesso no Firebase.")
        except Exception as e:
            print(f"Erro ao salvar progresso do usuário '{username}' no Firebase: {str(e)}")

    def carregar_usuario(self, username):
        global Users, arquivoPickle, dados, df_ativos

        try:
            # Recupere os dados do Firebase para o usuário específico
            ref = db.reference(f'usuarios/{username}')
            user_data = ref.get()

            if user_data:
                df_ativos = pd.DataFrame(user_data)
                self.show_dialog("Usuario Válido!")
                self.atualizar_dados(df_ativos)
                self.atualizar_dropdowns()

            if len(df_ativos) == 0:
                global assets
                df_ativos = pd.DataFrame(assets)
                self.atualizar_dados(df_ativos)
                self.atualizar_dropdowns()
                # Restante do código para atualizar a tabela, etc.
            else:
                print(f"Usuário '{username}' não encontrado no Firebase.")
        except Exception as e:
            print(f"Erro ao carregar usuário '{username}' do Firebase: {str(e)}")
    
    def atualizar_grafico(self):
        ativo_selecionado = self.dropdown.currentText()
        df_filtrado = df_ativos[df_ativos['Ativo'] == ativo_selecionado]
        symbol = df_filtrado['Ativo'].iloc[0]
        today = datetime.today()
        start_date = today - timedelta(days=365)

        try:
            # Obter os dados históricos da ação usando o yfinance
            title = f"Evolução de Preço - {symbol}"
            chart_path = f"chart_{symbol}.png"

            # Obter os dados históricos da ação usando o yfinance
            df = yf.download(symbol, start=today - timedelta(days=180), end=today)

            # Limpar o gráfico anterior
            self.ax.clear()

            Open = []
            High = []
            Low = []
            Close = []

            for i in range(len(df)):
                Open.append(df['Open'].iloc[i])
                High.append(df['High'].iloc[i])
                Low.append(df['Low'].iloc[i])
                Close.append(df['Close'].iloc[i])

            data = pd.DataFrame({
                "Open": Open,
                "High": High,
                "Low" : Low,
                "Close": Close
            }, index=df.index)

            mpf.plot(data, type='candle', style='yahoo', ax=self.ax)

            # Atualizar o gráfico exibido no widget FigureCanvasQtAgg
            self.canvas.draw()

            msg_box = QMessageBox()
            msg_box.setText("A função terminou de executar!")
            msg_box.exec_()
            self.salvar_progresso()

        except Exception as e:
            print(f"Erro ao obter os dados da ação: {e}")

    def atualizar_dropdowns(self):
        try:
            ativos = df_ativos['Ativo']
            self.dropdown.clear()
            self.remover_ativo.clear()
            for ativo in ativos:
                self.dropdown.addItem(ativo)
                self.remover_ativo.addItem(ativo)
        except:
            pass

    def atualizar_dados(self, df):

        global df_ativos
        
        df_ativos = pd.DataFrame(df)
        try:
            ativos = df_ativos['Ativo']
            self.dropdown.addItems(ativos)
            self.remover_ativo.addItems(ativos)
        except:
            pass
        
        self.show_dialog(f"{df_ativos}")

        self.update_table()

        msg_box = QMessageBox()
        msg_box.setText("Dados Atualizados!")
        msg_box.exec_()

        self.salvar_progresso()

        return df_ativos

    def addAtivo(self):

        global df_ativos, assets, data_compra, Dividendos, frequencias
        msg_box = QMessageBox()
        nomeAtivo = self.nomeAtivo.text()
        Cotas = self.cotas.text()
        Cotas = int(Cotas)
        dataCompra = pd.to_datetime(self.dataCompra.text())
        dividendos = float(self.dividendos.text())
        frequencia_dropdown = str(self.frequencia.currentText())

        msg_box.setText(f"{nomeAtivo} , {Cotas}, {dataCompra}, {dividendos}, {frequencia_dropdown}")
        msg_box.exec_()

        #Análise inicial do ativo
        try:
            symbol = nomeAtivo
            dados = yf.download(symbol, start=(today - timedelta(180)), end=today)
            dados['Returns'] = dados['Close'].pct_change()
            volatility = dados['Returns'].std() * (252 ** 0.5)  # Assumindo 252 dias de negociação em um ano
            result = volatility
            volatilidade = (f"{result:.0%}")
            stock = yf.Ticker(symbol)
            latest_price = stock.history(period='1d')['Close'][0]
            estimate = float(round(latest_price, 2))
            Rentabilidade = round(((dividendos / estimate) * 100), 2)

            # Criação da caixa de mensagem
            msg_box = QMessageBox()
            msg_box.setIcon(QMessageBox.Question)
            msg_box.setWindowTitle("Opções")
            msg_box.setText(f"Deseja adicionar o ativo? A Ação {symbol} com o preço atual de {estimate}, tem uma rentabilidade de {Rentabilidade}% e uma volatilidade de {volatilidade}")
            msg_box.setStandardButtons(QMessageBox.Yes | QMessageBox.No)
            msg_box.setDefaultButton(QMessageBox.Yes)

            opcao = msg_box.exec_()

            if opcao == QMessageBox.Yes:
                try:
                    start_date = pd.to_datetime(dataCompra)
                    end_date = pd.to_datetime("today")
                    df = yf.download(nomeAtivo, start=start_date, end=end_date)
                    df['Returns'] = df['Close'].pct_change()
                    volatility = df['Returns'].std() * (252 ** 0.5)
                    result = volatility * Cotas
                    volatilidades = (f"{result:.0%}")
                except:
                    volatilidades = None

                try:
                    stock = yf.Ticker(nomeAtivo)
                    latest_price = stock.history(period='1d')['Close'][0]
                    preco = round(latest_price, 2)
                    hist = stock.history(period='20d')['Close'][0]
                    hist = round(hist, 2)
                except:
                    preco = None
                    hist = None

                quantMin = round(((preco / dividendos)), 0)
                resultado = round((((preco - hist) / hist) * 100), 2)

                days = (today - dataCompra).days

                try:
                    stock = yf.Ticker(nomeAtivo)
                    latest_price = stock.history(period=f"{days}d")['Close'][0]
                    precoCompra = round(latest_price, 2)
                except:
                    pass

                variacao = round(((preco - precoCompra) / precoCompra * 100), 0).astype(str) + "%"

                novo_ativo = {
                    'Ativo': nomeAtivo,
                    'Cotas': int(Cotas),
                    'Dividendos': dividendos,
                    'Volatilidade': volatilidades,
                    'Preco': preco,
                    'Quant Min': quantMin,
                    'Progresso': round(((Cotas / quantMin) * 100), 0),
                    'Data Compra': dataCompra,
                    'Historico': hist,
                    'Resultado': resultado,
                    'Total': round(Cotas * preco,2),
                    'Frequencia': frequencia_dropdown, 
                    'Div Total': dividendos * Cotas,
                    'Rentabilidade(%)' : f"{Rentabilidade}%",
                    'Preco Compra': precoCompra,
                    'ROI': variacao
                } 

                df_ativos = df_ativos._append(novo_ativo, ignore_index=True)
                
                self.dropdown.addItem(nomeAtivo)
                self.remover_ativo.addItem(nomeAtivo)
                self.update_table()

                msg_box = QMessageBox()
                msg_box.setText(f"Ativo {symbol} Adicionado")
                msg_box.exec_()
        except:
            self.show_dialog("Ação Não Econtrada")

        else:
            msg_box.setText(f"Ativo {symbol} não Adicionado")

    def update_table(self):
        msg_box = QMessageBox()
        self.tabela.clear()

        for _, row in df_ativos.iterrows():
            item = QTreeWidgetItem(self.tabela)
            item.setTextAlignment(0, Qt.AlignLeft)
            item.setTextAlignment(1, Qt.AlignLeft)
            item.setTextAlignment(2, Qt.AlignCenter)
            item.setTextAlignment(3, Qt.AlignCenter)
            item.setTextAlignment(4, Qt.AlignCenter)
            item.setTextAlignment(5, Qt.AlignCenter)
            item.setTextAlignment(6, Qt.AlignCenter)
            item.setTextAlignment(7, Qt.AlignCenter)
            item.setTextAlignment(8, Qt.AlignCenter)
            item.setTextAlignment(9, Qt.AlignCenter)
            item.setTextAlignment(10, Qt.AlignCenter)
            item.setTextAlignment(11, Qt.AlignCenter)
            item.setTextAlignment(12, Qt.AlignCenter)
            item.setTextAlignment(13, Qt.AlignCenter)

            for i in range(len(row)):
                item.setText(i, str(row[i]))
        
        self.tabela.show

        msg_box.setText("Tabela Atualizada")
        msg_box.exec_()
        self.atualizar_dropdowns()

    def gerar_relatorio_pdf(self):
        global df_ativos, name
        self.salvar_progresso()
        PDF.inicio_pdf(df_ativos, name)

class LoginDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Login")
        self.setFixedSize(300, 150)

        layout = QVBoxLayout(self)

        self.username_label = QLabel("Nome de usuário:")
        self.username_input = QLineEdit()
        self.login_button = QPushButton("Login")

        self.login_button.clicked.connect(self.login)

        layout.addWidget(self.username_label)
        layout.addWidget(self.username_input)
        layout.addWidget(self.login_button)

    def login(self):
        global username
        username = self.username_input.text()

        if username:
            if check_user_exists(username):
                app = App()
                msg = QMessageBox()
                msg.setText(f"{name}")
                msg.exec_()
                app.carregar_usuario(username)
                self.accept()
            else:
                create_new_user(username)
                self.accept()
        else:
            QMessageBox.warning(self, "Login", "Por favor, insira um nome de usuário válido.")

def check_user_exists(username):
    if username in pickles:
        name = username
        msg_box = QMessageBox()
        msg_box.setText(f"Bem vindo de volta {name}")
        msg_box.exec_()
        return True

def create_new_user(username):
    nome = username
    assets = []
    data_compra = []
    Dividendos = []
    frequencias = []
    Users.append(username)
    usuario = dados[username] = [assets, 
                                data_compra, 
                                Dividendos, 
                                frequencias]
    nome = str(username)
    pickles[nome] = usuario
    with open(arquivoPickle, "wb") as arquivo:
        pickle.dump(pickles, arquivo)
    pass

# Configurar e exibir a janela principal
if __name__ == "__main__":
    app = QApplication([])
    login_dialog = LoginDialog()

    if login_dialog.exec_() == QDialog.Accepted:
        main_window = App()
        main_window.salvar_progresso()
        main_window.show()

    app.exec_()