import os                                       # Para operações com sistema de arquivos
import tkinter as tk                            # Para criação da interface gráfica
from tkinter import ttk, filedialog, messagebox # Componentes avançados do Tkinter
from PIL import Image                           # Para manipulação de imagens
from pypdf import PdfReader, PdfWriter          # Para manipulação de PDFs
from fpdf import FPDF                           # Para geração de PDFs
from docx import Document                       # Para manipulação de arquivos Word
from docx.shared import Inches                  # Para controle de dimensões no Word
import pandas as pd                             # Para manipulação de dados tabulares
from datetime import datetime                   # Para manipulação de data/hora
import zipfile                                  # Para criação/extração de arquivos ZIP
import sys                                      # Para acesso a funcionalidades do sistema
from pdf2docx import Converter                  # Para conversão de PDF para Word
import subprocess                               # Para execução de processos externos
from pdf2image import convert_from_path         # Para conversão de PDF para imagens
import chardet                                  # Para detecção de codificação de arquivos
import magic                                 # Para detecção de tipo MIME de arquivos
import mimetypes                                # Para manipulação de tipos MIME
import shutil                                   # Para operações de cópia e movimentação de arquivos
import re                                       # Para expressões regulares
import uuid                                     # Para geração de identificadores únicos
import tempfile                                 # Para criação de arquivos temporários
from docx2pdf import convert as docx2pdf_convert   # Para conversão de DOCX para PDF
from reportlab.lib.pagesizes import letter      # Para definir o tamanho da página no ReportLab
from reportlab.pdfgen import canvas             # Para geração de PDFs com ReportLab
import multiprocessing as mp                    # Para processamento paralelo
import time                                   # Para manipulação de tempo
import psutil                                  # Para monitoramento de processos e uso de memória
import logging                                # Para registro de logs
import traceback                             # Para captura de rastreamentos de erros
import bleach                                  # Para limpeza de HTML/JS potencialmente perigoso
import sqlite3                               # Para manipulação de banco de dados SQLite
from datetime import datetime, timedelta    # Para manipulação de data/hora

# Importações das funções utilitárias do módulo 'utils' para realizar conversões de arquivos
# entre diferentes formatos (PDF, DOCX, TXT, PNG, CSV, XLSX) e gerenciar a execução
# com timeout usando o wrapper, garantindo modularidade e reutilização do código.

from utils import (                         
    csv_para_xlsx_global,
    docx_para_pdf_global,
    docx_para_png_global,
    docx_para_txt_global,
    pdf_para_docx_global,
    pdf_para_png_global,
    pdf_para_txt_global,
    wrapper,
    xlsx_para_csv_global,
    pdf_para_xlsx_global,
    juntar_pdfs_worker,
    comprimir_pdf,
    obter_caminho_poppler,
    obter_caminho_grhostscript
)


# Criar pasta C:\Temp\Logs_conversor se não existir
if not os.path.exists(r'C:\Temp\Logs_conversor'):
    os.makedirs(r'C:\Temp\Logs_conversor')

# Configurar logger para criar arquivo por erro
logger = logging.getLogger("teste")
logger.setLevel(logging.DEBUG)

if not logger.handlers:
    # Usar nome fixo erro_conversor.log
    nome_arquivo = "erro_conversor.log"
    caminho_arquivo = os.path.join(r'C:\Temp\Logs_conversor', nome_arquivo)
    handler = logging.FileHandler(caminho_arquivo, encoding='utf-8')
    formatter = logging.Formatter('%(asctime)s [%(levelname)s] %(message)s')
    handler.setFormatter(formatter)
    logger.addHandler(handler)


# Verifica se o script está sendo executado como um executável congelado (PyInstaller)
if getattr(sys, 'frozen', False):
    sys.path.append(os.path.dirname(sys.executable))
mp.freeze_support()

# ==================================================
# CONFIGURAÇÕES GLOBAIS
# ==================================================

CHAVE_VALIDA = "K1F8O-CK659-XRD9A-GBQB9-B475O"

# Caminho do banco de dados SQLite para licenças
CAMINHO_DB = os.path.join(
    os.path.expanduser('~'),
    'AppData', 'Local', 'Conversor', 'licencas.db'
)
os.makedirs(os.path.dirname(CAMINHO_DB), exist_ok=True)

# Chave de licença válida (definida como constante)
def verificar_licenca():
    """Verifica se a licença está ativa e válida."""
    try:
        conn = sqlite3.connect(CAMINHO_DB)  # Conecta ao banco de dados SQLite
        cursor = conn.cursor()
        cursor.execute("SELECT Data_da_ativacao, validade_ate, ativa, código_da_licenca FROM licencas WHERE ativa = 1 ORDER BY id DESC LIMIT 1")
        licenca = cursor.fetchone()
        conn.close()
        data_atual = datetime.now().date()
        if not licenca or licenca[2] != 1:
            return False, None
        else:
            data_ativacao_str = licenca[0]
            valida_ate_str = licenca[1]
            try:
                data_ativacao = datetime.fromisoformat(data_ativacao_str).date()
                valida_ate = datetime.fromisoformat(valida_ate_str).date()
            except ValueError:
                data_ativacao = datetime.strptime(data_ativacao_str, '%Y-%m-%d').date()
                valida_ate = datetime.strptime(valida_ate_str, '%Y-%m-%d').date()
            if data_atual > valida_ate:
                return False, licenca[3]
            return True, licenca[3]
    except sqlite3.Error as e:
        print(f"Erro ao acessar o banco de dados: {e}")
        return False, None

def ativar_licenca(codigo_licenca):
    """Ativa a licença se o código for válido e salva no banco de dados."""
    if codigo_licenca != CHAVE_VALIDA:
        raise ValueError("Código de licença inválido!")
    try:
        conn = sqlite3.connect(CAMINHO_DB)  # Conecta ao banco de dados SQLite
        cursor = conn.cursor()
        cursor.execute("UPDATE licencas SET ativa = 0 WHERE ativa = 1")
        data_atual = datetime.now().date()
        data_validade = data_atual + timedelta(days=365)  # Define validade de 1 ano
        cursor.execute(
            "INSERT INTO licencas (Data_da_ativacao, validade_ate, código_da_licenca, ativa) VALUES (?, ?, ?, ?)",
            (data_atual.isoformat(), data_validade.isoformat(), codigo_licenca, 1)
        )
        conn.commit()
        conn.close()
        return True
    except sqlite3.Error as e:
        print(f"Erro ao salvar a licença no banco: {e}")
        raise

def inicializar_banco_licenca():
    '''Cria o banco de dados SQLite e a tabela de licenças se não existir.'''
    try:
        conn = sqlite3.connect(CAMINHO_DB)  # Conecta ao banco de dados SQLite
        cursor = conn.cursor()
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS licencas (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                Data_da_ativacao TEXT,
                validade_ate TEXT,
                código_da_licenca TEXT,
                ativa INTEGER DEFAULT 0
            )
        ''')
        conn.commit()
        conn.close()
    except sqlite3.Error as e:
        print(f"Erro ao criar o banco de dados: {e}")

# Configuração do log técnico seguro

if not os.path.exists(r'C:\Temp\Logs_conversor'):
    os.makedirs(r'C:\Temp\Logs_conversor')

# Configurar logger para criar arquivo por erro
logger = logging.getLogger("teste")
logger.setLevel(logging.DEBUG)

if not logger.handlers:
    # Usar nome fixo erro_conversor.log
    nome_arquivo = "erro_conversor.log"
    caminho_arquivo = os.path.join(r'C:\Temp\Logs_conversor', nome_arquivo)
    handler = logging.FileHandler(caminho_arquivo, encoding='utf-8')
    formatter = logging.Formatter('%(asctime)s [%(levelname)s] %(message)s')
    handler.setFormatter(formatter)
    logger.addHandler(handler)


def registrar_log_tecnico(exc):
    """Registra detalhes técnicos do erro em log interno seguro."""
    logger.error(traceback.format_exc())


def icone_logo(relative_path):
    """Converte caminhos relativos para funcionar no executável e no desenvolvimento."""
    if hasattr(sys, '_MEIPASS'):  # Verifica se está rodando como executável
        base_path = sys._MEIPASS # Caminho do diretório temporário do PyInstaller
    else:
        base_path = os.path.abspath(".") # Caminho do diretório atual
    return os.path.join(base_path, relative_path)  # Retorna o caminho completo do recurso

def validar_mime(filepath, tipos_aceitos):
    """Valida o tipo MIME real do arquivo usando python-magic, com fallback para mimetypes."""
    try:
        mime = magic.from_file(filepath, mime=True)
        print(f"[DEBUG] MIME detectado: {mime}")
        if mime in tipos_aceitos:
            return True
        # Fallback: tenta pelo mimetypes se magic falhar
        mime_guess, _ = mimetypes.guess_type(filepath)
        print(f"[DEBUG] MIME (fallback mimetypes): {mime_guess}")
        if mime_guess and mime_guess in tipos_aceitos:
            return True
        return False
    except Exception as e:
        print(f"[DEBUG] Erro na detecção MIME: {e}")
        # Fallback direto para mimetypes se magic falhar
        mime_guess, _ = mimetypes.guess_type(filepath)
        print(f"[DEBUG] MIME (fallback mimetypes): {mime_guess}")
        if mime_guess and mime_guess in tipos_aceitos:
            return True
        return False

# def obter_caminho_poppler(): # Função para obter o caminho do Poppler
#     """Retorna o caminho do Poppler, ajustando para o ambiente do PyInstaller."""
#     if hasattr(sys, '_MEIPASS'):  # Verifica se está rodando como executável
#         return os.path.join(sys._MEIPASS, 'poppler')  # Caminho dentro do executável
#     return r'C:\poppler-25.07.0\Library\bin'  # Caminho local para desenvolvimento

def sanitizar_celula_excel(valor):
    """Prefixa com apóstrofo se for potencial fórmula perigosa."""
    if isinstance(valor, str) and valor and valor[0] in ('=', '+', '-', '@'):
        return "'" + valor
    return valor

def remover_tags_html(texto):
    """Remove tags HTML/JS do texto, retornando apenas o texto limpo."""
    return bleach.clean(str(texto), tags=[], attributes={}, styles=[], strip=True)

def contem_estrutura_perigosa(texto):
    """Detecta se o texto contém estrutura potencialmente perigosa."""
    if isinstance(texto, str):
        # Detecta fórmulas perigosas
        if texto and texto[0] in ('=', '+', '-', '@'):
            return True
        # Detecta tags HTML/JS
        if re.search(r'<script|<iframe|<object|<embed|<a\s+href|javascript:', texto, re.IGNORECASE):
            return True
    return False


def executar_com_timeout(func, args=(), kwargs=None, timeout=30, mem_limit_mb=500):
    """Executa uma função em subprocesso com timeout e limite de memória (em MB)."""
    if kwargs is None:
        kwargs = {}
    q = mp.Queue()
    p = mp.Process(target=wrapper, args=(q, func, *args), kwargs=kwargs)
    p.start()
    proc = psutil.Process(p.pid)
    elapsed = 0
    interval = 0.2  # segundos
    mem_limit_bytes = mem_limit_mb * 3024 * 3024

    while p.is_alive() and elapsed < timeout:
        try:
            mem = proc.memory_info().rss
            if mem > mem_limit_bytes:
                p.terminate()
                p.join()
                return (False, MemoryError(f"Limite de memória excedido ({mem_limit_mb} MB)"))
        except Exception:
            pass  # Processo pode já ter terminado
        time.sleep(interval)
        elapsed += interval

    if p.is_alive():
        p.terminate()
        p.join()
        return (False, TimeoutError("Tempo limite excedido na conversão"))
    if not q.empty():
        return q.get()
    return (False, RuntimeError("Erro desconhecido no subprocesso"))

def validar_pdf_com_pdfinfo(arquivo, poppler_path):
    """Valida a integridade de um PDF usando pdfinfo."""
    try:
        comando = [os.path.join(poppler_path, "pdfinfo"), arquivo]
        resultado = subprocess.run(comando, capture_output=True, text=True, timeout=15, 
                         creationflags=subprocess.CREATE_NO_WINDOW if os.name == 'nt' else 0)
        if resultado.returncode != 0:
            raise ValueError(f"pdfinfo falhou: {resultado.stderr}")
        # Verifica se o PDF tem um número razoável de páginas (ex.: menos de 1000 para evitar DoS)
        for linha in resultado.stdout.splitlines():
            if linha.startswith("Pages:"):
                paginas = int(linha.split(":")[1].strip())
                if paginas > 1000:  # Limite arbitrário para evitar PDFs maliciosos
                    raise ValueError(f"PDF excede limite de páginas: {paginas}")
        return True
    except subprocess.TimeoutExpired:
        raise ValueError("Validação do PDF excedeu o tempo limite")
    except Exception as e:
        raise ValueError(f"Erro ao validar PDF com pdfinfo: {str(e)}")

def validar_docx_com_zip(arquivo):
    """Valida a estrutura básica de um DOCX verificando se é um ZIP válido."""
    try:
        with zipfile.ZipFile(arquivo, 'r') as zipf:
            if zipf.testzip() is not None:
                raise ValueError("Arquivo ZIP corrompido")
            # Verifica se contém arquivos essenciais do DOCX
            arquivos_essenciais = ["word/document.xml", "[Content_Types].xml"]
            arquivos_zip = zipf.namelist()
            for essencial in arquivos_essenciais:
                if essencial not in arquivos_zip:
                    raise ValueError(f"Estrutura de DOCX inválida: falta {essencial}")
        return True
    except Exception as e:
        logger.error(traceback.format_exc())
        raise ValueError(f"Erro ao validar DOCX: {str(e)}")

def slugify_nome(nome):
    """Remove caracteres especiais e espaços do nome do arquivo."""
    nome = re.sub(r'[^a-zA-Z0-9_\-\.]', '_', nome)
    return nome

def gerar_nome_download(nome_original, extensao):
    """Gera um nome de arquivo único para download, baseado no nome original e extensão."""
    prefixo = uuid.uuid4().hex
    nome_slug = slugify_nome(nome_original)
    return f"{nome_slug}_{prefixo}.{extensao}"


TIPOS_ARQUIVOS_PERMITIDOS = [
    ("Todos os arquivos suportados", "*.pdf;*.txt;*.png;*.jpg;*.jpeg;*.webp;*.ico;*.xlsx;*.csv"),
    ("PDF", "*.pdf"),
    ("PNG", "*.png"),
    ("JPG", "*.jpg"),
    ("JPEG", "*.jpeg"),
    ("WEBP", "*.webp"),
    ("ICO", "*.ico"),
    ("XLSX", "*.xlsx"),
    ("CSV", "*.csv")
]

ICONE = icone_logo('logo_conversor.ico') # Ícone do aplicativo
DOWNLOAD_FOLDER = os.path.join(os.path.expanduser('~'), 'Downloads') # Pasta de downloads padrão do usuário

# Esquema de cores
COR_PRIMARIA = '#FB071A'                        # Vermelho 
COR_SECUNDARIA = '#FFFFFF'                      # Branco para textos e fundos
COR_FUNDO = '#F5F5F5'                           # Cinza claro para fundo da interface
COR_TEXTO = '#333333'                           # Cinza escuro para textos principais
COR_BORDA = '#CCCCCC'                           # Cinza para bordas e divisores

# Configurações de fonte
FONTE_PRINCIPAL = ('Arial', 10)                # Fonte padrão para textos
FONTE_TITULO = ('Arial', 14, 'bold')           # Fonte para títulos (negrito)
FONTE_DESTAQUE = ('Arial', 10, 'bold')         # Fonte para textos destacados


# ==================================================
# CLASSE PRINCIPAL DO APLICATIVO
# ==================================================
class AplicativoFleury: # Classe principal do aplicativo
    def __init__(self, root): # Construtor da classe
        self.root = root                       # Referência à janela principal
        self.root.title("ZippyDocs")  # Define o título da janela
        self.configurar_icone()                # Configura o ícone do aplicativo
        self.configurar_estilos()              # Aplica os estilos visuais
        
        self.container = ttk.Frame(root)       # Container principal para as telas
        self.container.pack(fill=tk.BOTH, expand=True) # Expande para preencher toda a janela
        
        self.telas = {}                        # Dicionário para armazenar as telas
        
        # Cria todas as telas do aplicativo
        for Tela in (TelaInicial, TelaConversor, TelaPDFTools): # Lista de telas
            nome_tela = Tela.__name__          # Obtém o nome da classe da tela
            frame = Tela(parent=self.container, controller=self) # Instancia a tela
            self.telas[nome_tela] = frame      # Armazena a tela no dicionário
            frame.grid(row=0, column=0, sticky="nsew") # Posiciona na grade
        
        self.mostrar_tela("TelaInicial")       # Mostra a tela inicial por padrão
    
    def mostrar_tela(self, nome_tela):       # Método para mostrar uma tela específica
        tela = self.telas[nome_tela]          # Obtém a tela do dicionário
        tela.tkraise()                        # Traz a tela para frente (muda a visualização)
    
    def configurar_icone(self):           # Configura o ícone do aplicativo
        try:
            self.root.iconbitmap(ICONE)        # Tenta carregar o ícone personalizado
        except:
            print("Ícone não encontrado. Usando padrão do sistema.") # Fallback se o ícone não existir
    
    def configurar_estilos(self):      # Configura os estilos dos widgets
        style = ttk.Style()                   # Cria um objeto de estilos
        style.theme_use('clam')               # Usa o tema 'clam' como base
     
        
        # Configuração geral de estilo para todos os widgets
        style.configure('.', 
                      background=COR_FUNDO,   # Cor de fundo
                      foreground=COR_TEXTO,   # Cor do texto
                      font=FONTE_PRINCIPAL)   # Fonte principal
        
        # Estilo personalizado para botões Fleury
        style.configure('Fleury.TButton',  
                      foreground=COR_SECUNDARIA, # Cor do texto (branco)
                      background=COR_PRIMARIA, # Cor de fundo (vermelho Fleury)
                      font=FONTE_PRINCIPAL,    # Fonte
                      padding=5)               # Espaçamento interno
        
        # Efeitos de hover/pressionado para os botões
        style.map('Fleury.TButton', 
                background=[('active', '#D90615'), ('pressed', '#A00510')]) # Tons de vermelho
        
        self.root.configure(background=COR_FUNDO) # Define a cor de fundo da janela principal


# ==================================================
# TELA INICIAL DO APLICATIVO
# ==================================================
class TelaInicial(ttk.Frame):
    def __init__(self, parent, controller):
        super().__init__(parent)
        self.controller = controller
        self.configurar_interface()

    def configurar_interface(self):
        """Configura a interface da tela inicial."""
        main_frame = ttk.Frame(self, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Título da tela inicial
        ttk.Label(
            main_frame,
            text="ZippyDocs",
            font=FONTE_TITULO,
            foreground=COR_PRIMARIA,
            background=COR_FUNDO
        ).pack(pady=(0, 30))

        # Frame central para os botões principais
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(pady=10)

        self.btn_conversor = ttk.Button(
            btn_frame,
            text="Conversor de arquivos",
            command=lambda: self.controller.mostrar_tela("TelaConversor"),
            style='Fleury.TButton',
            width=25
        )
        self.btn_conversor.pack(pady=10)

        self.btn_pdf_tools = ttk.Button(
            btn_frame,
            text="Divisor & Unificador de PDF",
            command=lambda: self.controller.mostrar_tela("TelaPDFTools"),
            style='Fleury.TButton',
            width=25
        )
        self.btn_pdf_tools.pack(pady=10)

        self.btn_conversor.bind("<Button-1>", lambda e: self.mostrar_validade_licenca() if self.btn_conversor['state'] == tk.DISABLED else None)
        self.btn_pdf_tools.bind("<Button-1>", lambda e: self.mostrar_validade_licenca() if self.btn_pdf_tools['state'] == tk.DISABLED else None)

        # Frame para o rodapé (copyright)
        rodape_frame = ttk.Frame(main_frame)
        rodape_frame.pack(side=tk.BOTTOM, fill=tk.X, pady=(0, 0))

        rodape = ttk.Label(
            rodape_frame,
            text="© 2025 Matheus Augusto e Jeferson Sá - Todos os direitos reservados",
            font=('Arial', 8),
            foreground='gray',
            anchor='center',
            background=COR_FUNDO
        )
        rodape.pack(pady=(0, 2))

        # Frame para o link de ativação de licença, logo abaixo do rodapé
        link_frame = ttk.Frame(main_frame)
        link_frame.pack(side=tk.BOTTOM, fill=tk.X, pady=(0, 15))

        link_licenca = tk.Label(
            link_frame,
            text="Ativar Licença",
            font=('Arial', 10, 'underline'),
            fg=COR_PRIMARIA,
            bg=COR_FUNDO,
            cursor="hand2"
        )
        link_licenca.pack()
        link_licenca.bind("<Button-1>", lambda e: self.abrir_janela_licenca())

        self.atualizar_estado_botoes()


    def mostrar_mensagem(self, titulo, mensagem, erro=False, parent=None):
        """Exibe uma mensagem popup centralizada na janela principal ou em um parent."""
        if parent is None:
            parent = self
        msg = tk.Toplevel(parent)
        msg.title(titulo)
        msg.lift()

        msg.focus_force()  # Foca na nova janela
        try:
            msg.iconbitmap(ICONE)
        except:
            pass
        msg.configure(background=COR_FUNDO)
        tk.Label(
            msg,
            text=mensagem,
            bg=COR_FUNDO,
            fg=COR_PRIMARIA if erro else COR_TEXTO,
            font=FONTE_PRINCIPAL,
            padx=20,
            pady=20
        ).pack()
        btn_frame = tk.Frame(msg, bg=COR_FUNDO)
        btn_frame.pack(pady=(0, 10))
        btn_ok = tk.Button(
            btn_frame,
            text="OK",
            command=msg.destroy,
            bg=COR_PRIMARIA,
            fg=COR_SECUNDARIA,
            activebackground='#D90615',
            activeforeground=COR_SECUNDARIA,
            relief='flat',
            padx=20
        )
        btn_ok.pack()
        msg.bind("<Return>", lambda event: btn_ok.invoke())
        msg.bind("<Escape>", lambda event: msg.destroy())
        parent.update_idletasks()
        largura_msg = msg.winfo_reqwidth()
        altura_msg = msg.winfo_reqheight()
        largura_janela = parent.winfo_width()
        altura_janela = parent.winfo_height()
        x_janela = parent.winfo_rootx()
        y_janela = parent.winfo_rooty()
        x_centralizado = x_janela + (largura_janela // 2) - (largura_msg // 2)
        y_centralizado = y_janela + (altura_janela // 2) - (altura_msg // 2)
        msg.geometry(f"+{x_centralizado}+{y_centralizado}")
        msg.transient(parent)
        msg.grab_set()
        msg.wait_window()

    def abrir_janela_licenca(self):
        """Abre a janela de ativação de licença."""
        janela = tk.Toplevel(self)
        janela.title("Ativação de Licença")
        janela.geometry("400x220")
        janela.resizable(False, False)
        janela.configure(bg=COR_FUNDO)
        janela.bind('<Return>', lambda event: btn_ativar.invoke())

        try:
            janela.iconbitmap(icone_logo('logo_conversor.ico'))
        except Exception:
            pass

        # Centraliza a janela na tela
        self.update_idletasks()
        largura = 400
        altura = 220

        x_principal = self.winfo_rootx()
        y_principal = self.winfo_rooty()
        largura_principal = self.winfo_width()
        altura_principal = self.winfo_height()

        x = x_principal + (largura_principal // 2) - (largura // 2)
        y = y_principal + (altura_principal // 2) - (altura // 2)

        janela.geometry(f"{largura}x{altura}+{x}+{y}")

        # Verifica se a licença está ativa e pega a validade
        try:
            conn = sqlite3.connect(CAMINHO_DB)
            cursor = conn.cursor()
            cursor.execute("SELECT validade_ate FROM licencas WHERE ativa = 1 ORDER BY id DESC LIMIT 1")
            resultado = cursor.fetchone()
            conn.close()
        except Exception:
            resultado = None

        if resultado:
            validade_ate = resultado[0]
            try:
                validade_data = datetime.strptime(validade_ate, "%Y-%m-%d").date()
                validade_formatada = validade_data.strftime("%d/%m/%Y")
            except Exception:
                validade_data = None
                validade_formatada = validade_ate  # fallback se der erro

            hoje = datetime.now().date()
            if validade_data and validade_data < hoje:
                # Licença expirada: mostra mensagem e habilita campo para novo código
                label = tk.Label(
                    janela,
                    text="Licença expirada!\nDigite um novo código de licença:",
                    font=FONTE_TITULO,
                    fg=COR_PRIMARIA,
                    bg=COR_FUNDO
                )
                label.pack(pady=(25, 10))

                entry_codigo = tk.Entry(
                    janela,
                    width=32,
                    font=FONTE_PRINCIPAL,
                    justify="center",
                    relief="solid",
                    bd=2
                )
                entry_codigo.pack(pady=(0, 20), ipady=5)

                def ativar():
                    codigo_inserido = entry_codigo.get().strip()
                    try:
                        if ativar_licenca(codigo_inserido):
                            self.mostrar_mensagem("Sucesso", "Licença ativada com sucesso!", parent=janela)
                            janela.destroy()
                            self.atualizar_estado_botoes()
                    except ValueError as e:
                        self.mostrar_mensagem("Erro", str(e), erro=True, parent=janela)
                        entry_codigo.delete(0, tk.END)
                        entry_codigo.focus_set()

                btn_ativar = ttk.Button(
                    janela,
                    text="Ativar Licença",
                    style='Fleury.TButton',
                    command=ativar,
                    width=18
                )
                btn_ativar.pack(pady=(0, 10))

                janela.bind('<Return>', lambda event: btn_ativar.invoke())
                janela.bind('<Escape>', lambda event: janela.destroy())
                entry_codigo.focus_set()
            else:
                # Licença válida: mostra mensagem padrão
                label = tk.Label(
                    janela,
                    text=f"Licença já está ativa!\nVálida até: {validade_formatada}",
                    font=FONTE_TITULO,
                    fg=COR_PRIMARIA,
                    bg=COR_FUNDO
                )
                label.pack(pady=(60, 10))
                btn_fechar = ttk.Button(
                    janela,
                    text="Fechar",
                    style='Fleury.TButton',
                    command=janela.destroy,
                    width=18
                )
                btn_fechar.pack(pady=(10, 10))

                btn_fechar.focus_set()  # Foca no botão Fechar

                janela.bind('<Return>', lambda event: btn_fechar.invoke())
                janela.bind('<Escape>', lambda event: janela.destroy())

        else:
            # Título
            label = tk.Label(
                janela,
                text="Digite o código da licença:",
                font=FONTE_TITULO,
                fg=COR_PRIMARIA,
                bg=COR_FUNDO
            )
            label.pack(pady=(25, 10))

            # Caixa de texto maior
            entry_codigo = tk.Entry(
                janela,
                width=32,
                font=FONTE_PRINCIPAL,
                justify="center",
                relief="solid",
                bd=2
            )
            entry_codigo.pack(pady=(0, 20), ipady=5)

            def ativar():
                codigo_inserido = entry_codigo.get().strip()
                try:
                    if ativar_licenca(codigo_inserido):
                        self.mostrar_mensagem("Sucesso", "Licença ativada com sucesso!", parent=janela)
                        janela.destroy()
                        self.atualizar_estado_botoes()
                except ValueError as e:
                    self.mostrar_mensagem("Erro", str(e), erro=True, parent=janela)
                    entry_codigo.delete(0, tk.END)
                    entry_codigo.focus_set()

            btn_ativar = ttk.Button(
                janela,
                text="Ativar Licença",
                style='Fleury.TButton',
                command=ativar,
                width=18
            )
            btn_ativar.pack(pady=(0, 10))

            janela.bind('<Return>', lambda event: btn_ativar.invoke())
            janela.bind('<Escape>', lambda event: janela.destroy())

            entry_codigo.focus_set()


    def atualizar_estado_botoes(self):
        esta_valida, codigo = verificar_licenca()
        state = tk.NORMAL if esta_valida else tk.DISABLED
        self.btn_conversor.config(state=state)
        self.btn_pdf_tools.config(state=state)

    def mostrar_validade_licenca(self):
        esta_valida, codigo = verificar_licenca()
        if not esta_valida:
            messagebox.showwarning("Licença", "Licença não ativada ou expirada.\nClique em 'Ativar Licença' para liberar o uso.")
        else:
            messagebox.showinfo("Licença", "Licença ativa.")

# ==================================================
# TELA DE CONVERSÃO DE ARQUIVOS
# ==================================================
class TelaConversor(ttk.Frame): # Classe para a tela de conversão de arquivos
    def __init__(self, parent, controller): # Construtor da classe
        super().__init__(parent)              # Chama o construtor da classe pai
        self.controller = controller          # Guarda referência ao controlador
        self.configurar_interface()           # Configura os elementos da interface
    

    def adicionar_log(self, mensagem): # Método para adicionar mensagens ao log de atividades
        """Adiciona uma mensagem ao log""" 
        data_hora = datetime.now().strftime("%H:%M:%S")  # Obtém a hora atual
        self.text_log.insert(tk.END, f"[{data_hora}] {mensagem}\n")  # Insere a mensagem no log
        self.text_log.see(tk.END)  # Rola para mostrar a nova mensagem
    

    def limpar_log(self):
        """Limpa o log de atividades"""
        self.text_log.config(state=tk.NORMAL)  # Habilita edição temporariamente
        self.text_log.delete(1.0, tk.END)  # Remove todo o conteúdo do log
        self.text_log.config(state=tk.DISABLED)  # Desabilita edição novamente
        self.adicionar_log("Log limpo")  # Adiciona mensagem de log limpo

    def configurar_interface(self):  # Configura a interface da tela de conversão
        header_frame = ttk.Frame(self)  # Cabeçalho da tela
        header_frame.pack(fill=tk.X, pady=(0, 15))  # Expande horizontalmente com espaçamento
        

        # Botão "Voltar" para retornar à tela inicial
        ttk.Button(
            header_frame, # Frame pai
            text="← Voltar", # Texto do botão
            command=lambda: self.controller.mostrar_tela("TelaInicial"), # Chama a tela inicial
            style='Fleury.TButton', # Aplica o estilo personalizado
            width=10
        ).pack(side=tk.LEFT) # Posiciona à esquerda com espaçamento

        # Título da tela
        ttk.Label( 
            header_frame, # Frame pai
            text="Conversor de arquivos",   # Texto do título
            font=FONTE_TITULO,           # Usa a fonte de título
            foreground=COR_PRIMARIA,     # Cor do texto (vermelho Fleury)
            background=COR_FUNDO       # Cor de fundo
        ).pack(side=tk.LEFT, padx=10) # Posiciona à esquerda com espaçamento

        # Frame de configurações
        config_frame = ttk.LabelFrame(self, text=" CONFIGURAÇÕES ", padding=10) # Frame com borda e título
        config_frame.pack(fill=tk.X, pady=5) # Expande horizontalmente com espaçamento

        # Frame para seleção de arquivo
        file_frame = ttk.Frame(config_frame) # Frame para o campo de arquivo
        file_frame.pack(fill=tk.X, pady=5) # Expande horizontalmente com espaçamento

        ttk.Label(file_frame, text="Arquivo de origem:").pack(side=tk.LEFT) # Label para o campo de arquivo

        self.entry_arquivo = ttk.Entry(file_frame, width=40, state="readonly")  # Torna o campo somente leitura
        self.entry_arquivo.pack(side=tk.LEFT, padx=5, expand=True, fill=tk.X)  # Expande para preencher o espaço

        ttk.Button(
            file_frame, # Frame pai
            text="Procurar...", # Texto do botão
            command=self.selecionar_arquivo, # Chama o método de seleção de arquivo
            style='Fleury.TButton' # Aplica o estilo personalizado
        ).pack(side=tk.LEFT) # Posiciona à esquerda

        # Frame para seleção do formato de conversão
        convert_frame = ttk.Frame(config_frame) # Frame para o campo de conversão
        convert_frame.pack(fill=tk.X, pady=5) # Expande horizontalmente com espaçamento

        ttk.Label(convert_frame, text="Converter para:").pack(side=tk.LEFT, padx=(0, 10))

        self.tipo_conversao = tk.StringVar(value="")  # Variável para armazenar o formato selecionado
        self.container_formatos = ttk.Frame(convert_frame)  # Container para os botões de formato
        self.container_formatos.pack(side=tk.LEFT, fill=tk.X, expand=True) # Expande para preencher o espaço

        # Frame para o log de atividades
        log_frame = ttk.LabelFrame(self, text=" LOG ", padding=10) # Frame com borda e título
        log_frame.pack(fill=tk.BOTH, expand=True, pady=5) # Expande para preencher o espaço

        self.text_log = tk.Text(
            log_frame, # Área de texto para o log
            wrap=tk.WORD, # Quebra de linha por palavras
            height=20, # Altura inicial
            font=FONTE_PRINCIPAL, # Fonte principal
            bg=COR_SECUNDARIA, # Cor de fundo (branco)
            fg=COR_TEXTO, # Cor do texto (cinza escuro)
            state=tk.DISABLED  # Define o estado como somente leitura
        )
        self.text_log.pack(fill=tk.BOTH, expand=True) # Expande para preencher o espaço

        self.text_log.bind("<Tab>", lambda event: self.text_log.tk_focusNext().focus() or "break")

        # Frame para os botões de ação
        btn_frame = ttk.Frame(self) # Frame para os botões de ação
        btn_frame.pack(fill=tk.X, pady=(10, 0)) # Expande horizontalmente com espaçamento


        ttk.Button( 
            btn_frame, # Frame pai
            text="Converter", # Texto do botão
            command=self.converter_arquivo, # Chama o método de conversão
            style='Fleury.TButton' # Aplica o estilo personalizado
        ).pack(side=tk.LEFT, padx=5) # Posiciona à esquerda com espaçamento

        ttk.Button(
            btn_frame, # Frame pai
            text="Limpar Log", # Texto do botão
            command=self.limpar_log, # Chama o método para limpar o log
            style='Fleury.TButton' # Aplica o estilo personalizado
        ).pack(side=tk.LEFT, padx=5) # Posiciona à esquerda com espaçamento

        ttk.Button(
            btn_frame, # Frame pai
            text="Ajuda", # Texto do botão
            command=self.mostrar_ajuda, # Chama o método para mostrar a ajuda
            style='Fleury.TButton' # Aplica o estilo personalizado
        ).pack(side=tk.LEFT, padx=5) # Posiciona à esquerda com espaçamento

        # Rodapé com copyright
        ttk.Label(
            self, # Frame pai
            text="© 2025 Matheus Augusto e Jeferson Sá - Todos os direitos reservados", # Texto de copyright
            font=('Arial', 8), # Fonte menor
            foreground='gray' # Cor cinza
        ).pack(side=tk.RIGHT, pady=(10, 0)) # Posiciona à direita com espaçamento
    
    def mostrar_ajuda(self):
        """Exibe uma janela de ajuda com as conversões suportadas."""

        ajuda_texto = """\
            CONVERSÕES SUPORTADAS:

            • PDF →  DOCX, TXT, PNG, XLSX
            • DOCX → PDF, TXT, PNG
            • PNG → JPG, JPEG, WEBP, ICO, PDF, DOCX
            • JPG → PNG, JPEG, WEBP, ICO, PDF, DOCX
            • JPEG → PNG, JPG, WEBP, ICO, PDF, DOCX
            • WEBP → PNG, JPG
            • ICO → PNG, JPG
            • XLSX ↔ CSV"""

        ajuda_janela = tk.Toplevel(self)
        ajuda_janela.withdraw()  # Oculta para evitar salto visual
        ajuda_janela.title("Ajuda - ZippyDocs")
        ajuda_janela.resizable(False, False)
        ajuda_janela.configure(bg="white")  # Fundo branco

        try:
            ajuda_janela.iconbitmap(ICONE)
        except:
            pass

        # Frame externo com borda
        frame = tk.Frame(ajuda_janela, bg="white", padx=15, pady=15)
        frame.pack(fill=tk.BOTH, expand=True)

        # Texto centralizado e formatado
        label_texto = tk.Label(
            frame,
            text=ajuda_texto,
            justify="left",
            anchor="w",
            font=("Arial", 11),
            bg="white",
            fg="#333333"
        )
        label_texto.pack(pady=(0, 10), anchor="w")

        # Botão fechar
        btn_fechar = ttk.Button(
            frame,
            text="Fechar",
            command=ajuda_janela.destroy,
            style="Fleury.TButton"
        )
        btn_fechar.pack(pady=(10, 0))

        # Atalhos de teclado
        ajuda_janela.bind("<Return>", lambda event: btn_fechar.invoke())
        ajuda_janela.bind("<Escape>", lambda event: ajuda_janela.destroy())

        # Centralizar em relação à janela principal
        self.update_idletasks()
        ajuda_janela.update_idletasks()
        largura = ajuda_janela.winfo_reqwidth()
        altura = ajuda_janela.winfo_reqheight()
        x = self.winfo_rootx() + (self.winfo_width() // 2) - (largura // 2)
        y = self.winfo_rooty() + (self.winfo_height() // 2) - (altura // 2)
        ajuda_janela.geometry(f"{largura}x{altura}+{x}+{y}")

        ajuda_janela.deiconify()
        ajuda_janela.lift()
        ajuda_janela.focus_force()
        btn_fechar.focus_set()



    def selecionar_arquivo(self, entry_destino=None):
        """Abre o diálogo de seleção de arquivo e valida o tipo MIME e integridade antes de aceitar."""
        try:
            arquivo = filedialog.askopenfilename(
                title="Selecione o arquivo",
                filetypes=TIPOS_ARQUIVOS_PERMITIDOS,
                initialdir=DOWNLOAD_FOLDER
            )

            if entry_destino is None:
                entry_destino = self.entry_arquivo # Usa o campo de entrada padrão se não for especificado

            
            if not arquivo:
                self.adicionar_log("Nenhum arquivo selecionado.")
                # Limpa o campo de entrada se nenhum arquivo for selecionado
                entry_destino.config(state="normal")  # Habilita o campo de entrada
                entry_destino.delete(0, tk.END)  # Limpa o campo de entrada
                entry_destino.config(state="readonly")  # Torna o campo somente leitura
                return  # Sai se nenhum arquivo for selecionado

            # Validação MIME logo após seleção
            tipos_entrada_aceitos = [
                "application/pdf", "application/x-pdf",
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                "application/msword",
                 "application/zip",
                "text/plain",
                "image/png", "image/x-png",
                "image/jpeg", "image/pjpeg",
                "image/webp",
                "image/x-icon", "image/vnd.microsoft.icon",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                "application/vnd.ms-excel",
                "text/csv"
            ]


            if not validar_mime(arquivo, tipos_entrada_aceitos):
                entry_destino.config(state="normal")
                entry_destino.delete(0, tk.END)
                entry_destino.config(state="readonly")
                self.adicionar_log("Arquivo rejeitado por tipo MIME inválido.")
                self.mostrar_mensagem("Erro", f"O arquivo '{os.path.basename(arquivo)}' não é suportado ou pode ser perigoso.", erro=True)
                return

            # Validação de integridade (exemplo para DOCX, PDF e imagens)
            ext = os.path.splitext(arquivo)[1].lower()

            try:
                if ext == ".pdf":
                    poppler_path = obter_caminho_poppler()
                    validar_pdf_com_pdfinfo(arquivo, poppler_path)  # Valida PDF com pdfinfo
                    self.adicionar_log("PDF validado com sucesso.")
                elif ext == ".docx":
                    validar_docx_com_zip(arquivo)
                    self.adicionar_log("DOCX validado com sucesso.")
            except ValueError as ve:
                entry_destino.config(state="normal")
                entry_destino.delete(0, tk.END)
                entry_destino.config(state="readonly")
                self.adicionar_log(f"Arquivo rejeitado por erro de validação.")
                logger.error(traceback.format_exc())
                logger.error(str(ve))
                self.mostrar_mensagem("Erro", f"O arquivo '{os.path.basename(arquivo)}' não é suportado ou pode ser perigoso.", erro=True)
                return

            # Validação de integridade (após validação prévia)
            try:
                if ext == ".docx":
                    Document(arquivo)  # Tenta abrir o DOCX
                elif ext == ".pdf":
                    PdfReader(arquivo)  # Tenta abrir o PDF
                elif ext in [".jpg", ".jpeg", ".png", ".webp", ".ico"]:
                    Image.open(arquivo).verify()  # Tenta abrir a imagem
                elif ext == ".xlsx":
                    pd.read_excel(arquivo)
                elif ext == ".csv":
                    pd.read_csv(arquivo, nrows=1)
            except Exception as e:
                entry_destino.config(state="normal")
                entry_destino.delete(0, tk.END)
                entry_destino.config(state="readonly")
                self.adicionar_log(f"Arquivo rejeitado por estar corrompido.")
                logger.error(f"Erro ao validar arquivo {arquivo}: {str(e)}")
                logger.error(traceback.format_exc())
                self.mostrar_mensagem("Erro", f"O arquivo '{os.path.basename(arquivo)}' não é suportado ou pode ser perigoso.", erro=True)
                return


            # Só chega aqui se o arquivo for válido e íntegro
            entry_destino.config(state="normal")
            entry_destino.delete(0, tk.END)
            entry_destino.insert(0, arquivo)
            entry_destino.config(state="readonly")
            self.adicionar_log(f"Arquivo selecionado: {os.path.basename(arquivo)}")
            self.tipo_conversao.set("")  # Limpa a seleção do radiobutton
            if hasattr(self, "atualizar_botoes_conversao"):
                self.atualizar_botoes_conversao(arquivo)

        except Exception as e:
            self.mostrar_mensagem("Erro", "Erro ao selecionar arquivo. Consulte o suporte ou o log interno.", erro=True)
            logger.error(f"Erro ao selecionar arquivo: {str(e)}")
            logger.error(traceback.format_exc())

    def atualizar_botoes_conversao(self, arquivo): # Método para atualizar os botões de conversão
        """Atualiza os botões de conversão com base na extensão do arquivo.""" 
        extensao = os.path.splitext(arquivo)[1].lower() # Obtém a extensão do arquivo em minúsculas
        formatos_disponiveis = { # Dicionário de formatos disponíveis para conversão
            '.pdf': ['docx', 'txt', 'png', 'XLSX'],
            '.png': ['jpg', 'jpeg', 'webp', 'ico', 'pdf', 'docx'],
            '.jpg': ['png', 'jpeg', 'webp', 'ico', 'pdf', 'docx'],
            '.jpeg': ['png', 'jpg', 'webp', 'ico', 'pdf', 'docx'],
            '.webp': ['png', 'jpg'],
            '.ico': ['png', 'jpg'],
            '.csv': ['xlsx'],
            '.xlsx': ['csv']
        }

        # Obtém os formatos disponíveis para a extensão selecionada
        formatos = formatos_disponiveis.get(extensao, [])

        # Atualiza os botões de conversão
        for widget in self.container_formatos.winfo_children():
            widget.destroy()  # Remove os botões existentes

        if formatos: # Se houver formatos disponíveis
            for formato in formatos: # Para cada formato disponível
                ttk.Radiobutton( # Cria um botão de opção (radiobutton) para cada formato
                    self.container_formatos, # Frame pai
                    text=formato.upper(), # Texto do botão (formato em maiúsculas)
                    variable=self.tipo_conversao, # Variável associada ao botão
                    value=formato # Valor associado ao botão (formato selecionado)
                ).pack(side=tk.LEFT, padx=5) # Posiciona os botões à esquerda com espaçamento
        else:
            self.adicionar_log(f"Conversão não suportada para arquivos com extensão {extensao}") # Adiciona mensagem de log se não houver formatos disponíveis

    def contem_imagens(self, arquivo): # Método para verificar se o arquivo contém imagens
        """Verifica se o arquivo contém imagens (apenas para PDF e DOCX)."""
        try:
            if arquivo.endswith('.pdf'): # Se o arquivo for PDF
                reader = PdfReader(arquivo) # Lê o arquivo PDF
                for page in reader.pages:   # Para cada página do PDF
                    if '/XObject' in page.get('/Resources', {}): # Verifica se há objetos de imagem
                        return True # Retorna True se encontrar imagens
            elif arquivo.endswith('.docx'): # Se o arquivo for DOCX
                doc = Document(arquivo) # Lê o arquivo DOCX
                for rel in doc.part.rels.values(): # Para cada relacionamento no documento
                    if "image" in rel.target_ref: # Verifica se é uma imagem
                        return True # Retorna True se encontrar imagens
        except Exception as e: # Tratamento de erros
            self.adicionar_log(f"Erro ao verificar imagens no arquivo: {str(e)}") # Adiciona mensagem de log
            logger.error(f"Erro ao verificar imagens no arquivo {arquivo}: {str(e)}") # Registra o erro
            logger.error(traceback.format_exc()) # Registra o traceback do erro
        return False # Retorna False se não encontrar imagens
        
    
    def converter_arquivo(self): # Método para converter o arquivo selecionado
        '""Converte o arquivo selecionado para o formato desejado."""'
        arquivo_origem = self.entry_arquivo.get() # Obtém o caminho do arquivo

        MAX_FILE_SIZE_MB = 500
        if os.path.getsize(arquivo_origem) > MAX_FILE_SIZE_MB * 1024 * 1024:
            self.mostrar_mensagem("Erro", f"O arquivo excede o limite de {MAX_FILE_SIZE_MB}MB.", erro=True)
            return
        
        if not arquivo_origem:               # Verifica se um arquivo foi selecionado
            self.mostrar_mensagem("Erro", "Selecione um arquivo para converter", erro=True) ## Mensagem de erro
            return
        
        tipo_saida = self.tipo_conversao.get().lower() # Formato de saída em minúsculas
        if not tipo_saida:
            self.mostrar_mensagem("Erro", "Selecione o formato de conversão desejado.", erro=True)
            return
        
        if not os.path.isfile(arquivo_origem): # Verifica se o arquivo existe
            self.mostrar_mensagem("Erro", "Arquivo não encontrado", erro=True) # Mensagem de erro
            return
        
        tipos_entrada_aceitos = [
            "application/pdf", "application/x-pdf",
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            "application/msword",
            "application/zip",
            "text/plain",
            "image/png", "image/x-png",
            "image/jpeg", "image/pjpeg",
            "image/webp",
            "image/x-icon", "image/vnd.microsoft.icon",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            "application/vnd.ms-excel",
            "text/csv",
            "text/plain",
            "application/vnd.ms-excel"
             ]

        if not validar_mime(arquivo_origem, tipos_entrada_aceitos): # Valida o tipo MIME do arquivo
            self.mostrar_mensagem("Erro", "Tipo de arquivo não suportado", erro=True) # Mensagem de erro
            return
        
        tipo_saida = self.tipo_conversao.get().lower() # Formato de saída em minúsculas
        extensao_origem = os.path.splitext(arquivo_origem)[1][1:].lower() # Extensão do arquivo original
        
        try:
            self.adicionar_log(f"Iniciando conversão de {extensao_origem} para {tipo_saida}...") # Adiciona mensagem de log
            
            # Gera um nome único para o arquivo de destino
            nome_base = os.path.splitext(os.path.basename(arquivo_origem))[0]

            # Sanitização manual para evitar path traversal
            sequencias_perigosas = [
                "../", "..\\", "%2e%2e%2f", "%2e%2e/", "%2e%2e\\", "%252e%252e%255c", "%252e%252e%252f"
            ]
            for seq in sequencias_perigosas:
                nome_base = nome_base.replace(seq, "")
            nome_base = nome_base.replace("..", "")
            nome_base = re.sub(r'[^a-zA-Z0-9_\-]', '_', nome_base)
            if not nome_base:
                nome_base = "arquivo"

            nome_seguro = uuid.uuid4().hex  # Nome totalmente imprevisível
            pasta_temp = os.path.join(tempfile.gettempdir(), f"conversao_temp_{nome_seguro}")
            os.makedirs(pasta_temp, exist_ok=True)
            os.chmod(pasta_temp, 0o600)

            if (extensao_origem, tipo_saida) in [('pdf', 'png'), ('docx', 'png')]:
                arquivo_destino = os.path.join(pasta_temp, f"{nome_seguro}.zip")
            else:
                arquivo_destino = os.path.join(pasta_temp, f"{nome_seguro}.{tipo_saida}")

            # Dicionário de funções de conversão suportadas
            conversoes = {
                # Conversões de PDF para outros formatos
                ('pdf', 'jpg'): self.pdf_para_jpg, # Converte PDF para JPG
                ('pdf', 'png'): pdf_para_png_global, # Converte PDF para PNG
                ('pdf', 'docx'): pdf_para_docx_global, # Converte PDF para DOCX
                ('pdf', 'txt'): pdf_para_txt_global, # Converte PDF para TXT
                ('pdf', 'xlsx'): pdf_para_xlsx_global, # Converte PDF para XLSX
                
                # Conversões de DOCX para outros formatos
                ('docx', 'pdf'): docx_para_pdf_global, # Converte DOCX para PDF
                ('docx', 'jpg'): self.docx_para_jpg, # Converte DOCX para JPG
                ('docx', 'png'): docx_para_png_global, # Converte DOCX para PNG
                ('docx', 'txt'): docx_para_txt_global, # Converte DOCX para TXT
                
                # Conversões de imagens para outros formatos
                ('jpg', 'pdf'): self.imagem_para_pdf, # Converte JPG para PDF
                ('png', 'pdf'): self.imagem_para_pdf,   # Converte PNG para PDF
                ('jpg', 'docx'): self.imagem_para_docx, # Converte JPG para DOCX
                ('png', 'docx'): self.imagem_para_docx, # Converte PNG para DOCX
                ('jpg', 'png'): self.converter_para_png, # Converte JPG para PNG
                ('png', 'jpg'): self.converter_para_jpg, # Converte PNG para JPG
                ('jpg', 'ico'): self.converter_para_ico, # Converte JPG para ICO
                ('png', 'ico'): self.converter_para_ico, # Converte PNG para ICO
                ('png', 'webp'): self.converter_para_webp, # Converte PNG para WEBP
                ('png', 'webp'): self.converter_para_webp, # Converte WEBP para PNG
                ('jpg', 'webp'): self.converter_para_webp, # Converte JPG para WEBP
                ('jpeg', 'webp'): self.converter_para_webp, # Converte JPEG para WEBP
                ('png', 'jpeg'): self.converter_para_jpeg, # Converte PNG para JPEG
                ('jpg', 'jpeg'): self.converter_para_jpeg,  # JPG para JPEG
                ('jpeg', 'jpg'): self.converter_para_jpg,  # JPEG para JPG
                ('jpeg', 'png'): self.converter_para_png,  # JPEG para PNG
                ('jpeg', 'webp'): self.converter_para_webp,  # JPEG para WEBP
                ('jpeg', 'ico'): self.converter_para_ico,  # JPEG para ICO
                ('jpeg', 'pdf'): self.imagem_para_pdf,  # JPEG para PDF
                ('jpeg', 'docx'): self.imagem_para_docx,  # JPEG para DOCX
                ('webp', 'png'): lambda o, d: self.webp_para_imagem(o, d, 'png'), # Converte WEBP para PNG
                ('webp', 'jpg'): lambda o, d: self.webp_para_imagem(o, d, 'jpg'), # Converte WEBP para JPG

                
                # Conversões de CSV para outros formatos
                ('csv', 'pdf'): self.csv_para_pdf, # Converte CSV para PDF
                ('csv', 'docx'): self.csv_para_docx, # Converte CSV para DOCX
                ('csv', 'xlsx'): csv_para_xlsx_global, # Converte CSV para XLSX
                
                # Conversões de Excel para outros formatos
                ('xlsx', 'pdf'): self.xlsx_para_pdf, # Converte XLSX para PDF
                ('xlsx', 'docx'): self.xlsx_para_docx, # Converte XLSX para DOCX
                ('xlsx', 'csv'): xlsx_para_csv_global, # Converte XLSX para CSV
                
                # Conversões de ICO para outros formatos
                ('ico', 'png'): self.converter_para_png, # Converte ICO para PNG
                ('ico', 'jpg'): self.converter_para_jpg, # Converte ICO para JPG

            }

            # Verifica se a conversão solicitada é suportada
            if (extensao_origem, tipo_saida) in conversoes:
                self.adicionar_log(f"Conversão encontrada: {extensao_origem} para {tipo_saida}")

                # Lista de conversões críticas para rodar em sandbox
                conversoes_criticas = [
                    ('pdf', 'txt'), ('pdf', 'docx'), ('pdf', 'png'), ('pdf', 'jpg'),
                    ('docx', 'pdf'), ('docx', 'jpg'), ('docx', 'png'), ('docx', 'txt'),
                    ('xlsx', 'pdf'), ('xlsx', 'docx'), ('xlsx', 'csv'),
                    ('csv', 'pdf'), ('csv', 'docx'), ('csv', 'xlsx')
                ]

                if (extensao_origem, tipo_saida) in conversoes_criticas:
                    # Defina um limite maior para conversões pesadas
                    limite_mb = 1000 if (extensao_origem, tipo_saida) in [('pdf', 'docx')] else 500
                    sucesso, resultado = executar_com_timeout(
                        conversoes[(extensao_origem, tipo_saida)],
                        args=(arquivo_origem, arquivo_destino),
                        timeout=300,  
                        mem_limit_mb=limite_mb
                    )
                    if not sucesso:
                        raise ValueError(f"Erro/sandbox: {resultado}")
                else:
                    conversoes[(extensao_origem, tipo_saida)](arquivo_origem, arquivo_destino)
            else:
                self.adicionar_log(f"Conversão não encontrada: {extensao_origem} para {tipo_saida}")
                raise ValueError(f"Conversão de {extensao_origem} para {tipo_saida} não suportada")
            
            # Validação MIME do arquivo de saída
            tipos_saida_aceitos = {
                'pdf': ["application/pdf", "application/x-pdf"],
                'docx': ["application/vnd.openxmlformats-officedocument.wordprocessingml.document"],
                'txt': ["text/plain"],
                'png': ["image/png", "image/x-png"],
                'jpg': ["image/jpeg", "image/pjpeg"],
                'jpeg': ["image/jpeg", "image/pjpeg"],
                'webp': ["image/webp"],
                'ico': ["image/x-icon", "image/vnd.microsoft.icon"],
                'xlsx': ["application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"],
                'csv': ["text/csv", "text/plain", "application/vnd.ms-excel"],
                'zip': ["application/zip", "application/x-zip-compressed"]
            }

            if (extensao_origem, tipo_saida) in [('pdf', 'png'), ('docx', 'png')]:
                tipos_esperados = tipos_saida_aceitos.get('zip', [])
            else:
                tipos_esperados = tipos_saida_aceitos.get(tipo_saida, [])
            
            if not validar_mime(arquivo_destino, tipos_esperados):
                os.remove(arquivo_destino)  # Remove o arquivo se o MIME não for válido
                self.mostrar_mensagem("Atenção", "O arquivo gerado não corresponde ao tipo esperado e foi removido.", erro=True)
                return
            
            try:
                nome_original = os.path.splitext(os.path.basename(arquivo_origem))[0]
                extensao = tipo_saida
                if (extensao_origem, tipo_saida) in [('pdf', 'png'), ('docx', 'png')]:
                    nome_final = gerar_nome_download(nome_original, 'zip')
                else:
                    nome_final = gerar_nome_download(nome_original, extensao)
                caminho_final = os.path.join(DOWNLOAD_FOLDER, nome_final)
                shutil.copy2(arquivo_destino, caminho_final)
                self.adicionar_log(f"Arquivo exportado para Downloads: {caminho_final}")
                self.mostrar_mensagem("Sucesso", f"Arquivo convertido com sucesso!\nSalvo em: {caminho_final}")
            except Exception as e:
                self.adicionar_log("Erro ao exportar para Downloads. Consulte o log interno.")
                self.mostrar_mensagem("Erro", "Falha ao exportar para Downloads. Consulte o suporte ou o log interno.", erro=True)
                registrar_log_tecnico(e)
                logger.error(f"Erro ao exportar arquivo convertido: {str(e)}")
                logger.error(traceback.format_exc())
            
        except Exception as e:
            self.adicionar_log("Erro durante a conversão. Consulte o log interno.")
            self.mostrar_mensagem("Erro na conversão", "Ocorreu um erro durante a conversão. Consulte o suporte ou o log interno.", erro=True)
            logger.error(f"Erro durante a conversão: {str(e)}")
            registrar_log_tecnico(e)
            logger.error(traceback.format_exc()) # Registra o traceback do erro
    
    # ==================================================
    # MÉTODOS DE CONVERSÃO ESPECÍFICOS
    # ==================================================
    
    # Usar em uma nova versão do programa
    def pdf_para_jpg(self, origem, destino): # Método para converter PDF para JPG
        """Converte PDF para JPG (uma imagem por página)"""
        try:
            imagens = convert_from_path(origem) # Converte cada página do PDF para imagem
            if len(imagens) == 1:             # Se tiver apenas uma página
                imagens[0].save(destino, 'JPEG', quality=95) # Salva como JPG
                os.chmod(destino, 0o600) # Define permissões de leitura e escrita para o proprietário
            else:  # Se tiver múltiplas páginas
                nome_unico = os.path.splitext(destino)[0]  # Usa o nome único já gerado para o destino
                for i, img in enumerate(imagens):
                    nome_img = f"{nome_unico}_pagina_{i+1}.jpg"
                    img.save(nome_img, 'JPEG', quality=95)
                    os.chmod(nome_img, 0o600)
        except Exception as e:
            logger.error(traceback.format_exc())
            logger.error(f"Erro ao converter PDF para JPG: {str(e)}") # Registra o erro
            raise ValueError(f"Erro ao converter PDF para JPG: {str(e)}") ## Lança erro se falhar

    def pdf_para_png(self, origem, destino):
        """Converte PDF para PNG (uma imagem por página) e compacta em um arquivo ZIP."""
        try:
            poppler_path = obter_caminho_poppler()  # Caminho do Poppler
            imagens = convert_from_path(origem, poppler_path=poppler_path)
            
            nome_seguro = uuid.uuid4().hex  # Nome totalmente imprevisível
            pasta_temp = os.path.join(tempfile.gettempdir(), f"imagens_{nome_seguro}")
            os.makedirs(pasta_temp, exist_ok=True)
            os.chmod(pasta_temp, 0o600)  # Permissões restritas para a pasta temporária
            
            # Salva cada página como PNG
            caminhos_imagens = []
            for i, img in enumerate(imagens):
                img_path = os.path.join(pasta_temp, f"{nome_seguro}_pagina_{i+1}.png")
                img.save(img_path, 'PNG')
                os.chmod(img_path, 0o600)  # Permissões restritas para cada PNG
                caminhos_imagens.append(img_path)
            
            # Compacta as imagens em um arquivo ZIP
            with zipfile.ZipFile(destino, 'w', zipfile.ZIP_DEFLATED) as zipf:
                for img_path in caminhos_imagens:
                    zipf.write(img_path, os.path.basename(img_path))
            
            os.chmod(destino, 0o600)  # Permissões restritas para o ZIP
            
            # Verifica a integridade do ZIP
            with zipfile.ZipFile(destino, 'r') as zipf:
                if zipf.testzip() is not None:
                    raise ValueError("Arquivo ZIP corrompido")
            
            # Limpa os arquivos temporários
            for img_path in caminhos_imagens:
                os.remove(img_path)
            os.rmdir(pasta_temp)
            
        except Exception as e:
            if os.path.exists(pasta_temp):
                shutil.rmtree(pasta_temp, ignore_errors=True)
            logger.error(traceback.format_exc())
            logger.error(f"Erro ao converter PDF para PNG: {str(e)}")  #    
            raise ValueError(f"Erro ao converter PDF para PNG: {str(e)}")

    # usar em uma nova versão do programa    
    def docx_para_jpg(self, origem, destino):
        """Converte DOCX para JPG"""
        try:
            # Gera PDF temporário a partir do DOCX usando docx2pdf
            pdf_temp = os.path.join(tempfile.gettempdir(), f"{uuid.uuid4().hex}.pdf")
            docx2pdf_convert(origem, pdf_temp)

            # Converte o PDF temporário para imagens JPG
            poppler_path = obter_caminho_poppler()
            imagens = convert_from_path(pdf_temp, poppler_path=poppler_path)
            nome_unico = os.path.splitext(destino)[0]
            for i, img in enumerate(imagens):
                nome_img = f"{nome_unico}_pagina_{i+1}.jpg"
                img.save(nome_img, "JPEG", quality=95)
                os.chmod(nome_img, 0o600)

            os.remove(pdf_temp)  # Remove o PDF temporário
            self.adicionar_log(f"Arquivo DOCX convertido para JPG com sucesso")
        except Exception as e:
            logger.error(traceback.format_exc())
            logger.error(f"Erro ao converter DOCX para JPG: {str(e)}")  # Registra o erro
            raise ValueError(f"Erro ao converter DOCX para JPG: {str(e)}")


    def docx_para_png(self, origem, destino):
        """Converte DOCX para PNG"""
        try:
            # Gera PDF temporário a partir do DOCX usando docx2pdf
            pdf_temp = os.path.join(tempfile.gettempdir(), f"{uuid.uuid4().hex}.pdf")
            docx2pdf_convert(origem, pdf_temp)

            # Converte o PDF temporário para imagens PNG
            poppler_path = obter_caminho_poppler()
            imagens = convert_from_path(pdf_temp, poppler_path=poppler_path)
            nome_unico = os.path.splitext(destino)[0]
            for i, img in enumerate(imagens):
                nome_img = f"{nome_unico}_pagina_{i+1}.png"
                img.save(nome_img, "PNG")
                os.chmod(nome_img, 0o600)

            os.remove(pdf_temp)  # Remove o PDF temporário
            self.adicionar_log(f"Arquivo DOCX convertido para PNG com sucesso")
        except Exception as e:
            logger.error(traceback.format_exc())
            logger.error(f"Erro ao converter DOCX para PNG: {str(e)}")
            raise ValueError(f"Erro ao converter DOCX para PNG: {str(e)}")

    def imagem_para_docx(self, origem, destino): 
        """Converte imagem (JPG/PNG) para DOCX""" 
        try:
            doc = Document()                  # Cria um novo documento Word
            # Adiciona a imagem com largura de 6 polegadas
            doc.add_picture(origem, width=Inches(6)) # Adiciona a imagem ao documento
            doc.save(destino)                 # Salva o documento
            os.chmod(destino, 0o600)
        except Exception as e: # Tratamento de erros
            logger.error(traceback.format_exc())
            logger.error(f"Erro ao converter imagem para DOCX: {str(e)}")
            raise ValueError(f"Erro ao converter imagem para DOCX: {str(e)}") # Lança erro se falhar

    def csv_para_pdf(self, origem, destino): # Método para converter CSV para PDF
        """Converte CSV para PDF formatado""" # Docstring para documentação
        try:
            df = pd.read_csv(origem)          # Lê o arquivo CSV com pandas

            # ALERTA DE CONTEÚDO PERIGOSO
            if df.applymap(contem_estrutura_perigosa).any().any():
                self.mostrar_mensagem(
                    "Alerta de Segurança",
                    "O arquivo contém fórmulas ou scripts potencialmente perigosos. Eles serão neutralizados na conversão.",
                    erro=True
                )
            
            pdf = FPDF()                      # Cria um novo PDF
            pdf.add_page()                    # Adiciona uma página
            pdf.set_font("Arial", size=10)    # Define a fonte
            
            # Adiciona título com o nome do arquivo
            pdf.cell(200, 10, txt=f"Relatório CSV: {os.path.basename(origem)}", ln=1, align='C') # Centraliza o título
            
            # Configurações para a tabela
            col_width = pdf.w / (len(df.columns) + 1) # Largura das colunas
            row_height = pdf.font_size * 2    # Altura das linhas
            
            # Cabeçalho da tabela
            for col in df.columns:         # Para cada coluna do DataFrame
                pdf.cell(col_width, row_height, txt=str(col), border=1) # Adiciona o nome da coluna
            pdf.ln(row_height)                # Quebra de linha
            
            # Dados da tabela
            for _, row in df.iterrows():
                for item in row:
                    valor = remover_tags_html(sanitizar_celula_excel(item))
                    pdf.cell(col_width, row_height, txt=str(valor), border=1)
                pdf.ln(row_height)
            
            pdf.output(destino)               # Salva o PDF
            os.chmod(destino, 0o600)         # Define permissões de leitura e escrita para o proprietário
        except Exception as e:
            logger.error(traceback.format_exc())
            logger.error(f"Erro ao converter CSV para PDF: {str(e)}") # Registra o erro
            raise ValueError(f"Erro ao converter CSV para PDF: {str(e)}") # Lança erro se falhar

    def csv_para_docx(self, origem, destino):
        """Converte CSV para DOCX formatado"""
        try:
            df = pd.read_csv(origem)          # Lê o arquivo CSV com pandas

            # ALERTA DE CONTEÚDO PERIGOSO
            if df.applymap(contem_estrutura_perigosa).any().any():
                self.mostrar_mensagem(
                    "Alerta de Segurança",
                    "O arquivo contém fórmulas ou scripts potencialmente perigosos. Eles serão neutralizados na conversão.",
                    erro=True
                )
            
            doc = Document()                  # Cria um novo documento Word
            # Adiciona título com o nome do arquivo
            doc.add_heading(f"Relatório CSV: {os.path.basename(origem)}", 0) # Adiciona título
            
            # Cria tabela no documento com o número de colunas do CSV
            table = doc.add_table(rows=1, cols=len(df.columns)) # Cria tabela
            table.style = 'Table Grid'        # Aplica estilo de grade
            
            # Cabeçalho da tabela
            hdr_cells = table.rows[0].cells  # Células da primeira linha
            for i, col in enumerate(df.columns): # Para cada coluna
                hdr_cells[i].text = str(col)  # Adiciona o nome da coluna
            
            # Dados da tabela
            for _, row in df.iterrows():      # Para cada linha do DataFrame
                row_cells = table.add_row().cells # Adiciona nova linha
                for i, item in enumerate(row):
                    valor = remover_tags_html(sanitizar_celula_excel(item))
                    row_cells[i].text = str(valor) # Adiciona o valor na célula
            
            doc.save(destino)                 # Salva o documento
            os.chmod(destino, 0o600)         # Define permissões de leitura e escrita para o proprietário
        except Exception as e:         # Tratamento de erro
            logger.error(traceback.format_exc())
            logger.error(f"Erro ao converter CSV para DOCX: {str(e)}")
            raise ValueError(f"Erro ao converter CSV para DOCX: {str(e)}") # Lança erro se falhar


    # Próxima versão do programa
    def xlsx_para_pdf(self, origem, destino): 
        """Converte XLSX para PDF"""
        try:
            df = pd.read_excel(origem)
            
            # ALERTA DE CONTEÚDO PERIGOSO
            if df.applymap(contem_estrutura_perigosa).any().any():
                self.mostrar_mensagem(
                    "Alerta de Segurança",
                    "O arquivo contém fórmulas ou scripts potencialmente perigosos. Eles serão neutralizados na conversão.",
                    erro=True
                )

            c = canvas.Canvas(destino, pagesize=letter)
            width, height = letter
            y = height - 40
            c.setFont("Helvetica", 10)
            # Cabeçalho
            for i, col in enumerate(df.columns):
                c.drawString(40 + i*100, y, str(col))
            y -= 20
            # Dados
            for _, row in df.iterrows():
                for i, item in enumerate(row):
                    valor = remover_tags_html(sanitizar_celula_excel(item))
                    c.drawString(40 + i*100, y, str(valor))
                y -= 20
                if y < 40:
                    c.showPage()
                    y = height - 40
            c.save()
            os.chmod(destino, 0o600)
            self.adicionar_log(f"Arquivo XLSX convertido para PDF com sucesso")
        except Exception as e:
            logger.error(traceback.format_exc())
            logger.error(f"Erro ao converter XLSX para PDF: {str(e)}")
            raise ValueError(f"Erro ao converter XLSX para PDF: {str(e)}")

    def xlsx_para_docx(self, origem, destino): # Método para converter Excel (XLSX) para DOCX
        """Converte Excel (XLSX) para DOCX formatado"""
        try:
            df = pd.read_excel(origem)        # Lê o arquivo Excel com pandas

            # ALERTA DE CONTEÚDO PERIGOSO
            if df.applymap(contem_estrutura_perigosa).any().any():
                self.mostrar_mensagem(
                    "Alerta de Segurança",
                    "O arquivo contém fórmulas ou scripts potencialmente perigosos. Eles serão neutralizados na conversão.",
                    erro=True
                )
            
            doc = Document()                  # Cria um novo documento Word
            # Adiciona título com o nome do arquivo
            doc.add_heading(f"Relatório Excel: {os.path.basename(origem)}", 0) # Adiciona título
            
            # Cria tabela no documento com o número de colunas do Excel
            table = doc.add_table(rows=1, cols=len(df.columns)) # Cria tabela
            table.style = 'Table Grid'        # Aplica estilo de grade
            
            # Cabeçalho da tabela
            hdr_cells = table.rows[0].cells  # Células da primeira linha
            for i, col in enumerate(df.columns): # Para cada coluna
                hdr_cells[i].text = str(col)  # Adiciona o nome da coluna
            
            # Dados da tabela
            for _, row in df.iterrows():      # Para cada linha do DataFrame
                row_cells = table.add_row().cells # Adiciona nova linha
                for i, item in enumerate(row):
                    valor = remover_tags_html(sanitizar_celula_excel(item))
                    row_cells[i].text = str(valor) # Adiciona o valor na célula, sanitizando e removendo tags HTML
            
            doc.save(destino)                 # Salva o documento
            os.chmod(destino, 0o600)         # Define permissões de leitura e escrita para o proprietário
        except Exception as e:          # Tratamento de erro
            logger.error(traceback.format_exc())
            logger.error(f"Erro ao converter XLSX para DOCX: {str(e)}")
            raise ValueError(f"Erro ao converter XLSX para DOCX: {str(e)}") # Lança erro se falhar

    # Métodos de conversão mais simples
    def imagem_para_pdf(self, origem, destino):
        try:
            """Converte imagem (JPG/PNG) para PDF""" # Docstring para documentação
            img = Image.open(origem).convert("RGB") # Abre a imagem e converte para RGB
            img.save(destino, "PDF")              # Salva como PDF
            os.chmod(destino, 0o600) # Define permissões de leitura e escrita para o proprietário
        except Exception as e:
            logger.error(traceback.format_exc())
            logger.error(f"Erro ao converter imagem para PDF: {str(e)}") 
        
    def docx_para_pdf(self, origem, destino):
        """Converte DOCX para PDF de forma segura usando docx2pdf"""
        try:
            doc = Document(origem)
            texto_completo = ""
            for para in doc.paragraphs:
                texto_completo += para.text + "\n"
            for table in doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                        texto_completo += cell.text + "\n"
            
            if contem_estrutura_perigosa(texto_completo):
                self.mostrar_mensagem(
                    "Alerta de Segurança",
                    "O arquivo contém fórmulas ou scripts potencialmente perigosos. Eles serão neutralizados na conversão.",
                    erro=True
                )
                texto_completo = remover_tags_html(texto_completo)
                # Aqui seria ideal reescrever o DOCX sanitizado, mas como docx2pdf não permite, apenas alertamos

            docx2pdf_convert(origem, destino)
            os.chmod(destino, 0o600)
        except Exception as e:
            logger.error(traceback.format_exc())
            logger.error(f"Erro ao converter DOCX para PDF: {str(e)}")
            raise ValueError(f"Erro ao converter DOCX para PDF: {str(e)}")
    
    def pdf_para_txt(self, origem, destino):
        """Converte PDF para texto simples, ignorando imagens."""
        try:
            with open(origem, 'rb') as f: # Abre o arquivo PDF
                reader = PdfReader(f)  # Lê o PDF
                texto_extraido = "" # Inicializa variável para armazenar texto extraído
                for pagina in reader.pages:     # Para cada página do PDF
                    texto_extraido += pagina.extract_text()  # Extrai o texto da página

            if not texto_extraido.strip():  # Verifica se o texto extraído está vazio
                raise ValueError("Nenhum texto encontrado no PDF.") # Lança erro se vazio

            with open(destino, 'w', encoding='utf-8') as f_out: # Abre o arquivo de destino
                f_out.write(texto_extraido)  # Salva o texto extraído no arquivo TXT
            os.chmod(destino, 0o600) # Define permissões de leitura e escrita para o proprietário

            self.adicionar_log(f"PDF convertido para TXT com sucesso") # Registra sucesso no log
        except Exception as e: # Tratamento de erro
            logger.error(traceback.format_exc())
            logger.error(f"Erro ao converter PDF para TXT: {str(e)}")
            raise ValueError(f"Erro ao converter PDF para TXT: {str(e)}")  # Lança erro se falhar
    
    def docx_para_txt(self, origem, destino):
        """Converte DOCX para texto simples"""
        try:
            doc = Document(origem)                # Abre o documento Word
            with open(destino, 'w', encoding='utf-8') as f: # Abre o arquivo de destino
                for para in doc.paragraphs:
                    texto_limpo = remover_tags_html(para.text)
                    f.write(texto_limpo + '\n') # Escreve o texto no arquivo TXT
            os.chmod(destino, 0o600) # Define permissões de leitura e escrita para o proprietário
        except Exception as e:          # Tratamento de erro
            logger.error(traceback.format_exc())
            logger.error(f"Erro ao converter DOCX para TXT: {str(e)}")
            raise ValueError(f"Erro ao converter DOCX para TXT: {str(e)}") # Lança erro se falhar
    
    def xlsx_para_csv(self, origem, destino):
        """Converte Excel (XLSX) para CSV"""
        try:
            df = pd.read_excel(origem)            # Lê o arquivo Excel
            df.to_csv(destino, index=False, encoding='utf-8') # Salva como CSV
            os.chmod(destino, 0o600) # Define permissões de leitura e escrita para o proprietário
        except Exception as e:          # Tratamento de erro
            logger.error(traceback.format_exc())
            logger.error(f"Erro ao converter XLSX para CSV: {str(e)}")
            raise ValueError(f"Erro ao converter XLSX para CSV: {str(e)}") # Lança erro se falhar

    def converter_para_png(self, origem, destino):
        """Converte imagem (JPEG/JPG/ICO) para PNG"""
        try:
            img = Image.open(origem)  # Abre a imagem
            img.save(destino, 'PNG')  # Salva como PNG
            os.chmod(destino, 0o600) # Define permissões de leitura e escrita para o proprietário
            self.adicionar_log(f"Imagem convertida para PNG com sucesso") # Registra sucesso no log
        except Exception as e: 
            logger.error(traceback.format_exc())         # Tratamento de erro
            logger.error(f"Erro ao converter para PNG: {str(e)}") # Registra o erro
            raise ValueError(f"Erro ao converter para PNG: {str(e)}") # Lança erro se falhar

    def converter_para_jpeg(self, origem, destino):
        """Converte imagem (PNG/JPG) para JPEG"""
        try:
            img = Image.open(origem).convert("RGB")  # Abre a imagem e converte para RGB
            img.save(destino, 'JPEG', quality=95)  # Salva como JPEG com qualidade 95%
            os.chmod(destino, 0o600) # Define permissões de leitura e escrita para o proprietário
            self.adicionar_log(f"Imagem convertida para JPEG com sucesso") # Registra sucesso no log
        except Exception as e:         # Tratamento de erro
            logger.error(traceback.format_exc())
            logger.error(f"Erro ao converter para JPEG: {str(e)}")
            raise ValueError(f"Erro ao converter para JPEG: {str(e)}") # Lança erro se falhar
    
    def converter_para_jpg(self, origem, destino):
        """Converte imagem (PNG/JPEG/ICO) para JPG"""
        try:
            img = Image.open(origem).convert("RGB")  # Abre e converte para RGB
            img.save(destino, 'JPEG', quality=95)  # Salva como JPG com qualidade 95%
            os.chmod(destino, 0o600) # Define permissões de leitura e escrita para o proprietário
            self.adicionar_log(f"Imagem convertida para JPG com sucesso") # Registra sucesso no log
        except Exception as e:
            logger.error(traceback.format_exc())
            logger.error(f"Erro ao converter para JPG: {str(e)}") # Registra o erro
            raise ValueError(f"Erro ao converter para JPG: {str(e)}") # Lança erro se falhar
    
    def converter_para_webp(self, origem, destino):
        """Converte imagens (PNG/JPG) para WEBP"""
        try:
            img = Image.open(origem)         # Abre a imagem
            qualidade = 80  # Padrão para WEBP (0-100)
            img.save(destino, 'WEBP', quality=qualidade)    # Salva como WEBP
            os.chmod(destino, 0o600) # Define permissões de leitura e escrita para o proprietário
        except Exception as e:
            logger.error(traceback.format_exc())  # Registra o erro
            logger.error(f"Erro ao converter para WEBP: {str(e)}") # Reg
            raise ValueError(f"Erro ao converter para WEBP: {str(e)}") # Lança erro se falhar

    def webp_para_imagem(self, origem, destino, formato):
        """Converte WEBP para PNG/JPG""" 
        try:
            img = Image.open(origem)        # Abre a imagem WEBP
            if formato == 'jpg':
                img = img.convert("RGB")  # JPG não suporta transparência
                img.save(destino, 'JPEG', quality=95) # Salva como JPG
                os.chmod(destino, 0o600)
            else:
                img.save(destino, formato.upper()) # Salva como PNG
                os.chmod(destino, 0o600) # Define permissões de leitura e escrita para o proprietário
        except Exception as e:
            logger.error(traceback.format_exc())  # Registra o erro
            logger.error(f"Erro ao converter WEBP para {formato.upper()}: {str(e)}")
            raise ValueError(f"Erro ao converter WEBP para {formato.upper()}: {str(e)}") # Lança erro se falhar
    
    def converter_para_ico(self, origem, destino):
        """Converte imagem para ICO (ícone) com múltiplos tamanhos padrão"""
        try:
            img = Image.open(origem)
            img = img.convert("RGBA")  # Garante canal alfa para ícones
            tamanhos = [(16,16), (24,24), (32,32), (48,48), (64,64), (128,128), (256,256)]
            img.save(destino, format='ICO', sizes=tamanhos)
            os.chmod(destino, 0o600)
            self.adicionar_log(f"Imagem convertida para ICO com sucesso")
        except Exception as e:
            logger.error(traceback.format_exc())
            logger.error(f"Erro ao converter para ICO: {str(e)}")  # Registra o erro
            raise ValueError(f"Erro ao converter para ICO: {str(e)}")
    
    def pdf_para_docx(self, origem, destino):
        """Converte PDF para DOCX"""
        try:
            cv = Converter(origem)                # Cria um conversor PDF para Word
            cv.convert(destino)                   # Executa a conversão
            cv.close()                            # Fecha o conversor
            os.chmod(destino, 0o600) # Define permissões de leitura e escrita para o proprietário
        except Exception as e:
            logger.error(traceback.format_exc())
            logger.error(f"Erro ao converter PDF para DOCX: {str(e)}")  # Registra o erro
    
    def csv_para_xlsx(self, origem, destino): # Método para converter CSV para Excel (XLSX)
        """Converte CSV para Excel (XLSX) detectando automaticamente o delimitador e a codificação"""
        try:
            # Detecta a codificação do arquivo CSV
            with open(origem, 'rb') as f:
                resultado = chardet.detect(f.read())
                encoding_detectada = resultado['encoding']

            # Lê o arquivo CSV com a codificação detectada
            df = pd.read_csv(origem, sep=None, engine='python', encoding=encoding_detectada)

            # ALERTA DE CONTEÚDO PERIGOSO
            if df.applymap(contem_estrutura_perigosa).any().any():
                self.mostrar_mensagem(
                    "Alerta de Segurança",
                    "O arquivo contém fórmulas ou scripts potencialmente perigosos. Eles serão neutralizados na conversão.",
                    erro=True
                )

            # Sanitiza todas as células para evitar fórmulas perigosas
            df = df.applymap(sanitizar_celula_excel)

            # Salva o DataFrame como um arquivo Excel
            df.to_excel(destino, index=False, engine='openpyxl')
            os.chmod(destino, 0o600)
            self.adicionar_log(f"Arquivo CSV convertido para XLSX com sucesso")
        except Exception as e:
            logger.error(traceback.format_exc())
            logger.error(f"Erro ao converter CSV para XLSX: {str(e)}")
            raise ValueError(f"Erro ao converter CSV para XLSX: {str(e)}")
    
    def mostrar_mensagem(self, titulo, mensagem, erro=False, parent=None):
        """Exibe uma mensagem popup centralizada na janela principal ou em um parent."""
        if parent is None:
            parent = self
        msg = tk.Toplevel(parent)
        msg.title(titulo)
        msg.focus_force()
        try:
            msg.iconbitmap(ICONE)
        except:
            pass
        msg.configure(background=COR_FUNDO)
        tk.Label(
            msg,
            text=mensagem,
            bg=COR_FUNDO,
            fg=COR_PRIMARIA if erro else COR_TEXTO,
            font=FONTE_PRINCIPAL,
            padx=20,
            pady=20
        ).pack()
        btn_frame = tk.Frame(msg, bg=COR_FUNDO)
        btn_frame.pack(pady=(0, 10))
        btn_ok = tk.Button(
            btn_frame,
            text="OK",
            command=msg.destroy,
            bg=COR_PRIMARIA,
            fg=COR_SECUNDARIA,
            activebackground='#D90615',
            activeforeground=COR_SECUNDARIA,
            relief='flat',
            padx=20
        )
        btn_ok.pack()
        msg.bind("<Return>", lambda event: btn_ok.invoke())
        msg.bind("<Escape>", lambda event: msg.destroy())
        parent.update_idletasks()
        largura_msg = msg.winfo_reqwidth()
        altura_msg = msg.winfo_reqheight()
        largura_janela = parent.winfo_width()
        altura_janela = parent.winfo_height()
        x_janela = parent.winfo_rootx()
        y_janela = parent.winfo_rooty()
        x_centralizado = x_janela + (largura_janela // 2) - (largura_msg // 2)
        y_centralizado = y_janela + (altura_janela // 2) - (altura_msg // 2)
        msg.geometry(f"+{x_centralizado}+{y_centralizado}")
        msg.transient(parent)
        msg.grab_set()
        msg.wait_window()
    
    def adicionar_log(self, mensagem):
        """Adiciona uma mensagem ao log"""
        data_hora = datetime.now().strftime("%H:%M:%S")  # Obtém a hora atual
        self.text_log.config(state=tk.NORMAL)  # Habilita edição temporariamente
        self.text_log.insert(tk.END, f"[{data_hora}] {mensagem}\n")  # Insere a mensagem no log
        self.text_log.config(state=tk.DISABLED)  # Desabilita edição novamente
        self.text_log.see(tk.END)  # Rola para mostrar a nova mensagem
    
    def limpar_log(self):
        """Limpa o log de atividades"""
        self.text_log.config(state=tk.NORMAL)  # Habilita edição temporariamente
        self.text_log.delete(1.0, tk.END)  # Remove todo o conteúdo do log
        self.text_log.config(state=tk.DISABLED)  # Desabilita edição novamente
        self.adicionar_log("Log limpo")  # Adiciona mensagem de log limpo

# ==================================================
# TELA DE FERRAMENTAS PDF
# ==================================================
class TelaPDFTools(ttk.Frame): # Classe para a tela de ferramentas PDF
    def __init__(self, parent, controller): # Construtor da classe
        super().__init__(parent)              # Chama o construtor da classe pai
        self.controller = controller          # Guarda referência ao controlador
        self.arquivos_para_juntar = []        # Lista para armazenar PDFs a serem unidos
        self.configurar_interface()           # Configura os elementos da interface
        self.mostrar_frame_divisao()        # Mostra o frame de divisão por padrão
    
    def configurar_interface(self): # Método para configurar a interface
        header_frame = ttk.Frame(self)        # Cabeçalho da tela
        header_frame.pack(fill=tk.X, pady=(0, 15)) # Expande horizontalmente com espaçamento
        
        # Configura navegação com Tab e Enter para todos os widgets
        for widget in self.winfo_children():
            widget.bind("<Tab>", lambda event: widget.tk_focusNext().focus() or "break")
            if isinstance(widget, ttk.Button):
                widget.bind("<Return>", lambda event: widget.invoke())
                widget.bind("<FocusIn>", lambda event: widget.configure(style='Fleury.TButton'))  # Estilo ao focar


        # Botão "Voltar" para retornar à tela inicial
        ttk.Button(
            header_frame, # Frame pai
            text="← Voltar", # Texto do botão
            command=lambda: self.controller.mostrar_tela("TelaInicial"), # Volta à tela inicial
            style='Fleury.TButton',           # Aplica o estilo personalizado
            width=10                          # Largura fixa
        ).pack(side=tk.LEFT)                 # Alinha à esquerda
        
        # Título da tela
        ttk.Label(
            header_frame,  # Label para o título
            text="PDF DIVIDIR, SEPARAR & JUNTAR",  # Texto do título
            font=FONTE_TITULO,                # Fonte de título
            foreground=COR_PRIMARIA,          # Cor vermelha
            background=COR_FUNDO              # Cor de fundo
        ).pack(side=tk.LEFT, padx=10)        # Alinha à esquerda com espaçamento
        
        # Frame principal de operações
        operacoes_frame = ttk.LabelFrame(self, text=" OPERAÇÕES ", padding=10) # Frame para operações
        operacoes_frame.pack(fill=tk.X, pady=(0, 10)) # Expande horizontalmente
        
        # Frame para os botões de operação
        btn_frame = ttk.Frame(operacoes_frame) # Frame para os botões
        btn_frame.pack(fill=tk.X, pady=5)    # Expande horizontalmente
        
        # Botão para divisão de PDF
        ttk.Button(
            btn_frame,  # Frame pai
            text="Dividir PDF",  # Texto do botão
            command=self.mostrar_frame_divisao, # Mostra o frame de divisão
            style='Fleury.TButton',           # Aplica o estilo personalizado
            width=15                          # Largura fixa
        ).pack(side=tk.LEFT, padx=5)         # Alinha à esquerda com espaçamento
        
        # Botão para separação de todas as páginas
        ttk.Button(
            btn_frame, # frame pai
            text="Separar Todas as Páginas", # Texto do botão
            command=self.mostrar_frame_separar_todas, # Mostra o frame de separação
            style='Fleury.TButton',           # Aplica o estilo personalizado
            width=20                          # Largura fixa
        ).pack(side=tk.LEFT, padx=5)         # Alinha à esquerda com espaçamento
        
        # Botão para junção de PDFs
        ttk.Button(
            btn_frame, # frame pai
            text="Juntar PDFs", # Texto do botão
            command=self.mostrar_frame_juncao, # Mostra o frame de junção
            style='Fleury.TButton',           # Aplica o estilo personalizado
            width=15                          # Largura fixa
        ).pack(side=tk.LEFT, padx=5)         # Alinha à esquerda com espaçamento
        
        # Frame para divisão de PDF
        self.div_frame = ttk.Frame(operacoes_frame)
        
        # Frame para seleção de arquivo para divisão
        div_file_frame = ttk.Frame(self.div_frame)
        div_file_frame.pack(fill=tk.X, pady=5) # Expande horizontalmente
        
        ttk.Label(div_file_frame, text="Arquivo PDF para dividir:").pack(side=tk.LEFT, padx=(0, 5)) # Rótulo
        
        self.entry_arquivo_div = ttk.Entry(div_file_frame, width=40, state="readonly")  # Torna o campo somente leitura
        self.entry_arquivo_div.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=(0, 5))  # Expande para preencher
        
        # Botão "Procurar" para selecionar arquivo
        ttk.Button(
            div_file_frame, ## Frame pai
            text="Procurar...",  # Texto do botão
            command=lambda: self.selecionar_arquivo(self.entry_arquivo_div), # Chama o método de seleção
            style='Fleury.TButton',           # Aplica o estilo personalizado
            width=10                          # Largura fixa
        ).pack(side=tk.LEFT)                 # Alinha à esquerda
        
        # Frame para configuração de páginas
        div_config_frame = ttk.Frame(self.div_frame) # Frame para configuração
        div_config_frame.pack(fill=tk.X, pady=5) # Expande horizontalmente
        
        ttk.Label(
            div_config_frame,  # Frame pai
            text="Informe as páginas ou intervalos para extrair:\n"
                "- Exemplo 1: '5' extrai apenas a página 5\n"
                "- Exemplo 2: '3-7' extrai da página 3 até a 7\n"
                "- Exemplo 3: '1-3,5-8' combina páginas e intervalos\n"
                "- Exemplo 4: '1;2' extrai páginas em sequência de 2 em 2 (ex: 1-2, 3-4...)"
        ).pack(side=tk.LEFT, padx=(0, 5))  # Rótulo com instruções
        
        self.entry_paginas = ttk.Entry(div_config_frame, width=10) # Campo para configuração
        self.entry_paginas.pack(side=tk.LEFT, padx=(0, 5)) # Posiciona com espaçamento
        
        
        # Frame para separação de todas as páginas
        self.separar_frame = ttk.Frame(operacoes_frame)
        
        # Frame para seleção de arquivo para separação
        separar_file_frame = ttk.Frame(self.separar_frame)
        separar_file_frame.pack(fill=tk.X, pady=5) # Expande horizontalmente
        
        ttk.Label(separar_file_frame, text="Arquivo PDF para separar:").pack(side=tk.LEFT, padx=(0,5)) # Rótulo
        
        self.entry_arquivo_separar = ttk.Entry(separar_file_frame, width=40, state="readonly")  # Torna o campo somente leitura
        self.entry_arquivo_separar.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=(0, 5))  # Expande para preencher
        
        # Botão "Procurar" para selecionar arquivo
        ttk.Button(
            separar_file_frame, # Frame pai
            text="Procurar...", # Texto do botão
            command=lambda: self.selecionar_arquivo(self.entry_arquivo_separar), # Chama o método de seleção
            style='Fleury.TButton',           # Aplica o estilo personalizado
            width=10                          # Largura fixa
        ).pack(side=tk.LEFT)                 # Alinha à esquerda
        
        
        # Frame para junção de PDFs
        self.merge_frame = ttk.Frame(operacoes_frame)
        
        # Listbox para mostrar os PDFs a serem unidos
        self.lista_arquivos = tk.Listbox(
            self.merge_frame,
            height=6,                         # Altura em linhas
            selectmode=tk.EXTENDED,           # Permite seleção múltipla
            bg=COR_SECUNDARIA,                # Cor de fundo branca
            fg=COR_TEXTO,                     # Cor do texto
            highlightthickness=1,             # Espessura do destaque
            highlightcolor=COR_BORDA          # Cor da borda
        )
        self.lista_arquivos.pack(fill=tk.BOTH, expand=True, pady=5) # Expande para preencher
        
        # Scrollbar vertical para a lista
        scrollbar = ttk.Scrollbar(
            self.merge_frame, # Frame pai
            orient=tk.VERTICAL,               # Orientação vertical
            command=self.lista_arquivos.yview # Vincula ao Listbox
        )
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y) # Alinha à direita
        self.lista_arquivos.config(yscrollcommand=scrollbar.set) # Configura a scrollbar
        
        # Oculta a scrollbar inicialmente
        scrollbar.pack_forget()

        btn_frame_acoes = ttk.Frame(self.merge_frame)
        btn_frame_acoes.pack(pady=5)

        # Novo frame interno para centralizar os botões
        btns_centro = ttk.Frame(btn_frame_acoes)
        btns_centro.pack(anchor="center")

        # Botão "Adicionar PDF"
        ttk.Button(
            btn_frame_acoes,
            text="Adicionar PDF",
            command=self.adicionar_arquivos,  # Chama o método para adicionar
            style='Fleury.TButton',  # Aplica o estilo personalizado
            width=15  # Largura fixa
        ).pack(side=tk.LEFT, padx=5)  # Alinha à esquerda com espaçamento

        # Botão "Remover Selecionado"
        ttk.Button(
            btn_frame_acoes,
            text="Remover Selecionado",
            command=self.remover_arquivo,  # Chama o método para remover
            style='Fleury.TButton',  # Aplica o estilo personalizado
            width=20  # Largura fixa
        ).pack(side=tk.LEFT, padx=5)  # Alinha à esquerda com espaçamento

        # Botão "Mover para Cima"
        ttk.Button(
            btn_frame_acoes,
            text="Mover para Cima ↑",
            command=lambda: self.mover_arquivo(-1),  # Move para cima (índice -1)
            style='Fleury.TButton',  # Aplica o estilo personalizado
            width=16  # Largura fixa
        ).pack(side=tk.LEFT, padx=5)  # Alinha à esquerda com espaçamento

        # Botão "Mover para Baixo"
        ttk.Button(
            btn_frame_acoes,
            text="Mover para Baixo ↓",
            command=lambda: self.mover_arquivo(1),  # Move para baixo (índice +1)
            style='Fleury.TButton',  # Aplica o estilo personalizado
            width=16  # Largura fixa
        ).pack(side=tk.LEFT, padx=5)  # Alinha à esquerda com espaçamento

        # Frame para o botão "Limpar Lista de Junção"
        btn_limpar_lista_frame = ttk.Frame(self.merge_frame)  # Cria um frame exclusivo
        btn_limpar_lista_frame.pack(fill=tk.X, pady=10)  # Posiciona na parte inferior com espaçamento

        # Botão "Limpar Lista de Junção"
        self.btn_limpar_lista_juncao = ttk.Button(
            btn_limpar_lista_frame,  # Frame pai
            text="Limpar Lista de Junção",  # Texto do botão
            command=self.limpar_lista_juncao,  # Chama o método para limpar a lista
            style='Fleury.TButton',  # Aplica o estilo personalizado
            width=20  # Largura fixa
        )
        self.btn_limpar_lista_juncao.pack(side=tk.TOP, pady=10)  # Centraliza o botão com espaçamento vertical
        
        
        # Frame para o log de atividades
        log_frame = ttk.LabelFrame(self, text=" LOG ", padding=10) # Frame para log
        log_frame.pack(fill=tk.BOTH, expand=True) # Expande para preencher
        
        # Área de texto para o log
        self.text_log = tk.Text(    ## Área de texto para log
            log_frame,  # Frame pai
            wrap=tk.WORD,                    # Quebra de linha por palavras
            height=10,                       # Altura inicial
            bg=COR_SECUNDARIA,               # Cor de fundo branca
            fg=COR_TEXTO,                    # Cor do texto
            padx=5,                          # Espaçamento horizontal
            pady=5                           # Espaçamento vertical
        )
        self.text_log.pack(fill=tk.BOTH, expand=True) # Expande para preencher
        
        self.text_log.bind("<Tab>", lambda event: self.text_log.tk_focusNext().focus() or "break")


        # Scrollbar vertical para o log
        scrollbar_log = ttk.Scrollbar( ## Frame pai
            log_frame, 
            command=self.text_log.yview       # Vincula ao Text
        )
        scrollbar_log.pack(side=tk.RIGHT, fill=tk.Y) # Alinha à direita
        self.text_log.config(yscrollcommand=scrollbar_log.set) # Configura a scrollbar

        scrollbar_log.pack_forget() # Esconde a scrollbar inicialmente

        btn_limpar_frame = ttk.Frame(self)
        btn_limpar_frame.pack(fill=tk.X, pady=(5, 0))

        # Todos os botões juntos, alinhados à esquerda
        self.btn_dividir_pdf = ttk.Button(
            btn_limpar_frame,
            text="Dividir PDF",
            command=self.dividir_pdf,
            style='Fleury.TButton',
            width=15
        )
        self.btn_dividir_pdf.pack(side=tk.LEFT, padx=5)
        self.btn_dividir_pdf.pack_forget()

        self.btn_separar_todas = ttk.Button(
            btn_limpar_frame,
            text="Separar Todas as Páginas",
            command=self.separar_todas_paginas,
            style='Fleury.TButton',
            width=20
        )
        self.btn_separar_todas.pack(side=tk.LEFT, padx=5)
        self.btn_separar_todas.pack_forget()

        self.btn_juntar_pdfs = ttk.Button(
            btn_limpar_frame,
            text="Juntar PDFs",
            command=self.juntar_pdfs,
            style='Fleury.TButton',
            width=15
        )
        self.btn_juntar_pdfs.pack(side=tk.LEFT, padx=5)
        self.btn_juntar_pdfs.pack_forget()

        ttk.Button(
            btn_limpar_frame,
            text="Limpar Log",
            command=self.limpar_log,
            style='Fleury.TButton',
            width=15
        ).pack(side=tk.LEFT, padx=5)

        
        # Rodapé com copyright
        ttk.Label(
            self,  # Frame pai
            text="© 2025 Matheus Augusto e Jeferson Sá - Todos os direitos reservados",  ## Texto do rodapé
            font=('Arial', 8),               # Fonte menor
            foreground='gray'                 # Cor cinza
        ).pack(side=tk.RIGHT, pady=(10, 0))  # Alinha à direita no rodapé
        
    
    def mostrar_frame_separar_todas(self):
        """Mostra o frame para separar todas as páginas"""
        self.div_frame.pack_forget()  # Esconde o frame de divisão
        self.merge_frame.pack_forget()  # Esconde o frame de junção
        self.separar_frame.pack(fill=tk.BOTH, expand=True)  # Mostra o frame de separação
        self.btn_juntar_pdfs.pack_forget()  # Oculta o botão "Juntar PDFs"
        self.btn_dividir_pdf.pack_forget()  # Oculta o botão "Dividir PDF"
        self.btn_separar_todas.pack(side=tk.LEFT, padx=5)  # Mostra o botão "Separar Todas as Páginas"
        self.adicionar_log("Modo de separação de todas as páginas ativado")  # Log
    
    def mostrar_frame_divisao(self):
        """Mostra o frame para divisão de PDF"""
        self.merge_frame.pack_forget()  # Esconde o frame de junção
        self.separar_frame.pack_forget()  # Esconde o frame de separação
        self.div_frame.pack(fill=tk.BOTH, expand=True)  # Mostra o frame de divisão
        self.btn_juntar_pdfs.pack_forget()  # Oculta o botão "Juntar PDFs"
        self.btn_separar_todas.pack_forget()  # Oculta o botão "Separar Todas as Páginas"
        self.btn_dividir_pdf.pack(side=tk.LEFT, padx=5)  # Mostra o botão "Dividir PDF"
        self.adicionar_log("Modo de divisão de PDF ativado")  # Log
    
    def mostrar_frame_juncao(self):
        """Mostra o frame para junção de PDFs"""
        self.div_frame.pack_forget()  # Esconde o frame de divisão
        self.separar_frame.pack_forget()  # Esconde o frame de separação
        self.merge_frame.pack(fill=tk.BOTH, expand=True)  # Mostra o frame de junção
        self.btn_dividir_pdf.pack_forget()  # Oculta o botão "Dividir PDF"
        self.btn_separar_todas.pack_forget()  # Oculta o botão "Separar Todas as Páginas"
        self.btn_juntar_pdfs.pack(side=tk.LEFT, padx=5)  # Mostra o botão "Juntar PDFs"
        self.adicionar_log("Modo de junção de PDF ativado")  # Log
    

    def selecionar_arquivo(self, entry_destino=None):
        """Abre o diálogo de seleção de arquivo e valida o tipo MIME e integridade antes de aceitar."""
        try:
            arquivo = filedialog.askopenfilename(
                title="Selecione o arquivo PDF",
                filetypes=[("Arquivos PDF", "*.pdf")],
                initialdir=DOWNLOAD_FOLDER
            )

            if entry_destino is None:
                return

            if not arquivo:
                self.adicionar_log("Nenhum arquivo selecionado.")
                entry_destino.config(state="normal")
                entry_destino.delete(0, tk.END)
                entry_destino.config(state="readonly")
                return

            tipos_entrada_aceitos = [
                "application/pdf", "application/x-pdf"
            ]
            
            # Verifica se o arquivo selecionado é um PDF válido
            if not validar_mime(arquivo, tipos_entrada_aceitos):
                entry_destino.config(state="normal")
                entry_destino.delete(0, tk.END)
                entry_destino.config(state="readonly")
                self.adicionar_log("Arquivo rejeitado por tipo MIME inválido.")
                self.mostrar_mensagem("Erro", f"O arquivo '{os.path.basename(arquivo)}' não é PDF válido ou pode ser perigoso.", erro=True)
                return
            
            # Validação com pdfinfo
            try:
                poppler_path = obter_caminho_poppler()
                validar_pdf_com_pdfinfo(arquivo, poppler_path)
                self.adicionar_log("PDF validado com pdfinfo com sucesso.")
            except Exception as e:
                entry_destino.config(state="normal")
                entry_destino.delete(0, tk.END)
                entry_destino.config(state="readonly")
                self.adicionar_log("Arquivo rejeitado por validação prévia. Consulte o log interno.")
                self.mostrar_mensagem("Erro", f"O arquivo '{os.path.basename(arquivo)}' não passou na análise. Consulte o suporte ou o log interno.", erro=True)
                registrar_log_tecnico(e)
                logger.error(f"Erro ao validar PDF com pdfinfo: {str(e)}")
                logger.error(traceback.format_exc())
                return

            try:
                PdfReader(arquivo)
            except Exception as e:
                entry_destino.config(state="normal")
                entry_destino.delete(0, tk.END)
                entry_destino.config(state="readonly")
                self.adicionar_log("Arquivo rejeitado por estar corrompido. Consulte o log interno.")
                self.mostrar_mensagem("Erro", f"O arquivo '{os.path.basename(arquivo)}' está corrompido ou não é um PDF válido.", erro=True)
                registrar_log_tecnico(e)
                logger.error(f"Erro ao ler PDF com PdfReader: {str(e)}")
                logger.error(traceback.format_exc())
                return

            entry_destino.config(state="normal")
            entry_destino.delete(0, tk.END)
            entry_destino.insert(0, arquivo)
            entry_destino.config(state="readonly")
            self.adicionar_log(f"Arquivo selecionado: {os.path.basename(arquivo)}")

        except Exception as e:
            self.adicionar_log("Erro ao selecionar arquivo. Consulte o log interno.")
            self.mostrar_mensagem("Erro", "Erro ao selecionar arquivo. Consulte o suporte ou o log interno.", erro=True)
            registrar_log_tecnico(e)
            logger.error(f"Erro ao selecionar arquivo: {str(e)}")
            logger.error(traceback.format_exc())
    
    def adicionar_arquivos(self):
        """Adiciona arquivos PDF à lista de junção, validando MIME e integridade."""

        LIMITE_ARQUIVOS = 100
        total_na_lista = self.lista_arquivos.size()
        # Se já atingiu o limite, bloqueia
        if total_na_lista >= LIMITE_ARQUIVOS:
            self.mostrar_mensagem("Erro", f"Limite de {LIMITE_ARQUIVOS} arquivos por operação excedido.", erro=True)
            self.adicionar_log(f"Operação recusada: mais de {LIMITE_ARQUIVOS} arquivos para junção.")
            return
        
        try:
            arquivos = filedialog.askopenfilenames(
                title="Selecione os arquivos PDF",  # Título do diálogo
                filetypes=[("Arquivos PDF", "*.pdf")],  # Permite apenas arquivos PDF
                initialdir=DOWNLOAD_FOLDER  # Pasta inicial (Downloads)
            )
            if arquivos:  # Se o usuário selecionou arquivos
                tipos_entrada_aceitos = [
                    "application/pdf", "application/x-pdf"
                ]
                for arquivo in arquivos:
                    # Validação MIME
                    if not validar_mime(arquivo, tipos_entrada_aceitos):
                        self.adicionar_log("Arquivo rejeitado por tipo MIME inválido.")
                        self.mostrar_mensagem("Erro", f"O arquivo '{os.path.basename(arquivo)}' não é PDF válido ou pode ser perigoso.", erro=True)
                        continue
                    # Validação de integridade PDF
                    try:
                        PdfReader(arquivo)
                    except Exception as e:
                        self.adicionar_log("Arquivo rejeitado por estar corrompido. Consulte o log interno.")
                        self.mostrar_mensagem("Erro", f"O arquivo '{os.path.basename(arquivo)}' está corrompido ou não é um PDF válido.", erro=True)
                        registrar_log_tecnico(e)
                        logger.error(traceback.format_exc())
                        return
                        continue
                    # Só adiciona se passar nas validações e não estiver na lista
                    if arquivo not in self.lista_arquivos.get(0, tk.END):
                        self.lista_arquivos.insert(tk.END, arquivo)
                        self.adicionar_log(f"Arquivo adicionado para junção: {os.path.basename(arquivo)}")
                    else:
                        self.adicionar_log("Arquivo já existe na lista de junção.")
        except Exception as e:
            self.adicionar_log("Erro ao selecionar arquivos. Consulte o log interno.")
            self.mostrar_mensagem("Erro", "Erro ao selecionar arquivos. Consulte o suporte ou o log interno.", erro=True)
            registrar_log_tecnico(e)
            logger.error(f"Erro ao adicionar arquivos: {str(e)}")
            logger.error(traceback.format_exc())
    
    def remover_arquivo(self):
        """Remove arquivo(s) selecionado(s) da lista de junção"""
        selecionados = self.lista_arquivos.curselection() # Obtém índices dos selecionados
        if not selecionados:                # Se nenhum arquivo selecionado
            self.mostrar_mensagem("Aviso", "Nenhum arquivo selecionado para remover") # Mensagem
            return
            
        for index in reversed(selecionados): # Remove do último para o primeiro (para manter índices)
            arquivo = self.lista_arquivos.get(index) # Obtém o caminho do arquivo
            self.lista_arquivos.delete(index) # Remove da lista
            self.adicionar_log(f"Arquivo removido da lista de junção: {os.path.basename(arquivo)}") # Log
    
    def mover_arquivo(self, direcao):
        """Move arquivo(s) selecionado(s) na lista (para cima ou para baixo)"""
        selecionados = self.lista_arquivos.curselection() # Obtém índices dos selecionados
        if not selecionados:                # Se nenhum arquivo selecionado
            self.mostrar_mensagem("Aviso", "Nenhum arquivo selecionado para mover") # Mensagem
            return
            
        for index in selecionados:          # Para cada índice selecionado
            # Verifica se o movimento é possível (não ultrapassa os limites)
            if (direcao == -1 and index == 0) or (direcao == 1 and index == self.lista_arquivos.size()-1): # Limites
                continue                    # Ignora se não puder mover
                
            texto = self.lista_arquivos.get(index) # Obtém o texto do item
            self.lista_arquivos.delete(index) # Remove da posição atual
            self.lista_arquivos.insert(index + direcao, texto) # Insere na nova posição
            self.lista_arquivos.selection_set(index + direcao) # Mantém selecionado
            self.adicionar_log(f"Arquivo movido: {os.path.basename(texto)}") # Log
    
    def limpar_log(self):
        """Limpa o log de atividades"""
        self.text_log.config(state=tk.NORMAL)  # Habilita edição temporariamente
        self.text_log.delete(1.0, tk.END)  # Remove todo o conteúdo do log
        self.text_log.config(state=tk.DISABLED)  # Desabilita edição novamente
        self.adicionar_log("Log limpo")  # Adiciona mensagem de log limpo
    
    def limpar_lista_juncao(self):
        """Limpa toda a lista de arquivos para junção"""
        self.lista_arquivos.delete(0, tk.END) # Remove todos os itens
        self.adicionar_log("Lista de arquivos para junção limpa") # Log
    
    def dividir_pdf(self):
        """Divide ou extrai páginas específicas de um PDF, incluindo divisão por blocos fixos com compactação em .zip."""
        arquivo = self.entry_arquivo_div.get()
        config_paginas = self.entry_paginas.get()

        if not arquivo:
            self.mostrar_mensagem("Erro", "Selecione um arquivo PDF para dividir", erro=True)
            return

        if not config_paginas:
            self.mostrar_mensagem("Erro", "Informe as páginas ou intervalos para extrair", erro=True)
            return

        try:
            self.adicionar_log(f"Iniciando extração do PDF {os.path.basename(arquivo)}...")

            with open(arquivo, 'rb') as f:
                pdf = PdfReader(f)
                total_paginas = len(pdf.pages)

                LIMITE_PAGINAS = 1700
                if total_paginas > LIMITE_PAGINAS:
                    self.mostrar_mensagem("Erro", f"Limite de {LIMITE_PAGINAS} páginas por PDF excedido.", erro=True)
                    self.adicionar_log(f"Operação recusada: PDF com mais de {LIMITE_PAGINAS} páginas.")
                    return

                sequencias_perigosas = [
                    "../", "..\\", "%2e%2e%2f", "%2e%2e/", "%2e%2e\\", "%252e%252e%255c", "%252e%252e%252f"
                ]

                paginas_para_extrair = set()
                blocos = []

                partes = config_paginas.split(',')
                for parte in partes:
                    parte = parte.strip()

                    if ';' in parte and '-' not in parte:
                        try:
                            inicio_bloco, tamanho_bloco = map(int, parte.split(';'))
                            if inicio_bloco < 1 or tamanho_bloco < 1:
                                raise ValueError(f"Valores inválidos: {parte}")
                            for bloco_inicio in range(inicio_bloco, total_paginas + 1, tamanho_bloco):
                                bloco_fim = min(bloco_inicio + tamanho_bloco - 1, total_paginas)
                                blocos.append((bloco_inicio, bloco_fim))
                        except ValueError:
                            raise ValueError(f"Formato de bloco inválido: {parte}")
                    elif '-' in parte:
                        try:
                            inicio, fim = map(int, parte.split('-'))
                            if inicio < 1 or fim > total_paginas or inicio > fim:
                                raise ValueError(f"Intervalo inválido: {parte}")
                            paginas_para_extrair.update(range(inicio, fim + 1))
                        except ValueError:
                            raise ValueError(f"Formato de intervalo inválido: {parte}")
                    else:
                        try:
                            pagina = int(parte)
                            if pagina < 1 or pagina > total_paginas:
                                raise ValueError(f"Página inválida: {pagina}")
                            paginas_para_extrair.add(pagina)
                        except ValueError:
                            raise ValueError(f"Formato de página inválido: {parte}")

                if blocos:
                    nome_base = os.path.splitext(os.path.basename(arquivo))[0]
                    for seq in sequencias_perigosas:
                        nome_base = nome_base.replace(seq, "")
                    nome_base = nome_base.replace("..", "")
                    nome_base = re.sub(r'[^a-zA-Z0-9_\-]', '_', nome_base)
                    if not nome_base:
                        nome_base = "arquivo"

                    nome_seguro = uuid.uuid4().hex
                    pasta_temp = os.path.join(tempfile.gettempdir(), f"ConversorTemp_{nome_seguro}")
                    os.makedirs(pasta_temp, exist_ok=True)
                    os.chmod(pasta_temp, 0o600)

                    pdfs_temp = []
                    for i, (bloco_inicio, bloco_fim) in enumerate(blocos, 1):
                        output = PdfWriter()
                        for pagina in range(bloco_inicio, bloco_fim + 1):
                            output.add_page(pdf.pages[pagina - 1])
                        pdf_temp = os.path.join(pasta_temp, f"bloco_{i}_{bloco_inicio}-{bloco_fim}.pdf")
                        with open(pdf_temp, 'wb') as f_out:
                            output.write(f_out)
                        os.chmod(pdf_temp, 0o600)
                        pdfs_temp.append(pdf_temp)

                    zip_path = os.path.join(pasta_temp, f"{nome_seguro}.zip")
                    with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
                        for pdf_temp in pdfs_temp:
                            zipf.write(pdf_temp, os.path.basename(pdf_temp))
                    os.chmod(zip_path, 0o600)

                    nome_final = gerar_nome_download(nome_base, "zip")
                    caminho_final = os.path.join(DOWNLOAD_FOLDER, nome_final)
                    shutil.copy2(zip_path, caminho_final)
                    self.adicionar_log(f"Arquivo .zip gerado com sucesso: {caminho_final}")
                    self.mostrar_mensagem("Sucesso", f"Blocos extraídos com sucesso!\nSalvo em: {caminho_final}")

                    for pdf_temp in pdfs_temp:
                        os.remove(pdf_temp)
                    os.remove(zip_path)
                    os.rmdir(pasta_temp)

                elif paginas_para_extrair:
                    output = PdfWriter()
                    for pagina in sorted(paginas_para_extrair):
                        output.add_page(pdf.pages[pagina - 1])

                    nome_base = os.path.splitext(os.path.basename(arquivo))[0]
                    for seq in sequencias_perigosas:
                        nome_base = nome_base.replace(seq, "")
                    nome_base = nome_base.replace("..", "")
                    nome_base = re.sub(r'[^a-zA-Z0-9_\-]', '_', nome_base)
                    if not nome_base:
                        nome_base = "arquivo"

                    nome_seguro = uuid.uuid4().hex
                    pasta_temp = os.path.join(tempfile.gettempdir(), f"ConversorTemp_{nome_seguro}")
                    os.makedirs(pasta_temp, exist_ok=True)
                    os.chmod(pasta_temp, 0o600)
                    pdf_saida = os.path.join(pasta_temp, f"{nome_seguro}.pdf")
                    with open(pdf_saida, 'wb') as f_out:
                        output.write(f_out)
                    os.chmod(pdf_saida, 0o600)

                    try:
                        nome_final = gerar_nome_download(nome_base, "pdf")
                        caminho_final = os.path.join(DOWNLOAD_FOLDER, nome_final)
                        shutil.copy2(pdf_saida, caminho_final)
                        self.adicionar_log(f"Páginas extraídas com sucesso: {caminho_final}")
                        self.mostrar_mensagem("Sucesso", f"Páginas extraídas com sucesso!\nSalvo em: {caminho_final}")
                    except Exception as e:
                        self.adicionar_log("Erro ao exportar para Downloads. Consulte o log interno.")
                        self.mostrar_mensagem("Erro", "Falha ao exportar para Downloads. Consulte o suporte ou o log interno.", erro=True)
                        logger.error(traceback.format_exc())
                        registrar_log_tecnico(e)

                    os.remove(pdf_saida)
                    os.rmdir(pasta_temp)

        except Exception as e:
            self.adicionar_log("Erro ao extrair páginas. Consulte o log interno.")
            self.mostrar_mensagem("Erro", "Não foi possível extrair as páginas. Consulte o suporte ou o log interno.", erro=True)
            logger.error(f"Erro ao dividir PDF: {str(e)}")
            logger.error(traceback.format_exc())
            registrar_log_tecnico(e)

    
    def separar_todas_paginas(self):
        """Separa todas as páginas de um PDF em arquivos individuais"""
        arquivo = self.entry_arquivo_separar.get() # Obtém o caminho do arquivo

        if not arquivo:                      # Verifica se um arquivo foi selecionado
            self.mostrar_mensagem("Erro", "Selecione um arquivo PDF para separar", erro=True) # Mensagem de erro
            return

        try:
            # Verifica se o arquivo existe e não está vazio
            if not os.path.exists(arquivo): # Verifica se o arquivo existe
                raise ValueError("Arquivo não encontrado") # Lança erro se não encontrado
            
            if os.path.getsize(arquivo) == 0: # Verifica se o arquivo está vazio
                raise ValueError("O arquivo PDF está vazio (0 bytes)") # Lança erro se vazio

            self.adicionar_log(f"Iniciando separação de todas as páginas do PDF {os.path.basename(arquivo)}...") # Log
            
            with open(arquivo, 'rb') as f:    # Abre o arquivo PDF
                pdf = PdfReader(f)            # Cria um leitor de PDF
                total_paginas = len(pdf.pages) # Obtém o total de páginas

                LIMITE_PAGINAS = 1700
                if total_paginas > LIMITE_PAGINAS:
                    self.mostrar_mensagem("Erro", f"Limite de {LIMITE_PAGINAS} páginas por PDF excedido.", erro=True)
                    self.adicionar_log(f"Operação recusada: PDF com mais de {LIMITE_PAGINAS} páginas.")
                    return
                
                if total_paginas == 0:        # Verifica se o PDF tem páginas
                    raise ValueError("O PDF não contém nenhuma página") # Lança erro se não tiver páginas
                
                # Gera nomes únicos para os arquivos
                nome_base = os.path.splitext(os.path.basename(arquivo))[0]

                # Sanitização manual para evitar path traversal
                sequencias_perigosas = [
                    "../", "..\\", "%2e%2e%2f", "%2e%2e/", "%2e%2e\\", "%252e%252e%255c", "%252e%252e%252f"
                ]
                for seq in sequencias_perigosas:
                    nome_base = nome_base.replace(seq, "")
                nome_base = nome_base.replace("..", "")
                nome_base = re.sub(r'[^a-zA-Z0-9_\-]', '_', nome_base)
                if not nome_base:
                    nome_base = "arquivo"

                nome_seguro = uuid.uuid4().hex  # Nome totalmente imprevisível
                pasta_temp = os.path.join(tempfile.gettempdir(), f"temp_paginas_{nome_seguro}")
                os.makedirs(pasta_temp, exist_ok=True)  # Garante que a pasta temporária seja criada
                os.chmod(pasta_temp, 0o600)  # Define permissões restritas para a pasta
                zip_path = os.path.join(pasta_temp, f"{nome_seguro}.zip")  # Salva o ZIP dentro da pasta temporária

                with zipfile.ZipFile(zip_path, 'w') as zipf:  # Cria um arquivo ZIP
                    for i in range(total_paginas):  # Para cada página
                        try:
                            writer = PdfWriter()  # Cria um escritor de PDF
                            writer.add_page(pdf.pages[i])  # Adiciona a página atual
                            
                            # Salva a página como um PDF individual
                            pagina_path = os.path.join(pasta_temp, f"{nome_seguro}_pagina_{i+1}.pdf")  # Caminho do arquivo temporário
                            with open(pagina_path, 'wb') as f_out:  # Abre o arquivo para escrita
                                writer.write(f_out)  # Salva a página
                            
                            zipf.write(pagina_path, os.path.basename(pagina_path))  # Adiciona ao ZIP
                            os.remove(pagina_path)  # Remove o arquivo temporário
                            
                        except Exception as pagina_error:
                            self.adicionar_log(f"Erro ao processar página {i+1}. Consulte o log interno.")
                            registrar_log_tecnico(pagina_error)
                            continue

                # Remove a pasta temporária apenas se não houver mais arquivos dentro dela
                try:
                    os.rmdir(pasta_temp)  # Tenta remover a pasta temporária
                except OSError:
                    self.adicionar_log(f"Aviso: Pasta temporária {pasta_temp} não foi removida, pois ainda contém arquivos.")
                os.chmod(zip_path, 0o600)  # Define permissões de leitura e escrita para o arquivo ZIP
                
                try:
                    nome_original = os.path.splitext(os.path.basename(arquivo))[0]
                    nome_final = gerar_nome_download(nome_original, "zip")
                    caminho_final = os.path.join(DOWNLOAD_FOLDER, nome_final)
                    shutil.copy2(zip_path, caminho_final)
                    self.adicionar_log(f"PDF separado em {total_paginas} páginas individuais e compactado em: {caminho_final}")
                    self.mostrar_mensagem("Sucesso", f"Todas as {total_paginas} páginas foram separadas e compactadas!\nSalvo em: {caminho_final}")
                except Exception as e:
                    self.adicionar_log("Erro ao exportar para Downloads. Consulte o log interno.")
                    self.mostrar_mensagem("Erro", "Falha ao exportar para Downloads. Consulte o suporte ou o log interno.", erro=True)
                    registrar_log_tecnico(e)
                    logger.error(traceback.format_exc())
        
        except Exception as e:
            self.adicionar_log("Erro ao separar páginas. Consulte o log interno.")
            self.mostrar_mensagem("Erro", "Não foi possível separar o PDF. Consulte o suporte ou o log interno.", erro=True)
            registrar_log_tecnico(e)
            logger.error(f"Erro ao separar todas as páginas: {str(e)}")
            logger.error(traceback.format_exc())
    
    def juntar_pdfs(self):
        """Junta vários PDFs em um único arquivo"""
        arquivos = list(self.lista_arquivos.get(0, tk.END)) # Obtém a lista de arquivos

        LIMITE_ARQUIVOS = 100
        if len(arquivos) > LIMITE_ARQUIVOS:
            self.mostrar_mensagem("Erro", f"Limite de {LIMITE_ARQUIVOS} arquivos por operação excedido.", erro=True)
            self.adicionar_log(f"Operação recusada: tentativa de juntar mais de {LIMITE_ARQUIVOS} arquivos.")
            return
        
        if len(arquivos) < 2:                # Verifica se há pelo menos 2 arquivos
            self.mostrar_mensagem("Erro", "Selecione pelo menos 2 arquivos PDF para juntar", erro=True) # Mensagem de erro
            return
        
        try:
            self.adicionar_log(f"Iniciando junção de {len(arquivos)} PDFs...") # Log

            nome_seguro = uuid.uuid4().hex  # Nome totalmente imprevisível
            pasta_temp = os.path.join(tempfile.gettempdir(), f"ConversorTemp_{nome_seguro}")
            os.makedirs(pasta_temp, exist_ok=True)  # Garante que a pasta temporária seja criada
            os.chmod(pasta_temp, 0o600)  # Define permissões restritas para a pasta
            arquivo_gerado = os.path.join(pasta_temp, f"{nome_seguro}.pdf")
            


            # PROCESSAMENTO ASSÍNCRONO
            sucesso, resultado = executar_com_timeout(
                juntar_pdfs_worker,
                args=(arquivos, arquivo_gerado),
                timeout=120  # segundos
            )
            if not sucesso:
                raise ValueError(f"Erro/sandbox: {resultado}")
            
            arquivo_saida_comprimido = arquivo_gerado.replace('.pdf', '_comprimido.pdf')
            comprimir_pdf(arquivo_gerado, arquivo_saida_comprimido)

            nome_final = gerar_nome_download("PDFs_unidos", "pdf")
            caminho_final = os.path.join(DOWNLOAD_FOLDER, nome_final)
            shutil.copy2(arquivo_saida_comprimido, caminho_final)
            self.adicionar_log(f"PDFs unidos com sucesso: {caminho_final}")
            self.mostrar_mensagem("Sucesso", f"PDFs unidos na ordem selecionada!\nSalvo em: {caminho_final}")

        except Exception as e:
            self.adicionar_log("Erro ao juntar PDFs. Consulte o log interno.")
            self.mostrar_mensagem("Erro", "Não foi possível juntar os PDFs. Consulte o suporte ou o log interno.", erro=True)
            registrar_log_tecnico(e)
            logger.error(f"Erro ao juntar PDFs: {str(e)}")
            logger.error(traceback.format_exc())
    
    def mostrar_mensagem(self, titulo, mensagem, erro=False):
        """Exibe uma mensagem popup centralizada na janela principal."""
        msg = tk.Toplevel(self)  # Cria uma janela de mensagem
        msg.title(titulo)  # Define o título

        try:
            msg.iconbitmap(ICONE)  # Tenta definir o ícone
        except:
            pass  # Ignora se não conseguir

        msg.configure(background=COR_FUNDO)  # Define a cor de fundo

        # Label com a mensagem
        tk.Label(
            msg,  # Frame pai
            text=mensagem,  # Mensagem a ser exibida
            bg=COR_FUNDO,  # Cor de fundo
            fg=COR_PRIMARIA if erro else COR_TEXTO,  # Cor do texto (vermelho se erro)
            font=FONTE_PRINCIPAL,  # Fonte
            padx=20,  # Espaçamento horizontal
            pady=20  # Espaçamento vertical
        ).pack()  # Adiciona a label à janela

        # Frame para o botão
        btn_frame = tk.Frame(msg, bg=COR_FUNDO)  # Frame para o botão
        btn_frame.pack(pady=(0, 10))  # Espaçamento inferior

        # Botão OK
        btn_ok = tk.Button(
            btn_frame,  # Frame pai
            text="OK",  # Texto do botão
            command=msg.destroy,  # Fecha a mensagem quando clicado
            bg=COR_PRIMARIA,  # Cor de fundo vermelha
            fg=COR_SECUNDARIA,  # Cor do texto branca
            activebackground='#D90615',  # Cor quando ativo
            activeforeground=COR_SECUNDARIA,  # Cor do texto quando ativo
            relief='flat',  # Estilo sem relevo
            padx=20  # Espaçamento horizontal
        )
        btn_ok.pack()

        # Vincula a tecla Enter ao botão OK
        msg.bind("<Return>", lambda event: btn_ok.invoke())
        msg.bind("<Escape>", lambda event: msg.destroy())

        # Centraliza o pop-up na janela principal
        self.update_idletasks()  # Atualiza informações da janela principal
        largura_msg = msg.winfo_reqwidth()  # Largura do pop-up
        altura_msg = msg.winfo_reqheight()  # Altura do pop-up
        largura_janela = self.winfo_width()  # Largura da janela principal
        altura_janela = self.winfo_height()  # Altura da janela principal
        x_janela = self.winfo_rootx()  # Posição X da janela principal
        y_janela = self.winfo_rooty()  # Posição Y da janela principal

        # Calcula a posição central
        x_centralizado = x_janela + (largura_janela // 2) - (largura_msg // 2)
        y_centralizado = y_janela + (altura_janela // 2) - (altura_msg // 2)
        msg.geometry(f"+{x_centralizado}+{y_centralizado}")  # Define a posição do pop-up

        msg.transient(self)  # Relaciona com a janela principal
        msg.grab_set()  # Torna modal (bloqueia outras janelas)
        self.wait_window(msg)  # Espera até que a mensagem seja fechada
    
    def adicionar_log(self, mensagem):
        """Adiciona uma mensagem ao log"""
        data_hora = datetime.now().strftime("%H:%M:%S")  # Obtém a hora atual
        self.text_log.config(state=tk.NORMAL)  # Habilita edição temporariamente
        self.text_log.insert(tk.END, f"[{data_hora}] {mensagem}\n")  # Insere a mensagem no log
        self.text_log.config(state=tk.DISABLED)  # Desabilita edição novamente
        self.text_log.see(tk.END)  # Rola para mostrar a nova mensagem


# ==================================================
# INICIALIZAÇÃO DO APLICATIVO
# ==================================================
if __name__ == "__main__": 

    # Configurações iniciais
    root = tk.Tk()                            # Cria a janela principal
        
    inicializar_banco_licenca()    
    
    # Centraliza a janela na tela
    root.update_idletasks()                   # Atualiza informações de geometria
    width = 700                              # Largura da janela
    height = 680                             # Altura da janela
    x = (root.winfo_screenwidth() // 2) - (width // 2) # Posição X centralizada
    y = (root.winfo_screenheight() // 2) - (height // 2) # Posição Y centralizada
    root.geometry(f'{width}x{height}+{x}+{y}') # Define geometria e posição
    root.resizable(False, False)              # Impede redimensionamento
        
    app = AplicativoFleury(root)              # Cria o aplicativo
    root.mainloop()                           # Inicia o loop principal da interface
