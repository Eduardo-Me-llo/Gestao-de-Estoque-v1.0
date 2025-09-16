# Gestao-de-Estoque-v1.0
Sistema de estoque local simples 


import sqlite3
import hashlib
import re
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import ttkbootstrap as tb
from ttkbootstrap.constants import *
from datetime import datetime
from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font
import os
import sys
import shutil
import subprocess
from PIL import Image, ImageTk
import io
import json

# ====================== CONFIGURAÇÕES INICIAIS ======================
DB = "estoque.db"
ADMIN_EMAIL = "admin.admin@aeon.com.br"
UNIDADES = ['Peça', 'Unidade', 'Pacote', 'Caixa', 'Metro', 'Conjunto', 'Cento', 'Kilo', 'Rolo', 'Rolo com 30 Metros', 'Par']
ESTOQUES = ['Estoque Garagem', 'Estoque Escritório']

# Criar diretórios para anexos
ANEXOS_DIR = "anexos"
NOTAS_ENTRADA_DIR = os.path.join(ANEXOS_DIR, "notas_fiscais", "entrada")
NOTAS_SAIDA_DIR = os.path.join(ANEXOS_DIR, "notas_fiscais", "saida")
FOTOS_DIR = os.path.join(ANEXOS_DIR, "fotos_produtos")
LOGS_DIR = os.path.join(ANEXOS_DIR, "logs")

os.makedirs(NOTAS_ENTRADA_DIR, exist_ok=True)
os.makedirs(NOTAS_SAIDA_DIR, exist_ok=True)
os.makedirs(FOTOS_DIR, exist_ok=True)
os.makedirs(LOGS_DIR, exist_ok=True)

# ====================== INICIALIZAÇÃO DO BANCO DE DADOS ======================
def init_db():
    with sqlite3.connect(DB) as conn:
        c = conn.cursor()

        c.executescript('''
            CREATE TABLE IF NOT EXISTS usuarios (
                email TEXT PRIMARY KEY,
                nome TEXT NOT NULL,
                sobrenome TEXT NOT NULL,
                senha TEXT NOT NULL,
                admin BOOLEAN DEFAULT 0
            );

            CREATE TABLE IF NOT EXISTS materiais (
                id INTEGER PRIMARY KEY,
                nome TEXT NOT NULL,
                tipo TEXT NOT NULL,
                quantidade INTEGER NOT NULL,
                marca TEXT,
                unidade TEXT,
                patrimonio INTEGER,
                descricao TEXT,
                estoque TEXT DEFAULT 'Estoque Garagem',
                nota_fiscal_entrada TEXT,  -- Caminho do arquivo
                foto TEXT  -- Caminho do arquivo
            );

            CREATE TABLE IF NOT EXISTS logs (
                id INTEGER PRIMARY KEY,
                timestamp TEXT,
                usuario TEXT,
                acao TEXT,
                produto TEXT,
                quantidade INTEGER,
                observacao TEXT,
                estoque TEXT DEFAULT 'Estoque Garagem',
                nota_fiscal_saida TEXT  -- Caminho do arquivo
            );

            CREATE TABLE IF NOT EXISTS estoques (
                nome TEXT PRIMARY KEY
            );
        ''')

        # Verificar e adicionar colunas faltantes
        c.execute("PRAGMA table_info(usuarios)")
        colunas_usuarios = [col[1] for col in c.fetchall()]
        if 'admin' not in colunas_usuarios:
            c.execute("ALTER TABLE usuarios ADD COLUMN admin BOOLEAN DEFAULT 0")
        if 'sobrenome' not in colunas_usuarios:
            c.execute("ALTER TABLE usuarios ADD COLUMN sobrenome TEXT NOT NULL DEFAULT ''")

        c.execute("PRAGMA table_info(materiais)")
        colunas_materiais = [col[1] for col in c.fetchall()]
        if 'estoque' not in colunas_materiais:
            c.execute("ALTER TABLE materiais ADD COLUMN estoque TEXT DEFAULT 'Estoque Garagem'")
            c.execute("UPDATE materiais SET estoque='Estoque Garagem' WHERE estoque IS NULL")
        if 'unidade' not in colunas_materiais:
            c.execute("ALTER TABLE materiais ADD COLUMN unidade TEXT")
        if 'nota_fiscal_entrada' not in colunas_materiais:
            c.execute("ALTER TABLE materiais ADD COLUMN nota_fiscal_entrada TEXT")
        if 'foto' not in colunas_materiais:
            c.execute("ALTER TABLE materiais ADD COLUMN foto TEXT")

        c.execute("PRAGMA table_info(logs)")
        colunas_logs = [col[1] for col in c.fetchall()]
        if 'estoque' not in colunas_logs:
            c.execute("ALTER TABLE logs ADD COLUMN estoque TEXT DEFAULT 'Estoque Garagem'")
            c.execute("UPDATE logs SET estoque='Estoque Garagem' WHERE estoque IS NULL")
        if 'nota_fiscal_saida' not in colunas_logs:
            c.execute("ALTER TABLE logs ADD COLUMN nota_fiscal_saida TEXT")

        # Carregar estoques iniciais
        c.execute("SELECT nome FROM estoques")
        estoques_existentes = [row[0] for row in c.fetchall()]

        for estoque in ESTOQUES:
            if estoque not in estoques_existentes:
                c.execute("INSERT OR IGNORE INTO estoques (nome) VALUES (?)", (estoque,))

        # Inserir admin
        senha_admin = hashlib.sha256(b"admin123").hexdigest()
        c.execute("""
            INSERT OR IGNORE INTO usuarios (email, nome, sobrenome, senha, admin)
            VALUES (?, ?, ?, ?, ?)
        """, (ADMIN_EMAIL, "Administrador", "Admin", senha_admin, 1))

        conn.commit()

# Inicializar o banco de dados
init_db()

# ====================== FUNÇÕES AUXILIARES ======================
def hash_senha(senha):
    return hashlib.sha256(senha.encode()).hexdigest()

def validar_e_extrair_email(email):
    padrao = r'^([a-zA-ZÀ-ÿ]+)\.([a-zA-ZÀ-ÿ]+)@aeon\.com\.br$'
    match = re.match(padrao, email, re.IGNORECASE)

    if not match:
        return False, None, None

    nome = match.group(1).capitalize()
    sobrenome = match.group(2).capitalize()

    return True, nome, sobrenome

def carregar_estoques():
    try:
        with sqlite3.connect(DB) as conn:
            cursor = conn.cursor()
            cursor.execute("SELECT nome FROM estoques ORDER BY nome")
            return [row[0] for row in cursor.fetchall()]
    except sqlite3.Error:
        return ESTOQUES.copy()

def listar_materiais(estoque=None):
    try:
        with sqlite3.connect(DB) as conn:
            cursor = conn.cursor()

            if estoque:
                cursor.execute("SELECT nome FROM materiais WHERE estoque=? ORDER BY nome", (estoque,))
            else:
                cursor.execute("SELECT nome FROM materiais ORDER BY nome")

            return [item[0] for item in cursor.fetchall()]
    except sqlite3.Error:
        return []

def vincular_enter(widgets, botao_principal):
    for i in range(len(widgets) - 1):
        widgets[i].bind("<Return>", lambda e, next_widget=widgets[i+1]: next_widget.focus())

    if botao_principal and widgets:
        widgets[-1].bind("<Return>", lambda e: botao_principal.invoke())

def realizar_backup():
    try:
        os.makedirs("backup", exist_ok=True)
        data = datetime.now().strftime("%Y%m%d_%H%M")
        backup_file = f"backup/estoque_{data}.db"
        shutil.copyfile(DB, backup_file)
        return True, backup_file
    except Exception as e:
        return False, str(e)

# ====================== CLASSE PRINCIPAL ======================
class EstoqueApp(tb.Window):
    def __init__(self):
        super().__init__(themename="litera")
        self.title("Gestão de Estoque Aeon")
        self.geometry("1000x600")
        self.usuario = None
        self.usuario_email = None
        self.admin = False
        self.estoque_atual = "Estoque Garagem"
        self.ordenacao = {"coluna": None, "reverso": False}
        self._criar_login()
        self.protocol("WM_DELETE_WINDOW", self._on_close)

    def _on_close(self):
        # Realizar backup ao fechar o aplicativo
        success, backup_info = realizar_backup()
        if success:
            print(f"Backup realizado: {backup_info}")
        else:
            print(f"Erro no backup: {backup_info}")
        self.destroy()

    # -------------------- LOGIN --------------------
    def _criar_login(self):
        self._limpar_interface()
        frame = tb.Frame(self, padding=40)
        frame.pack(expand=True)

        tb.Label(frame, text="Email:").grid(row=0, column=0, pady=5)
        self.email_entry = tb.Entry(frame, width=30)
        self.email_entry.grid(row=0, column=1)

        tb.Label(frame, text="Senha:").grid(row=1, column=0, pady=5)
        self.senha_entry = tb.Entry(frame, show="*", width=30)
        self.senha_entry.grid(row=1, column=1)

        self.btn_entrar = tb.Button(frame, text="Entrar", command=self._fazer_login, 
                                   bootstyle="success")
        self.btn_entrar.grid(row=2, columnspan=2, pady=10)

        tb.Button(frame, text="Cadastrar", command=self._janela_cadastro,
                 bootstyle="secondary").grid(row=3, columnspan=2, pady=5)

        widgets = [self.email_entry, self.senha_entry]
        vincular_enter(widgets, self.btn_entrar)

    def _fazer_login(self):
        email = self.email_entry.get().strip()
        senha = self.senha_entry.get().strip()

        if not email or not senha:
            messagebox.showerror("Erro", "Preencha todos os campos!")
            return

        try:
            with sqlite3.connect(DB) as conn:
                c = conn.cursor()
                c.execute("SELECT nome, sobrenome, senha, admin FROM usuarios WHERE email=?", (email,))
                usuario = c.fetchone()

            if usuario and usuario[2] == hash_senha(senha):
                self.usuario = f"{usuario[0]} {usuario[1]}"
                self.usuario_email = email
                self.admin = bool(usuario[3])
                self._interface_principal()
            else:
                messagebox.showerror("Erro", "Credenciais inválidas!")

        except sqlite3.OperationalError as e:
            if "no such column" in str(e):
                init_db()
                self._fazer_login()
            else:
                messagebox.showerror("Erro", f"Falha no login: {str(e)}")
        except Exception as e:
            messagebox.showerror("Erro", f"Falha no login: {str(e)}")

    # -------------------- CADASTRO DE USUÁRIOS --------------------
    def _janela_cadastro(self):
        win = tk.Toplevel(self)
        win.title("Cadastro de Usuário")
        win.geometry("300x300")
        win.transient(self)
        win.grab_set()

        frame = tb.Frame(win, padding=20)
        frame.pack(expand=True, fill='both')

        campos = [
            ("Email", tb.Entry),
            ("Senha", tb.Entry),
            ("Confirmar Senha", tb.Entry)
        ]

        self.entries_cadastro = {}
        for i, (rotulo, widget) in enumerate(campos):
            tb.Label(frame, text=rotulo+":").grid(row=i, column=0, sticky='e', pady=5)
            w = widget(frame, show="*" if "Senha" in rotulo else "")
            w.grid(row=i, column=1, pady=5)
            self.entries_cadastro[rotulo] = w

        self.btn_registrar = tb.Button(frame, text="Registrar", 
                                      command=lambda: self._registrar_usuario(win),
                                      bootstyle="success")
        self.btn_registrar.grid(row=3, columnspan=2, pady=10)

        widgets = [
            self.entries_cadastro['Email'],
            self.entries_cadastro['Senha'],
            self.entries_cadastro['Confirmar Senha']
        ]
        vincular_enter(widgets, self.btn_registrar)

    def _registrar_usuario(self, janela):
        email = self.entries_cadastro['Email'].get().strip()
        senha = self.entries_cadastro['Senha'].get().strip()
        confirmar = self.entries_cadastro['Confirmar Senha'].get().strip()

        if not all([email, senha, confirmar]):
            messagebox.showerror("Erro", "Preencha todos os campos!")
            return

        if senha != confirmar:
            messagebox.showerror("Erro", "As senhas não coincidem!")
            return

        valido, nome, sobrenome = validar_e_extrair_email(email)
        if not valido:
            messagebox.showerror("Erro", 
                "Email inválido! Formato correto: nome.sobrenome@aeon.com.br\n"
                "Exemplo: joao.silva@aeon.com.br")
            return

        try:
            with sqlite3.connect(DB) as conn:
                c = conn.cursor()
                c.execute("INSERT INTO usuarios (email, nome, sobrenome, senha, admin) VALUES (?,?,?,?,?)",
                         (email, nome, sobrenome, hash_senha(senha), 0))
                conn.commit()

            messagebox.showinfo("Sucesso", "Usuário cadastrado com sucesso!")
            janela.destroy()

        except sqlite3.IntegrityError:
            messagebox.showerror("Erro", "Este email já está cadastrado!")

    # -------------------- INTERFACE PRINCIPAL --------------------
    def _interface_principal(self):
        self._limpar_interface()

        # Toolbar
        toolbar = tb.Frame(self)
        toolbar.pack(fill='x', padx=5, pady=5)

        tb.Button(toolbar, text="Exportar Excel", command=self._exportar_excel,
                 bootstyle="info").pack(side='left')

        estoque_frame = tb.Frame(toolbar)
        estoque_frame.pack(side='left', padx=20)

        tb.Label(estoque_frame, text="Estoque:").pack(side='left', padx=5)
        self.combo_estoque = ttk.Combobox(estoque_frame, state="readonly")
        self.combo_estoque.pack(side='left')
        self.combo_estoque.bind("<<ComboboxSelected>>", self._mudar_estoque)
        self._atualizar_lista_estoques()

        if self.admin:
            tb.Button(estoque_frame, text="+", command=self._janela_novo_estoque,
                     bootstyle="success-outline", width=2).pack(side='left', padx=5)

        perfil_frame = tb.Frame(toolbar)
        perfil_frame.pack(side='right')

        tb.Label(perfil_frame, text=f"Bem-Vindo,{self.usuario}!").pack(side='left', padx=15)
        tb.Button(perfil_frame, text="Sair", command=self._sair,
                 bootstyle="danger-outline").pack(side='right')

        # Abas
        self.abas = tb.Notebook(self)
        self.abas.pack(fill='both', expand=True)

        self._criar_aba_estoque()
        self._criar_aba_adicionar()
        self._criar_aba_retirada()
        self._criar_aba_editar()
        self._criar_aba_logs()

        if self.admin:
            self._criar_aba_usuarios()

    def _atualizar_lista_estoques(self):
        estoques = carregar_estoques()
        self.combo_estoque['values'] = estoques
        if self.estoque_atual in estoques:
            self.combo_estoque.set(self.estoque_atual)
        elif estoques:
            self.estoque_atual = estoques[0]
            self.combo_estoque.set(estoques[0])

    def _mudar_estoque(self, event=None):
        novo_estoque = self.combo_estoque.get()
        if novo_estoque != self.estoque_atual:
            self.estoque_atual = novo_estoque
            self._atualizar_interface_estoque()
            messagebox.showinfo("Estoque Alterado", f"Estoque atual: {self.estoque_atual}")

    def _atualizar_interface_estoque(self):
        self._atualizar_estoque()
        self._atualizar_lista_retirada()
        self._atualizar_lista_edicao()
        self._atualizar_logs()

    # -------------------- NOVO ESTOQUE --------------------
    def _janela_novo_estoque(self):
        win = tk.Toplevel(self)
        win.title("Novo Estoque")
        win.geometry("300x150")
        win.transient(self)
        win.grab_set()

        frame = tb.Frame(win, padding=20)
        frame.pack(expand=True, fill='both')

        tb.Label(frame, text="Nome do Novo Estoque:").pack(pady=5)
        self.entry_novo_estoque = tb.Entry(frame)
        self.entry_novo_estoque.pack(fill='x', pady=5)
        self.entry_novo_estoque.focus()

        btn_frame = tb.Frame(frame)
        btn_frame.pack(fill='x', pady=10)

        tb.Button(btn_frame, text="Cancelar", command=win.destroy,
                 bootstyle="secondary").pack(side='right', padx=5)
        self.btn_salvar_estoque = tb.Button(btn_frame, text="Salvar", 
                                          command=lambda: self._salvar_novo_estoque(win),
                                          bootstyle="primary")
        self.btn_salvar_estoque.pack(side='right', padx=5)

        self.entry_novo_estoque.bind("<Return>", lambda e: self.btn_salvar_estoque.invoke())

    def _salvar_novo_estoque(self, janela):
        nome_estoque = self.entry_novo_estoque.get().strip()

        if not nome_estoque:
            messagebox.showerror("Erro", "Digite um nome para o estoque!")
            return

        if nome_estoque in self.combo_estoque['values']:
            messagebox.showerror("Erro", "Este estoque já existe!")
            return

        try:
            with sqlite3.connect(DB) as conn:
                c = conn.cursor()
                c.execute("INSERT INTO estoques (nome) VALUES (?)", (nome_estoque,))
                conn.commit()

            messagebox.showinfo("Sucesso", "Estoque criado com sucesso!")
            janela.destroy()
            self._atualizar_lista_estoques()

        except sqlite3.IntegrityError:
            messagebox.showerror("Erro", "Este estoque já existe!")

    # -------------------- ABA ESTOQUE --------------------
    def _criar_aba_estoque(self):
        aba = tb.Frame(self.abas)
        self.abas.add(aba, text="Estoque")

        frame_filtros = tb.Frame(aba)
        frame_filtros.pack(fill='x', padx=10, pady=10)

        self.filtros = {}
        for i, texto in enumerate(["Nome:", "Tipo:", "Marca:"]):
            tb.Label(frame_filtros, text=texto).grid(row=0, column=i*2, padx=5, sticky='e')
            entry = tb.Entry(frame_filtros, width=20)
            entry.grid(row=0, column=i*2+1, padx=5, sticky='ew')
            self.filtros[texto.lower().replace(":", "")] = entry

        tb.Button(frame_filtros, text="Filtrar", command=self._atualizar_estoque,
                 bootstyle="primary").grid(row=0, column=6, padx=5)
        tb.Button(frame_filtros, text="Limpar", command=self._limpar_filtros,
                 bootstyle="secondary").grid(row=0, column=7, padx=5)

        frame_tabela = tb.Frame(aba)
        frame_tabela.pack(fill='both', expand=True)

        cols = ("ID", "Nome", "Tipo", "Quantidade", "Marca", "Unidade")
        self.tree = ttk.Treeview(frame_tabela, columns=cols, show='headings', selectmode='browse')

        for col in cols:
            self.tree.heading(col, text=col, 
                             command=lambda c=col: self._ordenar_por_coluna(c),
                             anchor='w')
            self.tree.column(col, width=100, anchor='w')

        scroll_y = ttk.Scrollbar(frame_tabela, orient='vertical', command=self.tree.yview)
        scroll_x = ttk.Scrollbar(frame_tabela, orient='horizontal', command=self.tree.xview)
        self.tree.configure(yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)

        self.tree.pack(side='left', fill='both', expand=True)
        scroll_y.pack(side='right', fill='y')
        scroll_x.pack(side='bottom', fill='x')

        self.tree.bind("<Double-1>", self._abrir_detalhes_material)

        self._atualizar_estoque()

    def _ordenar_por_coluna(self, coluna):
        if self.ordenacao["coluna"] == coluna:
            self.ordenacao["reverso"] = not self.ordenacao["reverso"]
        else:
            self.ordenacao["coluna"] = coluna
            self.ordenacao["reverso"] = False

        self._atualizar_estoque()

    # -------------------- ABA ADICIONAR --------------------
    def _criar_aba_adicionar(self):
        aba = tb.Frame(self.abas)
        self.abas.add(aba, text="Adicionar")

        campos = [
            ("Nome", tb.Entry),
            ("Tipo", lambda f: ttk.Combobox(f, values=["Eletrônico", "EPI","Ferramenta", "Obra","Marketing","Escritório","Móvel","Cabos","Outros"], state="readonly")),
            ("Quantidade", ttk.Spinbox),
            ("Marca", tb.Entry),
            ("Unidade", lambda f: ttk.Combobox(f, values=UNIDADES, state="readonly")),
            ("Patrimônio", tb.Entry),
            ("Descrição", tb.Entry)
        ]

        self.widgets_add = {}
        for i, (label, widget) in enumerate(campos):
            tb.Label(aba, text=label+":", anchor='e').grid(row=i, column=0, padx=10, pady=5, sticky='e')
            if widget == ttk.Spinbox:
                w = widget(aba, from_=1, to=9999)
            else:
                w = widget(aba)
            w.grid(row=i, column=1, padx=10, pady=5, sticky='ew')
            self.widgets_add[label.lower().replace(" ", "_")] = w

        # Campo para Nota Fiscal (Entrada)
        row_idx = len(campos)
        tb.Label(aba, text="Nota Fiscal Entrada:").grid(row=row_idx, column=0, padx=10, pady=5, sticky='e')
        frame_nf = tb.Frame(aba)
        frame_nf.grid(row=row_idx, column=1, padx=10, pady=5, sticky='ew')

        self.lbl_nf_entrada = tb.Label(frame_nf, text="Nenhum arquivo selecionado", bootstyle="secondary")
        self.lbl_nf_entrada.pack(side='left', fill='x', expand=True)

        self.btn_anexar_nf = tb.Button(frame_nf, text="Anexar", 
                                  command=lambda: self._selecionar_arquivo(
                                      self.lbl_nf_entrada, 
                                      [('Documentos', '*.pdf *.doc *.docx')]
                                  ),
                                  bootstyle="info", width=8)
        self.btn_anexar_nf.pack(side='right')

        # Campo para Foto do Produto
        row_idx += 1
        tb.Label(aba, text="Foto do Produto:").grid(row=row_idx, column=0, padx=10, pady=5, sticky='e')
        frame_foto = tb.Frame(aba)
        frame_foto.grid(row=row_idx, column=1, padx=10, pady=5, sticky='ew')

        self.lbl_foto = tb.Label(frame_foto, text="Nenhuma imagem selecionada", bootstyle="secondary")
        self.lbl_foto.pack(side='left', fill='x', expand=True)

        self.btn_anexar_foto = tb.Button(frame_foto, text="Anexar", 
                                    command=lambda: self._selecionar_arquivo(
                                        self.lbl_foto, 
                                        [('Imagens', '*.jpg *.jpeg *.png')]
                                    ),
                                    bootstyle="info", width=8)
        self.btn_anexar_foto.pack(side='right')

        self.btn_salvar_material = tb.Button(aba, text="Salvar", command=self._salvar_material,
                                       bootstyle="success")
        self.btn_salvar_material.grid(row=row_idx+1, columnspan=2, pady=10)

        widgets = list(self.widgets_add.values()) + [self.btn_anexar_nf, self.btn_anexar_foto]
        vincular_enter(widgets, self.btn_salvar_material)

    def _selecionar_arquivo(self, label_widget, filetypes):
        arquivo = filedialog.askopenfilename(
            title="Selecionar arquivo",
            filetypes=filetypes
        )

        if arquivo:
            label_widget.arquivo_selecionado = arquivo
            nome_arquivo = os.path.basename(arquivo)
            label_widget.config(text=nome_arquivo[:20] + "..." if len(nome_arquivo) > 20 else nome_arquivo)
        else:
            label_widget.arquivo_selecionado = None
            label_widget.config(text="Nenhum arquivo selecionado")

    def _salvar_anexo(self, origem, diretorio, prefixo):
        try:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            ext = os.path.splitext(origem)[1]
            nome_arquivo = f"{prefixo}_{timestamp}{ext}"
            destino = os.path.join(diretorio, nome_arquivo)

            shutil.copy2(origem, destino)

            return destino

        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao salvar anexo: {str(e)}")
            return None

    def _salvar_material(self):
        dados = {k: v.get() for k, v in self.widgets_add.items()}

        if not all([dados['nome'], dados['tipo'], dados['quantidade']]):
            messagebox.showerror("Erro", "Preencha os campos obrigatórios!")
            return

        try:
            dados['quantidade'] = int(dados['quantidade'])
            if dados['quantidade'] <= 0:
                raise ValueError
        except:
            messagebox.showerror("Erro", "Quantidade inválida!")
            return

        # Processar anexos
        nf_entrada_path = None
        foto_path = None

        if hasattr(self.lbl_nf_entrada, 'arquivo_selecionado') and self.lbl_nf_entrada.arquivo_selecionado:
            nf_entrada_path = self._salvar_anexo(
                self.lbl_nf_entrada.arquivo_selecionado, 
                NOTAS_ENTRADA_DIR,
                "nf_entrada"
            )

        if hasattr(self.lbl_foto, 'arquivo_selecionado') and self.lbl_foto.arquivo_selecionado:
            foto_path = self._salvar_anexo(
                self.lbl_foto.arquivo_selecionado, 
                FOTOS_DIR,
                "foto"
            )

        try:
            with sqlite3.connect(DB) as conn:
                c = conn.cursor()
                c.execute('''
                    INSERT INTO materiais (
                        nome, tipo, quantidade, marca, unidade, 
                        patrimonio, descricao, estoque, nota_fiscal_entrada, foto
                    ) VALUES (?,?,?,?,?,?,?,?,?,?)
                ''', (
                    dados['nome'], 
                    dados['tipo'], 
                    dados['quantidade'],
                    dados.get('marca'), 
                    dados.get('unidade'),
                    dados.get('patrimônio'), 
                    dados.get('descrição'),
                    self.estoque_atual,
                    nf_entrada_path,
                    foto_path
                ))

                # Log
                c.execute('''
                    INSERT INTO logs (
                        timestamp, usuario, acao, produto, quantidade, estoque
                    ) VALUES (?,?,?,?,?,?)
                ''', (
                    datetime.now().strftime("%Y-%m-%d %H:%M:%S"), 
                    self.usuario,
                    "ADICIONAR", 
                    dados['nome'], 
                    dados['quantidade'], 
                    self.estoque_atual
                ))

                conn.commit()

            messagebox.showinfo("Sucesso", "Material adicionado!")
            self._atualizar_estoque()
            self._limpar_campos_adicao()

            # Resetar labels de arquivos
            self.lbl_nf_entrada.config(text="Nenhum arquivo selecionado")
            self.lbl_foto.config(text="Nenhuma imagem selecionada")

            if hasattr(self.lbl_nf_entrada, 'arquivo_selecionado'):
                delattr(self.lbl_nf_entrada, 'arquivo_selecionado')
            if hasattr(self.lbl_foto, 'arquivo_selecionado'):
                delattr(self.lbl_foto, 'arquivo_selecionado')

            # Atualizar comboboxes
            self._atualizar_lista_retirada()
            self._atualizar_lista_edicao()

        except sqlite3.Error as e:
            messagebox.showerror("Erro", f"Erro no banco: {str(e)}")

    # -------------------- ABA RETIRADA --------------------
    def _criar_aba_retirada(self):
        aba = tb.Frame(self.abas)
        self.abas.add(aba, text="Retirar")

        tb.Label(aba, text="Material:").grid(row=0, column=0, padx=10, pady=5, sticky='e')
        self.combo_retirada = ttk.Combobox(aba, state="readonly")
        self.combo_retirada.grid(row=0, column=1, padx=10, pady=5, sticky='ew')
        self._atualizar_lista_retirada()

        tb.Label(aba, text="Quantidade:").grid(row=1, column=0, padx=10, pady=5, sticky='e')
        self.qt_retirada = ttk.Spinbox(aba, from_=1, to=9999)
        self.qt_retirada.grid(row=1, column=1, padx=10, pady=5, sticky='ew')

        tb.Label(aba, text="Destino:").grid(row=2, column=0, padx=10, pady=5, sticky='e')
        self.destino_retirada = tb.Entry(aba)
        self.destino_retirada.grid(row=2, column=1, padx=10, pady=5, sticky='ew')

        tb.Label(aba, text="Nota Fiscal Saída:").grid(row=3, column=0, padx=10, pady=5, sticky='e')
        frame_nf_saida = tb.Frame(aba)
        frame_nf_saida.grid(row=3, column=1, padx=10, pady=5, sticky='ew')

        self.lbl_nf_saida = tb.Label(frame_nf_saida, text="Nenhum arquivo selecionado", bootstyle="secondary")
        self.lbl_nf_saida.pack(side='left', fill='x', expand=True)

        self.btn_anexar_nf_saida = tb.Button(frame_nf_saida, text="Anexar", 
                                       command=lambda: self._selecionar_arquivo(
                                           self.lbl_nf_saida, 
                                           [('Documentos', '*.pdf *.doc *.docx')]
                                       ),
                                       bootstyle="info", width=8)
        self.btn_anexar_nf_saida.pack(side='right')

        self.btn_retirada = tb.Button(aba, text="Confirmar Retirada", 
                                command=self._processar_retirada,
                                bootstyle="warning")
        self.btn_retirada.grid(row=4, columnspan=2, pady=10)

        widgets = [
            self.combo_retirada,
            self.qt_retirada,
            self.destino_retirada,
            self.btn_anexar_nf_saida
        ]
        vincular_enter(widgets, self.btn_retirada)

    def _processar_retirada(self):
        material = self.combo_retirada.get()
        quantidade = self.qt_retirada.get()
        destino = self.destino_retirada.get().strip()

        if not all([material, quantidade, destino]):
            messagebox.showerror("Erro", "Preencha todos os campos obrigatórios!")
            return

        try:
            quantidade = int(quantidade)
            if quantidade <= 0:
                raise ValueError
        except:
            messagebox.showerror("Erro", "Quantidade inválida!")
            return

        # Processar nota fiscal de saída
        nf_saida_path = None
        if hasattr(self.lbl_nf_saida, 'arquivo_selecionado') and self.lbl_nf_saida.arquivo_selecionado:
            nf_saida_path = self._salvar_anexo(
                self.lbl_nf_saida.arquivo_selecionado, 
                NOTAS_SAIDA_DIR,
                "nf_saida"
            )

        try:
            with sqlite3.connect(DB) as conn:
                c = conn.cursor()

                # Verificar estoque
                c.execute("SELECT quantidade FROM materiais WHERE nome=? AND estoque=?", 
                         (material, self.estoque_atual))
                estoque = c.fetchone()

                if not estoque:
                    messagebox.showerror("Erro", "Material não encontrado neste estoque!")
                    return

                if estoque[0] < quantidade:
                    messagebox.showerror("Erro", f"Estoque insuficiente! Disponível: {estoque[0]}")
                    return

                # Atualizar estoque
                novo_estoque = estoque[0] - quantidade
                c.execute("UPDATE materiais SET quantidade=? WHERE nome=? AND estoque=?", 
                         (novo_estoque, material, self.estoque_atual))

                # Registrar log
                c.execute('''
                    INSERT INTO logs (
                        timestamp, usuario, acao, produto, quantidade, 
                        observacao, estoque, nota_fiscal_saida
                    ) VALUES (?,?,?,?,?,?,?,?)
                ''', (
                    datetime.now().strftime("%Y-%m-%d %H:%M:%S"), 
                    self.usuario,
                    "RETIRADA", 
                    material, 
                    quantidade, 
                    destino, 
                    self.estoque_atual,
                    nf_saida_path
                ))

                conn.commit()

            messagebox.showinfo("Sucesso", "Retirada registrada com sucesso!")
            self._atualizar_estoque()
            self._limpar_campos_retirada()

            # Resetar label de nota fiscal
            self.lbl_nf_saida.config(text="Nenhum arquivo selecionado")
            if hasattr(self.lbl_nf_saida, 'arquivo_selecionado'):
                delattr(self.lbl_nf_saida, 'arquivo_selecionado')

        except sqlite3.Error as e:
            messagebox.showerror("Erro", f"Erro no banco: {str(e)}")

    # -------------------- ABA EDIÇÃO --------------------
    def _criar_aba_editar(self):
        aba = tb.Frame(self.abas)
        self.abas.add(aba, text="Editar")

        tb.Label(aba, text="Selecionar Material:").grid(row=0, column=0, padx=10, pady=5, sticky='e')
        self.combo_editar = ttk.Combobox(aba, state="readonly")
        self.combo_editar.grid(row=0, column=1, padx=10, pady=5, sticky='ew')
        self._atualizar_lista_edicao()
        self.combo_editar.bind("<Return>", lambda e: self._carregar_dados_edicao())

        campos = [
            ("Nome", tb.Entry),
            ("Tipo", lambda f: ttk.Combobox(f, values=["Eletrônico", "EPI","Ferramenta", "Obra","Marketing","Escritório","Móvel","Cabos","Outros"], state="readonly")),
            ("Quantidade", ttk.Spinbox),
            ("Marca", tb.Entry),
            ("Unidade", lambda f: ttk.Combobox(f, values=UNIDADES, state="readonly")),
            ("Patrimônio", tb.Entry),
            ("Descrição", tb.Entry),
            ("Nota Fiscal Entrada", tb.Entry)
        ]

        self.widgets_editar = {}
        for i, (label, widget) in enumerate(campos, start=1):
            tb.Label(aba, text=label+":", anchor='e').grid(row=i, column=0, padx=10, pady=5, sticky='e')
            if widget == ttk.Spinbox:
                w = widget(aba, from_=0, to=9999)
            else:
                w = widget(aba)
            w.grid(row=i, column=1, padx=10, pady=5, sticky='ew')
            self.widgets_editar[label.lower().replace(" ", "_")] = w

        # Campo para Foto do Produto (edição)
        row_idx = len(campos) + 1
        tb.Label(aba, text="Foto do Produto:").grid(row=row_idx, column=0, padx=10, pady=5, sticky='e')
        frame_foto_editar = tb.Frame(aba)
        frame_foto_editar.grid(row=row_idx, column=1, padx=10, pady=5, sticky='ew')

        self.lbl_foto_editar = tb.Label(frame_foto_editar, text="Nenhuma imagem selecionada", bootstyle="secondary")
        self.lbl_foto_editar.pack(side='left', fill='x', expand=True)

        self.btn_anexar_foto_editar = tb.Button(frame_foto_editar, text="Anexar", 
                                    command=lambda: self._selecionar_arquivo(
                                        self.lbl_foto_editar, 
                                        [('Imagens', '*.jpg *.jpeg *.png')]
                                    ),
                                    bootstyle="info", width=8)
        self.btn_anexar_foto_editar.pack(side='right')

        btn_frame = tb.Frame(aba)
        btn_frame.grid(row=row_idx+1, column=0, columnspan=2, pady=10)

        self.btn_carregar = tb.Button(btn_frame, text="Carregar Dados", 
                                    command=self._carregar_dados_edicao,
                                    bootstyle="info")
        self.btn_carregar.pack(side='left', padx=5)

        self.btn_salvar_edicao = tb.Button(btn_frame, text="Salvar Alterações", 
                                         command=self._salvar_edicao,
                                         bootstyle="success")
        self.btn_salvar_edicao.pack(side='left', padx=5)

        tb.Button(btn_frame, text="Limpar", command=self._limpar_campos_edicao,
                 bootstyle="secondary").pack(side='left', padx=5)

        widgets_edicao = list(self.widgets_editar.values()) + [self.btn_anexar_foto_editar]
        vincular_enter(widgets_edicao, self.btn_salvar_edicao)

    def _carregar_dados_edicao(self):
        material = self.combo_editar.get()
        if not material:
            messagebox.showwarning("Aviso", "Selecione um material para editar!")
            return

        try:
            with sqlite3.connect(DB) as conn:
                c = conn.cursor()
                c.execute("SELECT * FROM materiais WHERE nome=? AND estoque=?", 
                         (material, self.estoque_atual))
                dados = c.fetchone()

            if dados:
                campos = [
                    ("nome", dados[1]),
                    ("tipo", dados[2]),
                    ("quantidade", dados[3]),
                    ("marca", dados[4]),
                    ("unidade", dados[5]),
                    ("patrimônio", dados[6]),
                    ("descrição", dados[7]),
                    ("nota_fiscal_entrada", dados[9])
                ]

                for campo, valor in campos:
                    widget = self.widgets_editar[campo]
                    if isinstance(widget, ttk.Combobox):
                        widget.set(valor or "")
                    elif isinstance(widget, ttk.Spinbox):
                        widget.delete(0, 'end')
                        widget.insert(0, valor or 0)
                    else:
                        widget.delete(0, 'end')
                        widget.insert(0, valor or "")

                # Carregar foto se existir
                if dados[10]:
                    self.lbl_foto_editar.config(text=os.path.basename(dados[10]))
                else:
                    self.lbl_foto_editar.config(text="Nenhuma imagem selecionada")

                self.widgets_editar['nome'].focus()

        except sqlite3.Error as e:
            messagebox.showerror("Erro", f"Erro ao carregar: {str(e)}")

    def _salvar_edicao(self):
        material_original = self.combo_editar.get()
        novos_dados = {k: v.get() for k, v in self.widgets_editar.items()}

        required = ['nome', 'tipo', 'quantidade']
        if any(not novos_dados[field] for field in required):
            messagebox.showerror("Erro", "Preencha todos os campos obrigatórios!")
            return

        try:
            novos_dados['quantidade'] = int(novos_dados['quantidade'])
            if novos_dados['quantidade'] < 0:
                raise ValueError
        except:
            messagebox.showerror("Erro", "Quantidade inválida! Deve ser um número positivo.")
            return

        # Processar nova foto se selecionada
        nova_foto_path = None
        if hasattr(self.lbl_foto_editar, 'arquivo_selecionado') and self.lbl_foto_editar.arquivo_selecionado:
            nova_foto_path = self._salvar_anexo(
                self.lbl_foto_editar.arquivo_selecionado, 
                FOTOS_DIR,
                "foto"
            )

        try:
            with sqlite3.connect(DB) as conn:
                c = conn.cursor()

                # Obter foto atual para não substituir se não for alterada
                c.execute("SELECT foto FROM materiais WHERE nome=? AND estoque=?", 
                         (material_original, self.estoque_atual))
                foto_atual = c.fetchone()[0]

                foto_final = nova_foto_path if nova_foto_path else foto_atual

                # Atualizar material
                c.execute('''
                    UPDATE materiais SET
                        nome=?, tipo=?, quantidade=?, marca=?, unidade=?,
                        patrimonio=?, descricao=?, nota_fiscal_entrada=?, foto=?
                    WHERE nome=? AND estoque=?
                ''', (
                    novos_dados['nome'], 
                    novos_dados['tipo'], 
                    novos_dados['quantidade'],
                    novos_dados['marca'], 
                    novos_dados['unidade'],
                    novos_dados['patrimônio'], 
                    novos_dados['descrição'],
                    novos_dados['nota_fiscal_entrada'],
                    foto_final,
                    material_original, 
                    self.estoque_atual
                ))

                # Registrar log
                c.execute('''
                    INSERT INTO logs (timestamp, usuario, acao, produto, quantidade, estoque)
                    VALUES (?,?,?,?,?,?)
                ''', (
                    datetime.now().strftime("%Y-%m-%d %H:%M:%S"), 
                    self.usuario,
                    "EDIÇÃO", 
                    novos_dados['nome'], 
                    novos_dados['quantidade'], 
                    self.estoque_atual
                ))

                conn.commit()

            messagebox.showinfo("Sucesso", "Material atualizado com sucesso!")
            self._atualizar_estoque()
            self._limpar_campos_edicao()
            self._atualizar_lista_edicao()
            self._atualizar_lista_retirada()

        except sqlite3.Error as e:
            messagebox.showerror("Erro", f"Erro no banco de dados: {str(e)}")

    # -------------------- POP-UP DE DETALHES DO MATERIAL --------------------
    def _abrir_detalhes_material(self, event):
        item_selecionado = self.tree.selection()
        if not item_selecionado:
            return

        item_id = self.tree.item(item_selecionado[0], 'values')[0]

        win = tk.Toplevel(self)
        win.title("Detalhes do Material")
        win.geometry("900x700")
        win.transient(self)
        win.grab_set()

        notebook = ttk.Notebook(win)
        notebook.pack(fill='both', expand=True, padx=10, pady=10)

        frame_info = tb.Frame(notebook)
        notebook.add(frame_info, text="Informações")

        frame_historico = tb.Frame(notebook)
        notebook.add(frame_historico, text="Histórico")

        try:
            with sqlite3.connect(DB) as conn:
                c = conn.cursor()
                c.execute("SELECT * FROM materiais WHERE id=?", (item_id,))
                material = c.fetchone()

                c.execute("""
                    SELECT timestamp, usuario, acao, quantidade, observacao, nota_fiscal_saida 
                    FROM logs 
                    WHERE produto = ?
                    ORDER BY timestamp DESC
                """, (material[1],))
                historico = c.fetchall()

        except sqlite3.Error as e:
            messagebox.showerror("Erro", f"Falha ao carregar detalhes: {str(e)}")
            win.destroy()
            return

        if material:
            campos = [
                ("ID", material[0]),
                ("Nome", material[1]),
                ("Tipo", material[2]),
                ("Quantidade", material[3]),
                ("Marca", material[4] or "-"),
                ("Unidade", material[5] or "-"),
                ("Patrimônio", material[6] or "-"),
                ("Descrição", material[7] or "-"),
                ("Estoque", material[8] or "-"),
                ("Nota Fiscal Entrada", material[9] or "-"),
                ("Foto", material[10] or "-")
            ]

            main_frame = tb.Frame(frame_info)
            main_frame.pack(fill='both', expand=True, padx=10, pady=10)

            info_frame = tb.Frame(main_frame)
            info_frame.pack(side='left', fill='both', expand=True)

            img_frame = tb.Frame(main_frame)
            img_frame.pack(side='right', padx=10, pady=10)

            for i, (rotulo, valor) in enumerate(campos):
                tb.Label(info_frame, text=f"{rotulo}:", font=("Helvetica", 10, "bold"), anchor='e'
                        ).grid(row=i, column=0, sticky='e', pady=5, padx=5)
                tb.Label(info_frame, text=valor, anchor='w'
                        ).grid(row=i, column=1, sticky='w', pady=5, padx=5)

            self.img_label = tb.Label(img_frame, text="Sem foto disponível", bootstyle="secondary")
            self.img_label.pack(pady=10)

            if material[10]:  # Caminho da foto
                try:
                    image = Image.open(material[10])
                    image.thumbnail((300, 300))
                    photo = ImageTk.PhotoImage(image)

                    self.img_label.configure(image=photo)
                    self.img_label.image = photo
                except Exception as e:
                    self.img_label.configure(text=f"Erro ao carregar foto: {str(e)}")

            if material[9]:  # Nota fiscal
                btn_nf = tb.Button(
                    info_frame, 
                    text="Abrir Nota Fiscal", 
                    command=lambda: self._abrir_arquivo(material[9]),
                    bootstyle="info"
                )
                btn_nf.grid(row=len(campos), columnspan=2, pady=10)

        # Preencher histórico
        cols = ("Data", "Ação", "Qtd", "Destino", "Nota Fiscal")
        tree_historico = ttk.Treeview(frame_historico, columns=cols, show='headings')

        for col in cols:
            tree_historico.heading(col, text=col, anchor='w')
            tree_historico.column(col, width=120, anchor='w')

        scroll_y = ttk.Scrollbar(frame_historico, orient='vertical', command=tree_historico.yview)
        tree_historico.configure(yscrollcommand=scroll_y.set)

        tree_historico.pack(side='left', fill='both', expand=True, padx=10, pady=10)
        scroll_y.pack(side='right', fill='y')

        for registro in historico:
            acao = registro[2]
            if acao == "ADICIONAR":
                acao = "➕ ADICIONAR"
            elif acao == "RETIRADA":
                acao = "➖ RETIRADA"
            elif acao == "EDIÇÃO":
                acao = "✏️ EDIÇÃO"

            tree_historico.insert('', 'end', values=(
                registro[0], 
                acao,
                registro[3], 
                registro[4], 
                registro[5] or "-"
            ))

    def _abrir_arquivo(self, caminho_arquivo):
        try:
            if os.name == 'nt':  # Windows
                os.startfile(caminho_arquivo)
            elif os.name == 'posix':  # macOS, Linux
                if sys.platform == 'darwin':
                    subprocess.run(['open', caminho_arquivo])
                else:
                    subprocess.run(['xdg-open', caminho_arquivo])
        except Exception as e:
            messagebox.showerror("Erro", f"Não foi possível abrir o arquivo: {str(e)}")

    # -------------------- ABA LOGS --------------------
    def _criar_aba_logs(self):
        aba = tb.Frame(self.abas)
        self.abas.add(aba, text="Logs")

        cols = ("Data", "Usuário", "Ação", "Produto", "Qtd", "Obs", "Estoque", "Nota Fiscal")
        self.tree_logs = ttk.Treeview(aba, columns=cols, show='headings')

        for col in cols:
            self.tree_logs.heading(col, text=col, anchor='w')
            self.tree_logs.column(col, width=120, anchor='w')

        scroll_y = ttk.Scrollbar(aba, orient='vertical', command=self.tree_logs.yview)
        scroll_x = ttk.Scrollbar(aba, orient='horizontal', command=self.tree_logs.xview)
        self.tree_logs.configure(yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)

        self.tree_logs.pack(side='left', fill='both', expand=True)
        scroll_y.pack(side='right', fill='y')
        scroll_x.pack(side='bottom', fill='x')

        self._atualizar_logs()

    def _atualizar_logs(self):
        self.tree_logs.delete(*self.tree_logs.get_children())
        try:
            with sqlite3.connect(DB) as conn:
                query = """
                    SELECT 
                        timestamp, usuario, acao, produto, quantidade, 
                        observacao, estoque, nota_fiscal_saida 
                    FROM logs
                """
                params = []

                if not self.admin:
                    query += " WHERE estoque = ?"
                    params.append(self.estoque_atual)

                query += " ORDER BY timestamp DESC"

                for row in conn.execute(query, params):
                    acao = row[2]
                    if acao == "ADICIONAR":
                        acao = "➕ ADICIONAR"
                    elif acao == "RETIRADA":
                        acao = "➖ RETIRADA"
                    elif acao == "EDIÇÃO":
                        acao = "✏️ EDIÇÃO"

                    self.tree_logs.insert('', 'end', values=(
                        row[0], row[1], acao, row[3], row[4], row[5], row[6], row[7] or "-"
                    ))
        except sqlite3.OperationalError as e:
            if "no such column" in str(e):
                init_db()
                self._atualizar_logs()

    # -------------------- ABA USUÁRIOS (ADMIN) --------------------
    def _criar_aba_usuarios(self):
        aba = tb.Frame(self.abas)
        self.abas.add(aba, text="Usuários")

        cols = ("Email", "Nome Completo", "Admin")
        self.tree_usuarios = ttk.Treeview(aba, columns=cols, show='headings')

        for col in cols:
            self.tree_usuarios.heading(col, text=col, anchor='w')
            self.tree_usuarios.column(col, width=150, anchor='w')

        scroll_y = ttk.Scrollbar(aba, orient='vertical', command=self.tree_usuarios.yview)
        scroll_x = ttk.Scrollbar(aba, orient='horizontal', command=self.tree_usuarios.xview)
        self.tree_usuarios.configure(yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)

        self.tree_usuarios.pack(side='left', fill='both', expand=True)
        scroll_y.pack(side='right', fill='y')
        scroll_x.pack(side='bottom', fill='x')

        btn_frame = tb.Frame(aba)
        btn_frame.pack(fill='x', pady=5)

        tb.Button(btn_frame, text="Adicionar", command=self._janela_novo_usuario,
                 bootstyle="success").pack(side='left', padx=5)
        tb.Button(btn_frame, text="Editar", command=self._editar_usuario,
                 bootstyle="warning").pack(side='left', padx=5)
        tb.Button(btn_frame, text="Excluir", command=self._excluir_usuario,
                 bootstyle="danger").pack(side='left', padx=5)

        self._atualizar_usuarios()

    def _janela_novo_usuario(self):
        win = tk.Toplevel(self)
        win.title("Novo Usuário")
        win.geometry("300x250")

        campos = [
            ("Email", tb.Entry),
            ("Senha", tb.Entry),
            ("Admin", ttk.Combobox)
        ]

        self.widgets_novo_user = {}
        for i, (label, widget) in enumerate(campos):
            tb.Label(win, text=label+":", anchor='e').grid(row=i, column=0, padx=10, pady=5, sticky='e')
            if widget == ttk.Combobox:
                w = widget(win, values=["Sim", "Não"], state="readonly")
                w.set("Não")
            else:
                w = widget(win, show="*" if label=="Senha" else "")
            w.grid(row=i, column=1, padx=10, pady=5, sticky='ew')
            self.widgets_novo_user[label.lower()] = w

        self.btn_salvar_user = tb.Button(win, text="Salvar", 
                                       command=lambda: self._salvar_novo_usuario(win),
                                       bootstyle="primary")
        self.btn_salvar_user.grid(row=3, columnspan=2, pady=10)

        widgets = [
            self.widgets_novo_user['email'],
            self.widgets_novo_user['senha'],
            self.widgets_novo_user['admin']
        ]
        vincular_enter(widgets, self.btn_salvar_user)

    def _salvar_novo_usuario(self, janela):
        dados = {k: v.get() for k, v in self.widgets_novo_user.items()}

        if not all(dados.values()):
            messagebox.showerror("Erro", "Todos os campos são obrigatórios!")
            return

        valido, nome, sobrenome = validar_e_extrair_email(dados['email'])
        if not valido:
            messagebox.showerror("Erro", 
                "Email inválido! Formato correto: nome.sobrenome@aeon.com.br\n"
                "Exemplo: joao.silva@aeon.com.br")
            return

        try:
            admin = 1 if dados['admin'] == "Sim" else 0
            senha_hash = hashlib.sha256(dados['senha'].encode()).hexdigest()

            with sqlite3.connect(DB) as conn:
                c = conn.cursor()
                c.execute('''
                    INSERT INTO usuarios (email, nome, sobrenome, senha, admin)
                    VALUES (?,?,?,?,?)
                ''', (dados['email'], nome, sobrenome, senha_hash, admin))
                conn.commit()

            messagebox.showinfo("Sucesso", "Usuário cadastrado!")
            janela.destroy()
            self._atualizar_usuarios()

        except sqlite3.IntegrityError:
            messagebox.showerror("Erro", "Email já cadastrado!")

    def _atualizar_usuarios(self):
        self.tree_usuarios.delete(*self.tree_usuarios.get_children())
        with sqlite3.connect(DB) as conn:
            for row in conn.execute("SELECT email, nome || ' ' || sobrenome, admin FROM usuarios"):
                self.tree_usuarios.insert('', 'end', values=(
                    row[0], 
                    row[1], 
                    "Sim" if row[2] else "Não"
                ))

    def _editar_usuario(self):
        selecionado = self.tree_usuarios.selection()
        if not selecionado:
            return

        email = self.tree_usuarios.item(selecionado[0], 'values')[0]

        with sqlite3.connect(DB) as conn:
            c = conn.cursor()
            c.execute("SELECT * FROM usuarios WHERE email=?", (email,))
            usuario = c.fetchone()

        win = tk.Toplevel(self)
        win.title("Editar Usuário")
        win.geometry("300x200")

        tb.Label(win, text="Email:").grid(row=0, column=0, padx=10, pady=5, sticky='e')
        tb.Label(win, text=usuario[0]).grid(row=0, column=1, padx=10, pady=5, sticky='w')

        tb.Label(win, text="Nome:").grid(row=1, column=0, padx=10, pady=5, sticky='e')
        tb.Label(win, text=f"{usuario[1]} {usuario[2]}").grid(row=1, column=1, padx=10, pady=5, sticky='w')

        tb.Label(win, text="Admin:").grid(row=2, column=0, padx=10, pady=5, sticky='e')
        admin_combo = ttk.Combobox(win, values=["Sim", "Não"], state="readonly")
        admin_combo.set("Sim" if usuario[4] else "Não")
        admin_combo.grid(row=2, column=1, padx=10, pady=5, sticky='ew')

        self.btn_salvar_edicao_user = tb.Button(win, text="Salvar", 
                                              command=lambda: self._salvar_edicao_usuario(
                                                  win, usuario[0], admin_combo.get()), 
                                              bootstyle="primary")
        self.btn_salvar_edicao_user.grid(row=3, columnspan=2, pady=10)

        admin_combo.bind("<Return>", lambda e: self.btn_salvar_edicao_user.invoke())

    def _salvar_edicao_usuario(self, janela, email, admin):
        try:
            admin = 1 if admin == "Sim" else 0
            with sqlite3.connect(DB) as conn:
                c = conn.cursor()
                c.execute('''
                    UPDATE usuarios SET
                        admin = ?
                    WHERE email = ?
                ''', (admin, email))
                conn.commit()

            messagebox.showinfo("Sucesso", "Usuário atualizado!")
            janela.destroy()
            self._atualizar_usuarios()

        except sqlite3.Error as e:
            messagebox.showerror("Erro", f"Erro ao atualizar: {str(e)}")

    def _excluir_usuario(self):
        selecionado = self.tree_usuarios.selection()
        if not selecionado:
            return

        email = self.tree_usuarios.item(selecionado[0], 'values')[0]

        if email == ADMIN_EMAIL:
            messagebox.showerror("Erro", "Não é possível excluir o administrador principal!")
            return

        if messagebox.askyesno("Confirmar", f"Excluir usuário {email} permanentemente?"):
            with sqlite3.connect(DB) as conn:
                c = conn.cursor()
                c.execute("DELETE FROM usuarios WHERE email=?", (email,))
                conn.commit()

            self._atualizar_usuarios()

    # ====================== UTILITÁRIOS ======================
    def _limpar_filtros(self):
        for entry in self.filtros.values():
            entry.delete(0, 'end')
        self._atualizar_estoque()

    def _limpar_campos_adicao(self):
        for widget in self.widgets_add.values():
            if isinstance(widget, ttk.Combobox):
                widget.set('')
            else:
                widget.delete(0, 'end')
        if 'nome' in self.widgets_add:
            self.widgets_add['nome'].focus()

    def _limpar_campos_retirada(self):
        self.combo_retirada.set('')
        self.qt_retirada.delete(0, 'end')
        self.destino_retirada.delete(0, 'end')
        self.combo_retirada.focus()

    def _limpar_campos_edicao(self):
        self.combo_editar.set('')
        for widget in self.widgets_editar.values():
            if isinstance(widget, ttk.Combobox):
                widget.set('')
            else:
                widget.delete(0, 'end')
        self.lbl_foto_editar.config(text="Nenhuma imagem selecionada")
        self.combo_editar.focus()

    def _atualizar_lista_retirada(self):
        materiais = listar_materiais(self.estoque_atual)
        self.combo_retirada['values'] = materiais

    def _atualizar_lista_edicao(self):
        materiais = listar_materiais(self.estoque_atual)
        self.combo_editar['values'] = materiais

    def _limpar_interface(self):
        for widget in self.winfo_children():
            widget.destroy()

    def _atualizar_estoque(self):
        self.tree.delete(*self.tree.get_children())
        query = "SELECT id, nome, tipo, quantidade, marca, unidade FROM materiais WHERE estoque = ?"
        params = [self.estoque_atual]

        if self.filtros['nome'].get():
            query += " AND nome LIKE ?"
            params.append(f"%{self.filtros['nome'].get()}%")

        if self.filtros['tipo'].get():
            query += " AND tipo LIKE ?"
            params.append(f"%{self.filtros['tipo'].get()}%")

        if self.filtros['marca'].get():
            query += " AND marca LIKE ?"
            params.append(f"%{self.filtros['marca'].get()}%")

        if self.ordenacao["coluna"]:
            coluna_sql = {
                "ID": "id",
                "Nome": "nome",
                "Tipo": "tipo",
                "Quantidade": "quantidade",
                "Marca": "marca",
                "Unidade": "unidade"
            }.get(self.ordenacao["coluna"], "id")

            direcao = "DESC" if self.ordenacao["reverso"] else "ASC"
            query += f" ORDER BY {coluna_sql} {direcao}"

        with sqlite3.connect(DB) as conn:
            for row in conn.execute(query, params):
                self.tree.insert('', 'end', values=row)

    # ====================== EXPORTAÇÃO EXCEL ======================
    def _exportar_excel(self):
        try:
            with sqlite3.connect(DB) as conn:
                cursor = conn.cursor()
                cursor.execute("PRAGMA table_info(materiais)")
                colunas = [col[1] for col in cursor.fetchall()]

                cursor.execute("SELECT * FROM materiais WHERE estoque=?", (self.estoque_atual,))
                dados = cursor.fetchall()

            wb = Workbook()
            ws = wb.active
            ws.title = self.estoque_atual[:30]

            titulo = f"Relatório do Estoque: {self.estoque_atual}"
            ws.append([titulo])
            ws.append([])

            cabecalhos = [col for col in colunas if col != "estoque" and col != "foto"]
            for col_idx, coluna in enumerate(cabecalhos, 1):
                celula = ws.cell(row=3, column=col_idx, value=coluna)
                celula.font = Font(bold=True)

            for linha_idx, linha in enumerate(dados, 4):
                valores = list(linha)
                del valores[colunas.index("estoque")]
                if "foto" in colunas:
                    del valores[colunas.index("foto") - 1]

                for col_idx, valor in enumerate(valores, 1):
                    valor_formatado = valor if valor is not None else ""
                    ws.cell(row=linha_idx, column=col_idx, value=valor_formatado)

            for col_idx, coluna in enumerate(cabecalhos, 1):
                col_letra = get_column_letter(col_idx)
                max_len = len(coluna)
                for row in dados:
                    valor = row[colunas.index(coluna)]
                    if valor is not None:
                        valor_str = str(valor)
                        if len(valor_str) > max_len:
                            max_len = len(valor_str)
                ws.column_dimensions[col_letra].width = min(max_len + 2, 50)

            ws.merge_cells(f'A1:{get_column_letter(len(cabecalhos))}1')
            titulo_cell = ws['A1']
            titulo_cell.value = titulo
            titulo_cell.font = Font(bold=True, size=14)

            nome_arquivo = f"estoque_{self.estoque_atual.replace(' ', '_')}_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"

            arquivo = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx")],
                title="Salvar relatório de estoque",
                initialfile=nome_arquivo
            )

            if arquivo:
                wb.save(arquivo)
                messagebox.showinfo("Sucesso", f"Relatório exportado com sucesso!\n{arquivo}")

                try:
                    if os.name == 'nt':
                        os.startfile(arquivo)
                    elif os.name == 'posix':
                        os.system(f'open "{arquivo}"' if sys.platform == 'darwin' else f'xdg-open "{arquivo}"')
                except:
                    pass

        except Exception as e:
            messagebox.showerror("Erro", f"Falha na exportação:\n{str(e)}")

    # ====================== CONTROLE DE PERFIL ======================
    def _sair(self):
        self.usuario = None
        self.usuario_email = None
        self.admin = False
        self.estoque_atual = "Estoque Garagem"
        self._criar_login()

if __name__ == "__main__":
    app = EstoqueApp()
    app.mainloop()
