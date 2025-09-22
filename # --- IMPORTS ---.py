# --- IMPORTS ---
import os
import time
import re
import unicodedata
import win32com.client as win32
from pathlib import Path
import tkinter as tk
from tkinter import messagebox, filedialog, ttk
import threading
from datetime import datetime, timedelta
import json
import csv
import openpyxl
from collections import defaultdict
from tkinter import simpledialog

# Inicializa o cliente Firebase
firebase = pyrebase.initialize_app(firebaseConfig)
auth = firebase.auth()
db = firebase.database()

# E-mails para cópia (CC)
copia_emails = ["email1@exemplo.com", "email2@exemplo.com", "email3@exemplo.com"]

# === TEXTO DO E-MAIL (HTML) PADRÃO ===
corpo_email_html_padrao = """
<p>Prezado (a) Coordenador (a) Estadual e Coordenador (a) de Local, bom dia!</p>
<p>Em virtude da aplicação das provas objetivas e de redação do Exame Nacional de Desempenho dos Estudantes (ENADE 2025), que ocorrerá no dia 23 de novembro de 2025, no período vespertino, encaminhamos anexo a distribuição de salas do seu município, referente à etapa de ensalamento e confirmação dos dados referentes ao espaço físico.</p>
<p><strong>Assunto: {{assunto}}</strong></p>
<p><strong>Arquivo(s) Anexado(s): {{nomes_arquivos}}</strong></p>
<p><strong>Procedimento de Validação:</strong><br>
O Coordenador deverá visualizar sua distribuição e verificar se as informações na distribuição de salas estão corretas.
<ol>
<li><strong>Validar:</strong> Confirmar se as informações estão corretas, tais como</li>
<li><strong>Recusar:</strong> Caso as informações não estejam corretas, recusar e informar o motivo e os ajustes necessários.</li>
</ol></p>
<table style="width: 100%; border-collapse: collapse;">
<thead>
<tr style="background-color: #007bff; color: white;">
<th style="border: 1px solid #ddd; padding: 8px;">Dados para confirmação</th>
<th style="border: 1px solid #ddd; padding: 8px;">Certo</th>
<th style="border: 1px solid #ddd; padding: 8px;">Errado</th>
<th style="border: 1px solid #ddd; padding: 8px;">Ajuste</th>
</tr>
</thead>
<tbody>
<tr>
<td style="border: 1px solid #ddd; padding: 8px;">Nome completo da instituição (nome exposto na fachada)?</td>
<td style="border: 1px solid #ddd; padding: 8px;"></td>
<td style="border: 1px solid #ddd; padding: 8px;"></td>
<td style="border: 1px solid #ddd; padding: 8px;"></td>
</tr>
<tr>
<td style="border: 1px solid #ddd; padding: 8px;">Endereço completo da instituição (inclusive a cidade)?</td>
<td style="border: 1px solid #ddd; padding: 8px;"></td>
<td style="border: 1px solid #ddd; padding: 8px;"></td>
<td style="border: 1px solid #ddd; padding: 8px;"></td>
</tr>
<tr>
<td style="border: 1px solid #ddd; padding: 8px;">Número de salas utilizadas e os respectivos andares?</td>
<td style="border: 1px solid #ddd; padding: 8px;"></td>
<td style="border: 1px solid #ddd; padding: 8px;"></td>
<td style="border: 1px solid #ddd; padding: 8px;"></td>
</tr>
<tr>
<td style="border: 1px solid #ddd; padding: 8px;">Capacidade de candidatos distribuídos em cada sala?</td>
<td style="border: 1px solid #ddd; padding: 8px;"></td>
<td style="border: 1px solid #ddd; padding: 8px;"></td>
<td style="border: 1px solid #ddd; padding: 8px;"></td>
</tr>
<tr>
<td style="border: 1px solid #ddd; padding: 8px;">Os Blocos foram agrupados de maneira correta?</td>
<td style="border: 1px solid #ddd; padding: 8px;"></td>
<td style="border: 1px solid #ddd; padding: 8px;"></td>
<td style="border: 1px solid #ddd; padding: 8px;"></td>
</tr>
<tr>
<td style="border: 1px solid #ddd; padding: 8px;">O quantitativo de sala por bloco está de acordo com a informação repassada por você?</td>
<td style="border: 1px solid #ddd; padding: 8px;"></td>
<td style="border: 1px solid #ddd; padding: 8px;"></td>
<td style="border: 1px solid #ddd; padding: 8px;"></td>
</tr>
<tr>
<td style="border: 1px solid #ddd; padding: 8px;">A escola com Atendimento Especializado tem a acessibilidade necessária?</td>
<td style="border: 1px solid #ddd; padding: 8px;"></td>
<td style="border: 1px solid #ddd; padding: 8px;"></td>
<td style="border: 1px solid #ddd; padding: 8px;"></td>
</tr>
</tbody>
</table>
<p>Ressaltamos que os participantes foram ensalados conforme o cadastro das instituições do seu município no SinCef. Solicitamos que você proceda à conferência dos dados.</p>
<p>Para garantir a qualidade e a excelência do nosso trabalho e cumprir os prazos estabelecidos, solicitamos a resposta a esse e-mail até o dia 20 de setembro de 2025, às 09:00h (horário de Brasília).</p>
<p>Em caso de dúvidas, entre em contato com o Cebraspe pelo e-mail <a href="mailto:enade2025@cebraspe.org.br">enade2025@cebraspe.org.br</a> ou telefone (61) 2109-5810.</p>
"""

# === FUNÇÕES ===

def normalizar_texto(texto):
    """Remove acentos e deixa tudo maiúsculo."""
    texto = unicodedata.normalize('NFKD', texto).encode('ASCII', 'ignore').decode('ASCII')
    return texto.strip().upper()

def is_valid_email(email):
    """Verifica se o formato do e-mail é válido. Regex corrigida para aceitar subdomínios."""
    regex = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
    return re.match(regex, email)

def enviar_email_com_anexos(destinatario, assunto, corpo_html, anexos_paths, cc_list=None, log_text=None):
    """Envia um e-mail via Outlook com múltiplos anexos."""
    try:
        if not destinatario or not is_valid_email(destinatario):
            if log_text: log_text.insert(tk.END, f"{datetime.now().strftime('%H:%M:%S')} ❌ Erro: E-mail do destinatário inválido: {destinatario}\n", 'red')
            return False

        outlook = win32.Dispatch('Outlook.Application')
        mail = outlook.CreateItem(0)
        
        mail.To = destinatario
        mail.Subject = assunto
        mail.HTMLBody = corpo_html
        
        if cc_list:
            mail.CC = "; ".join(cc_list)
        
        for anexo_path in anexos_paths:
            if anexo_path and os.path.exists(anexo_path):
                mail.Attachments.Add(str(anexo_path))
                if log_text: log_text.insert(tk.END, f"{datetime.now().strftime('%H:%M:%S')} 📎 Anexo adicionado: {Path(anexo_path).name}\n", 'green')
            else:
                if log_text: log_text.insert(tk.END, f"{datetime.now().strftime('%H:%M:%S')} ❌ Erro: Anexo '{anexo_path}' não encontrado ou caminho inválido. Pulando este anexo.\n", 'red')

        if not mail.Attachments.Count > 0:
            if log_text: log_text.insert(tk.END, f"{datetime.now().strftime('%H:%M:%S')} ❌ Erro: Nenhum anexo válido encontrado para este e-mail. Envio cancelado para {destinatario}.\n", 'red')
            return False

        mail.Send()
        if log_text: log_text.insert(tk.END, f"{datetime.now().strftime('%H:%M:%S')} ✅ E-mail enviado para: {destinatario}\n", 'green')
        return True
        
    except Exception as e:
        if log_text: log_text.insert(tk.END, f"{datetime.now().strftime('%H:%M:%S')} ❌ ERRO AO ENVIAR O E-MAIL para {destinatario}: {e}\n", 'red')
        return False

# === INTERFACE GRÁFICA ===
class EmailApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Automatizador de E-mails - v9.3")
        self.geometry("900x950")

        self.caminho_pasta = ""
        self.anexos_map = {}
        self.send_interval = 3 # Intervalo de 3 segundos entre envios
        
        self.current_user_uid = None
        self.token = None
        
        self.templates = defaultdict(lambda: corpo_email_html_padrao)
        self.templates["Padrão"] = corpo_email_html_padrao
        self.current_template = "Padrão"
        self.agendar_var = tk.BooleanVar()

        self.create_auth_widgets()
        
    def create_auth_widgets(self):
        self.auth_frame = tk.Frame(self)
        self.auth_frame.pack(expand=True, padx=20, pady=20)
        
        self.auth_title = tk.Label(self.auth_frame, text="Login", font=("Arial", 16))
        self.auth_title.pack(pady=10)

        tk.Label(self.auth_frame, text="E-mail:").pack()
        self.email_entry = tk.Entry(self.auth_frame, width=30)
        self.email_entry.pack(pady=5)
        
        tk.Label(self.auth_frame, text="Senha:").pack()
        self.password_entry = tk.Entry(self.auth_frame, width=30, show="*")
        self.password_entry.pack(pady=5)
        
        self.login_button = tk.Button(self.auth_frame, text="Entrar", command=self.login)
        self.login_button.pack(pady=5)
        
        self.register_button = tk.Button(self.auth_frame, text="Cadastrar", command=self.register_user)
        self.register_button.pack(pady=5)
        
        self.forgot_password_button = tk.Button(self.auth_frame, text="Esqueci a senha", command=self.forgot_password)
        self.forgot_password_button.pack(pady=5)
        
    def create_main_widgets(self):
        self.main_frame = tk.Frame(self)
        self.main_frame.pack(padx=20, pady=20, fill="both", expand=True)

        # Frame de Seleção de Pasta
        pasta_frame = tk.Frame(self.main_frame)
        pasta_frame.pack(fill="x", pady=(0, 10))
        tk.Label(pasta_frame, text="Selecione a pasta com os arquivos:").pack(side="left")
        tk.Button(pasta_frame, text="Selecionar Pasta", command=self.select_folder).pack(side="right")
        self.folder_path_label = tk.Label(self.main_frame, text="Nenhuma pasta selecionada", anchor="w")
        self.folder_path_label.pack(fill="x", padx=10)
        
        # Frame de Gerenciamento de Listas
        tk.Label(self.main_frame, text="Gerenciar Lista de Destinatários e Anexos:").pack(pady=(10, 5))
        list_controls_frame = tk.Frame(self.main_frame)
        list_controls_frame.pack(fill="x", pady=(0, 5))
        
        tk.Button(list_controls_frame, text="Adicionar Item", command=self.add_item).pack(side="left", padx=5)
        tk.Button(list_controls_frame, text="Remover Item", command=self.remove_item).pack(side="left", padx=5)
        tk.Button(list_controls_frame, text="Limpar Lista", command=self.clear_list).pack(side="left", padx=5)
        tk.Button(list_controls_frame, text="Editar Item", command=self.edit_item).pack(side="left", padx=5)
        tk.Button(list_controls_frame, text="Salvar Lista (DB)", command=self.save_list_db).pack(side="left", padx=5)
        tk.Button(list_controls_frame, text="Carregar Lista (DB)", command=self.load_list_db).pack(side="left", padx=5)
        
        # Novos botões de importar e exportar
        tk.Button(list_controls_frame, text="Exportar Lista (CSV)", command=self.export_list_csv).pack(side="left", padx=5)
        tk.Button(list_controls_frame, text="Importar Lista (CSV)", command=self.import_list_csv).pack(side="left", padx=5)

        # Treeview para a lista de envios
        self.tree = ttk.Treeview(self.main_frame, columns=("Destinatário", "Arquivos", "Assunto", "Corpo", "Agendamento"), show="headings", height=5)
        self.tree.heading("Destinatário", text="E-mail do Destinatário")
        self.tree.heading("Arquivos", text="Nome(s) dos Arquivos")
        self.tree.heading("Assunto", text="Assunto")
        self.tree.heading("Corpo", text="Corpo")
        self.tree.heading("Agendamento", text="Agendamento")
        self.tree.column("Destinatário", width=150, anchor=tk.W)
        self.tree.column("Arquivos", width=200, anchor=tk.W)
        self.tree.column("Assunto", width=200, anchor=tk.W)
        self.tree.column("Corpo", width=0, stretch=tk.NO)
        self.tree.column("Agendamento", width=120, anchor=tk.W)
        self.tree.pack(fill="both", expand=True)
        # Adiciona um evento de clique duplo para edição
        self.tree.bind("<Double-1>", self.on_double_click)

        # Campos de entrada
        input_frame = tk.Frame(self.main_frame)
        input_frame.pack(fill="x", pady=(10, 0))
        
        tk.Label(input_frame, text="E-mail:").pack(side="left", padx=5)
        self.destinatario_entry = tk.Entry(input_frame)
        self.destinatario_entry.pack(side="left", fill="x", expand=True, ipady=2)
        
        tk.Label(input_frame, text="Arquivos:").pack(side="left", padx=5)
        self.files_combobox = ttk.Combobox(input_frame, state="readonly")
        self.files_combobox.pack(side="left", fill="x", expand=True, ipady=2)
        
        self.add_file_button = tk.Button(input_frame, text="Adicionar Arquivo", command=self.add_selected_file)
        self.add_file_button.pack(side="left", padx=5)
        
        self.selected_files_label = tk.Label(self.main_frame, text="Arquivos selecionados para este item: Nenhum", anchor="w")
        self.selected_files_label.pack(fill="x", padx=10, pady=5)
        self.selected_files = []
        self.editing_item = None # Armazena o item que está sendo editado

        # Outros campos
        tk.Label(self.main_frame, text="Assunto do E-mail:").pack(pady=(10, 5))
        self.assunto_entry = tk.Entry(self.main_frame)
        self.assunto_entry.insert(0, "ENADE_2025_DISTRIBUIÇÃO - {{nomes_arquivos}}")
        self.assunto_entry.pack(fill="x", ipady=4)
        
        tk.Label(self.main_frame, text="E-mail(s) para Cópia (CC):").pack(pady=(10, 5))
        self.cc_entry = tk.Entry(self.main_frame)
        self.cc_entry.insert(0, ", ".join(copia_emails))
        self.cc_entry.pack(fill="x", ipady=4)

        # Frame para os templates de e-mail
        template_frame = tk.Frame(self.main_frame)
        template_frame.pack(fill="x", pady=(10, 5))
        tk.Label(template_frame, text="Template de E-mail:").pack(side="left", padx=(0, 5))
        self.template_combobox = ttk.Combobox(template_frame, state="readonly", values=list(self.templates.keys()))
        self.template_combobox.set("Padrão")
        self.template_combobox.bind("<<ComboboxSelected>>", self.load_template)
        self.template_combobox.pack(side="left", expand=True, fill="x")
        tk.Button(template_frame, text="Salvar Template", command=self.save_template).pack(side="left", padx=5)
        
        # Frame para o agendamento de envio com checkbox
        schedule_frame = tk.Frame(self.main_frame)
        schedule_frame.pack(fill="x", pady=(10, 5))
        self.agendar_checkbutton = tk.Checkbutton(schedule_frame, text="Agendar Envio:", variable=self.agendar_var)
        self.agendar_checkbutton.pack(side="left", padx=(0, 5))
        self.schedule_entry = tk.Entry(schedule_frame)
        self.schedule_entry.insert(0, "dd/mm/aaaa hh:mm")
        self.schedule_entry.pack(side="left", fill="x", expand=True)

        tk.Label(self.main_frame, text="Corpo do E-mail (HTML):").pack(pady=(10, 5))
        self.corpo_email_text = tk.Text(self.main_frame, height=15, width=70)
        self.corpo_email_text.pack(fill="both", expand=True, ipady=4)
        self.corpo_email_text.insert(tk.END, corpo_email_html_padrao)

        # Barra de Progresso
        self.progress_bar = ttk.Progressbar(self.main_frame, orient="horizontal", mode="determinate")
        self.progress_bar.pack(fill="x", pady=10)
        
        # Botão de Envio
        self.send_button = tk.Button(self.main_frame, text="Enviar E-mails", command=self.start_email_thread)
        self.send_button.pack(pady=20)

        # Log do Processo
        tk.Label(self.main_frame, text="Log do Processo:").pack(pady=(0, 5))
        self.log_text = tk.Text(self.main_frame, height=10, state="disabled")
        self.log_text.pack(fill="both", expand=True)
        self.log_text.tag_config('green', foreground='green')
        self.log_text.tag_config('red', foreground='red')

    # --- MÉTODOS DE AUTENTICAÇÃO ATUALIZADOS PARA PYREBASE ---

    def login(self):
        email = self.email_entry.get()
        password = self.password_entry.get()
        try:
            user = auth.sign_in_with_email_and_password(email, password)
            self.show_popup("Sucesso", f"Login bem-sucedido para {user['email']}")
            self.current_user_uid = user['localId']
            self.token = user['idToken']
            self.auth_frame.pack_forget()
            self.create_main_widgets()
        except Exception as e:
            self.show_popup("Erro de Login", f"Ocorreu um erro: {e}")

    def register_user(self):
        email = self.email_entry.get()
        password = self.password_entry.get()
        try:
            if not is_valid_email(email) or len(password) < 6:
                self.show_popup("Erro de Cadastro", "E-mail inválido ou senha muito curta (mínimo 6 caracteres).")
                return
            
            user = auth.create_user_with_email_and_password(email, password)
            self.show_popup("Sucesso", f"Usuário {email} criado. Você pode fazer o login agora.")
        except Exception as e:
            self.show_popup("Erro de Cadastro", f"Ocorreu um erro: {e}")

    def forgot_password(self):
        email = self.email_entry.get()
        try:
            auth.send_password_reset_email(email)
            self.show_popup("Esqueci a Senha", "Um link para redefinir a senha foi enviado para seu e-mail.")
        except Exception as e:
            self.show_popup("Erro", f"Ocorreu um erro: {e}")

    # --- MÉTODOS DE BANCO DE DADOS ATUALIZADOS PARA PYREBASE (REALTIME DATABASE) ---

    def load_template(self, event):
        selected_template = self.template_combobox.get()
        if selected_template:
            self.current_template = selected_template
            self.corpo_email_text.delete("1.0", tk.END)
            self.corpo_email_text.insert(tk.END, self.templates[self.current_template])

    def save_template(self):
        template_name = simpledialog.askstring("Salvar Template", "Nome do Template:")
        if template_name:
            self.templates[template_name] = self.corpo_email_text.get("1.0", tk.END).strip()
            self.template_combobox['values'] = list(self.templates.keys())
            self.template_combobox.set(template_name)
            self.current_template = template_name
            self.log(f"✅ Template '{template_name}' salvo com sucesso!", 'green')

    def clear_list(self):
        confirm = self.ask_yes_no("Confirmar Limpeza", "Tem certeza que deseja limpar toda a lista de envios?")
        if confirm:
            self.tree.delete(*self.tree.get_children())
            self.log("✅ Lista de envios limpa com sucesso.", 'green')
            self.progress_bar["value"] = 0
            self.editing_item = None
            self.add_file_button.config(text="Adicionar Arquivo")
            self.destinatario_entry.delete(0, tk.END)
            self.selected_files = []
            self.selected_files_label.config(text="Arquivos selecionados para este item: Nenhum")

    def on_double_click(self, event):
        selected_item = self.tree.selection()
        if not selected_item:
            return
        self.edit_item()

    def edit_item(self):
        selected_item = self.tree.selection()
        if not selected_item:
            self.show_popup("Atenção", "Nenhum item selecionado para editar.")
            return

        item_data = self.tree.item(selected_item)['values']
        
        self.destinatario_entry.delete(0, tk.END)
        self.destinatario_entry.insert(0, item_data[0])

        self.selected_files = item_data[1].split(', ')
        self.selected_files_label.config(text=f"Arquivos selecionados para este item: {', '.join(self.selected_files)}")

        self.assunto_entry.delete(0, tk.END)
        self.assunto_entry.insert(0, item_data[2])

        self.corpo_email_text.delete("1.0", tk.END)
        self.corpo_email_text.insert(tk.END, item_data[3])
        
        # Preenche o campo de agendamento se existir
        if len(item_data) > 4:
            agendamento = item_data[4]
            self.schedule_entry.delete(0, tk.END)
            if agendamento != "Não agendado":
                self.schedule_entry.insert(0, agendamento)
                self.agendar_var.set(True)
            else:
                self.schedule_entry.insert(0, "dd/mm/aaaa hh:mm")
                self.agendar_var.set(False)

        self.editing_item = selected_item[0]
        self.add_file_button.config(text="Atualizar Item")
        self.log(f"✏️ Modo de edição ativado para o item: {item_data[0]}", 'green')

    def add_item(self):
        destinatario_input = self.destinatario_entry.get().strip()
        anexo_nomes_str = ", ".join(self.selected_files)
        assunto_str = self.assunto_entry.get().strip()
        corpo_html_str = self.corpo_email_text.get("1.0", tk.END).strip()
        
        agendamento_str = self.schedule_entry.get().strip()
        agendamento_valido = False
        
        if self.agendar_var.get():
            if agendamento_str and agendamento_str.lower() != "dd/mm/aaaa hh:mm":
                try:
                    agendamento_dt = datetime.strptime(agendamento_str, "%d/%m/%Y %H:%M")
                    if agendamento_dt > datetime.now():
                        agendamento_valido = True
                    else:
                        self.show_popup("Erro de Agendamento", "A data e hora de agendamento devem ser no futuro.")
                        return
                except ValueError:
                    self.show_popup("Erro de Agendamento", "Formato de agendamento inválido. Use 'dd/mm/aaaa hh:mm'.")
                    return
            else:
                self.show_popup("Erro de Agendamento", "Marquei 'Agendar Envio', mas o campo está vazio.")
                return

        if not destinatario_input or not anexo_nomes_str:
            self.show_popup("Atenção", "Por favor, preencha o e-mail e selecione os arquivos para adicionar à lista.")
            return

        emails_list = [email.strip() for email in destinatario_input.split(',') if email.strip()]

        for nome in self.selected_files:
            if not self.anexos_map.get(normalizar_texto(nome)):
                self.show_popup("Erro de Validação", f"O arquivo '{nome}' não foi encontrado na pasta selecionada.")
                return

        for email in emails_list:
            if not is_valid_email(email):
                self.show_popup("Erro de Validação", f"O e-mail '{email}' não é válido.")
                return

        if self.editing_item:
            self.tree.delete(self.editing_item)
            self.editing_item = None
            self.add_file_button.config(text="Adicionar Item")
            self.log(f"✅ Item atualizado com sucesso.", 'green')
        
        for email in emails_list:
            self.tree.insert("", tk.END, values=(email, anexo_nomes_str, assunto_str, corpo_html_str, agendamento_str if agendamento_valido else "Não agendado"))
            
        self.destinatario_entry.delete(0, tk.END)
        self.selected_files = []
        self.selected_files_label.config(text="Arquivos selecionados para este item: Nenhum")
        self.schedule_entry.delete(0, tk.END)
        self.schedule_entry.insert(0, "dd/mm/aaaa hh:mm")
        self.agendar_var.set(False)

    def remove_item(self):
        selected_items = self.tree.selection()
        if not selected_items:
            self.show_popup("Atenção", "Nenhum item selecionado para remover.")
            return
        
        for item in selected_items:
            self.tree.delete(item)
            if item == self.editing_item:
                self.editing_item = None
                self.add_file_button.config(text="Adicionar Item")

    def save_list_db(self):
        if not self.current_user_uid or not self.token:
            self.show_popup("Erro", "Você precisa estar logado para salvar a lista.")
            return

        data = [self.tree.item(item, 'values') for item in self.tree.get_children()]
        
        try:
            db.child("lists").child(self.current_user_uid).set(data, self.token)
            self.show_popup("Sucesso", "Lista salva no banco de dados com sucesso!")
        except Exception as e:
            self.show_popup("Erro", f"Erro ao salvar a lista: {e}")

    def load_list_db(self):
        if not self.current_user_uid or not self.token:
            self.show_popup("Erro", "Você precisa estar logado para carregar a lista.")
            return

        try:
            list_data = db.child("lists").child(self.current_user_uid).get(self.token)
            
            if list_data and list_data.val():
                self.tree.delete(*self.tree.get_children())
                for item in list_data.val():
                    self.tree.insert("", tk.END, values=item)
                self.show_popup("Sucesso", "Última lista carregada com sucesso!")
            else:
                self.show_popup("Info", "Nenhuma lista encontrada para seu usuário.")
        except Exception as e:
            self.show_popup("Erro", f"Erro ao carregar a lista: {e}")

    def export_list_csv(self):
        file_path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
        )
        if not file_path:
            return

        with open(file_path, 'w', newline='', encoding='utf-8') as file:
            writer = csv.writer(file)
            writer.writerow(["Destinatário", "Arquivos", "Assunto", "Corpo", "Agendamento"])
            for item in self.tree.get_children():
                writer.writerow(self.tree.item(item)['values'])
        
        self.show_popup("Exportar", "Lista exportada com sucesso para CSV!")

    def import_list_csv(self):
        file_path = filedialog.askopenfilename(
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
        )
        if not file_path:
            return

        try:
            with open(file_path, 'r', newline='', encoding='utf-8') as file:
                reader = csv.reader(file)
                next(reader, None)
                
                self.tree.delete(*self.tree.get_children())
                for row in reader:
                    if len(row) == 5:
                        self.tree.insert("", tk.END, values=row)
            self.show_popup("Importar", "Lista importada com sucesso de CSV!")
        except Exception as e:
            self.show_popup("Erro", f"Erro ao importar o arquivo CSV: {e}")

    def select_folder(self):
        folder_path = filedialog.askdirectory()
        if folder_path:
            self.caminho_pasta = Path(folder_path)
            self.folder_path_label.config(text=str(self.caminho_pasta))
            self.log("📂 Pasta selecionada. Mapeando arquivos...", 'green')
            self.anexos_map = self.pre_process_anexos(self.caminho_pasta)
            
            nomes_arquivos = sorted(list(self.anexos_map.keys()))
            self.files_combobox['values'] = nomes_arquivos
            self.files_combobox.set('')
            
            if self.anexos_map:
                self.log(f"✅ {len(self.anexos_map)} arquivos encontrados e mapeados.", 'green')
            else:
                self.log(f"❌ Nenhum arquivo suportado encontrado na pasta.", 'red')

    def add_selected_file(self):
        selected_file = self.files_combobox.get()
        if selected_file and selected_file not in self.selected_files:
            self.selected_files.append(selected_file)
            self.selected_files_label.config(text=f"Arquivos selecionados para este item: {', '.join(self.selected_files)}")
            
    def add_item(self):
        destinatario_input = self.destinatario_entry.get().strip()
        anexo_nomes_str = ", ".join(self.selected_files)
        assunto_str = self.assunto_entry.get().strip()
        corpo_html_str = self.corpo_email_text.get("1.0", tk.END).strip()
        
        agendamento_str = self.schedule_entry.get().strip()
        agendamento_valido = False
        
        if self.agendar_var.get():
            if agendamento_str and agendamento_str.lower() != "dd/mm/aaaa hh:mm":
                try:
                    agendamento_dt = datetime.strptime(agendamento_str, "%d/%m/%Y %H:%M")
                    if agendamento_dt > datetime.now():
                        agendamento_valido = True
                    else:
                        self.show_popup("Erro de Agendamento", "A data e hora de agendamento devem ser no futuro.")
                        return
                except ValueError:
                    self.show_popup("Erro de Agendamento", "Formato de agendamento inválido. Use 'dd/mm/aaaa hh:mm'.")
                    return
            else:
                self.show_popup("Erro de Agendamento", "Marquei 'Agendar Envio', mas o campo está vazio.")
                return

        if not destinatario_input or not anexo_nomes_str:
            self.show_popup("Atenção", "Por favor, preencha o e-mail e selecione os arquivos para adicionar à lista.")
            return

        emails_list = [email.strip() for email in destinatario_input.split(',') if email.strip()]

        for nome in self.selected_files:
            if not self.anexos_map.get(normalizar_texto(nome)):
                self.show_popup("Erro de Validação", f"O arquivo '{nome}' não foi encontrado na pasta selecionada.")
                return

        for email in emails_list:
            if not is_valid_email(email):
                self.show_popup("Erro de Validação", f"O e-mail '{email}' não é válido.")
                return

        if self.editing_item:
            self.tree.delete(self.editing_item)
            self.editing_item = None
            self.add_file_button.config(text="Adicionar Item")
            self.log(f"✅ Item atualizado com sucesso.", 'green')
        
        for email in emails_list:
            self.tree.insert("", tk.END, values=(email, anexo_nomes_str, assunto_str, corpo_html_str, agendamento_str if agendamento_valido else "Não agendado"))
            
        self.destinatario_entry.delete(0, tk.END)
        self.selected_files = []
        self.selected_files_label.config(text="Arquivos selecionados para este item: Nenhum")
        self.schedule_entry.delete(0, tk.END)
        self.schedule_entry.insert(0, "dd/mm/aaaa hh:mm")
        self.agendar_var.set(False)

    def remove_item(self):
        selected_items = self.tree.selection()
        if not selected_items:
            self.show_popup("Atenção", "Nenhum item selecionado para remover.")
            return
        
        for item in selected_items:
            self.tree.delete(item)
            if item == self.editing_item:
                self.editing_item = None
                self.add_file_button.config(text="Adicionar Item")

    def save_list_db(self):
        if not self.current_user_uid or not self.token:
            self.show_popup("Erro", "Você precisa estar logado para salvar a lista.")
            return

        data = [self.tree.item(item, 'values') for item in self.tree.get_children()]
        
        try:
            db.child("lists").child(self.current_user_uid).set(data, self.token)
            self.show_popup("Sucesso", "Lista salva no banco de dados com sucesso!")
        except Exception as e:
            self.show_popup("Erro", f"Erro ao salvar a lista: {e}")

    def load_list_db(self):
        if not self.current_user_uid or not self.token:
            self.show_popup("Erro", "Você precisa estar logado para carregar a lista.")
            return

        try:
            list_data = db.child("lists").child(self.current_user_uid).get(self.token)
            
            if list_data and list_data.val():
                self.tree.delete(*self.tree.get_children())
                for item in list_data.val():
                    self.tree.insert("", tk.END, values=item)
                self.show_popup("Sucesso", "Última lista carregada com sucesso!")
            else:
                self.show_popup("Info", "Nenhuma lista encontrada para seu usuário.")
        except Exception as e:
            self.show_popup("Erro", f"Erro ao carregar a lista: {e}")

    def export_list_csv(self):
        file_path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
        )
        if not file_path:
            return

        with open(file_path, 'w', newline='', encoding='utf-8') as file:
            writer = csv.writer(file)
            writer.writerow(["Destinatário", "Arquivos", "Assunto", "Corpo", "Agendamento"])
            for item in self.tree.get_children():
                writer.writerow(self.tree.item(item)['values'])
        
        self.show_popup("Exportar", "Lista exportada com sucesso para CSV!")

    def import_list_csv(self):
        file_path = filedialog.askopenfilename(
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
        )
        if not file_path:
            return

        try:
            with open(file_path, 'r', newline='', encoding='utf-8') as file:
                reader = csv.reader(file)
                next(reader, None)
                
                self.tree.delete(*self.tree.get_children())
                for row in reader:
                    if len(row) == 5:
                        self.tree.insert("", tk.END, values=row)
            self.show_popup("Importar", "Lista importada com sucesso de CSV!")
        except Exception as e:
            self.show_popup("Erro", f"Erro ao importar o arquivo CSV: {e}")

    def select_folder(self):
        folder_path = filedialog.askdirectory()
        if folder_path:
            self.caminho_pasta = Path(folder_path)
            self.folder_path_label.config(text=str(self.caminho_pasta))
            self.log("📂 Pasta selecionada. Mapeando arquivos...", 'green')
            self.anexos_map = self.pre_process_anexos(self.caminho_pasta)
            
            nomes_arquivos = sorted(list(self.anexos_map.keys()))
            self.files_combobox['values'] = nomes_arquivos
            self.files_combobox.set('')
            
            if self.anexos_map:
                self.log(f"✅ {len(self.anexos_map)} arquivos encontrados e mapeados.", 'green')
            else:
                self.log(f"❌ Nenhum arquivo suportado encontrado na pasta.", 'red')

    def add_selected_file(self):
        selected_file = self.files_combobox.get()
        if selected_file and selected_file not in self.selected_files:
            self.selected_files.append(selected_file)
            self.selected_files_label.config(text=f"Arquivos selecionados para este item: {', '.join(self.selected_files)}")
            
    def add_item(self):
        destinatario_input = self.destinatario_entry.get().strip()
        anexo_nomes_str = ", ".join(self.selected_files)
        assunto_str = self.assunto_entry.get().strip()
        corpo_html_str = self.corpo_email_text.get("1.0", tk.END).strip()
        
        agendamento_str = self.schedule_entry.get().strip()
        agendamento_valido = False
        
        if self.agendar_var.get():
            if agendamento_str and agendamento_str.lower() != "dd/mm/aaaa hh:mm":
                try:
                    agendamento_dt = datetime.strptime(agendamento_str, "%d/%m/%Y %H:%M")
                    if agendamento_dt > datetime.now():
                        agendamento_valido = True
                    else:
                        self.show_popup("Erro de Agendamento", "A data e hora de agendamento devem ser no futuro.")
                        return
                except ValueError:
                    self.show_popup("Erro de Agendamento", "Formato de agendamento inválido. Use 'dd/mm/aaaa hh:mm'.")
                    return
            else:
                self.show_popup("Erro de Agendamento", "Marquei 'Agendar Envio', mas o campo está vazio.")
                return

        if not destinatario_input or not anexo_nomes_str:
            self.show_popup("Atenção", "Por favor, preencha o e-mail e selecione os arquivos para adicionar à lista.")
            return

        emails_list = [email.strip() for email in destinatario_input.split(',') if email.strip()]

        for nome in self.selected_files:
            if not self.anexos_map.get(normalizar_texto(nome)):
                self.show_popup("Erro de Validação", f"O arquivo '{nome}' não foi encontrado na pasta selecionada.")
                return

        for email in emails_list:
            if not is_valid_email(email):
                self.show_popup("Erro de Validação", f"O e-mail '{email}' não é válido.")
                return

        # Se estiver em modo de edição, remove o item antigo e insere o novo
        if self.editing_item:
            self.tree.delete(self.editing_item)
            self.editing_item = None
            self.add_file_button.config(text="Adicionar Item")
            self.log(f"✅ Item atualizado com sucesso.", 'green')
        
        for email in emails_list:
            self.tree.insert("", tk.END, values=(email, anexo_nomes_str, assunto_str, corpo_html_str, agendamento_str if agendamento_valido else "Não agendado"))
            
        self.destinatario_entry.delete(0, tk.END)
        self.selected_files = []
        self.selected_files_label.config(text="Arquivos selecionados para este item: Nenhum")
        self.schedule_entry.delete(0, tk.END)
        self.schedule_entry.insert(0, "dd/mm/aaaa hh:mm")
        self.agendar_var.set(False)

    def remove_item(self):
        selected_items = self.tree.selection()
        if not selected_items:
            self.show_popup("Atenção", "Nenhum item selecionado para remover.")
            return
        
        for item in selected_items:
            self.tree.delete(item)
            if item == self.editing_item:
                self.editing_item = None
                self.add_file_button.config(text="Adicionar Item")

    def save_list_db(self):
        if not self.current_user_uid or not self.token:
            self.show_popup("Erro", "Você precisa estar logado para salvar a lista.")
            return

        data = [self.tree.item(item, 'values') for item in self.tree.get_children()]
        
        try:
            db.child("lists").child(self.current_user_uid).set(data, self.token)
            self.show_popup("Sucesso", "Lista salva no banco de dados com sucesso!")
        except Exception as e:
            self.show_popup("Erro", f"Erro ao salvar a lista: {e}")

    def load_list_db(self):
        if not self.current_user_uid or not self.token:
            self.show_popup("Erro", "Você precisa estar logado para carregar a lista.")
            return

        try:
            list_data = db.child("lists").child(self.current_user_uid).get(self.token)
            
            if list_data and list_data.val():
                self.tree.delete(*self.tree.get_children())
                for item in list_data.val():
                    self.tree.insert("", tk.END, values=item)
                self.show_popup("Sucesso", "Última lista carregada com sucesso!")
            else:
                self.show_popup("Info", "Nenhuma lista encontrada para seu usuário.")
        except Exception as e:
            self.show_popup("Erro", f"Erro ao carregar a lista: {e}")

    def export_list_csv(self):
        file_path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
        )
        if not file_path:
            return

        with open(file_path, 'w', newline='', encoding='utf-8') as file:
            writer = csv.writer(file)
            writer.writerow(["Destinatário", "Arquivos", "Assunto", "Corpo", "Agendamento"])
            for item in self.tree.get_children():
                writer.writerow(self.tree.item(item)['values'])
        
        self.show_popup("Exportar", "Lista exportada com sucesso para CSV!")

    def import_list_csv(self):
        file_path = filedialog.askopenfilename(
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
        )
        if not file_path:
            return

        try:
            with open(file_path, 'r', newline='', encoding='utf-8') as file:
                reader = csv.reader(file)
                next(reader, None)
                
                self.tree.delete(*self.tree.get_children())
                for row in reader:
                    if len(row) == 5:
                        self.tree.insert("", tk.END, values=row)
            self.show_popup("Importar", "Lista importada com sucesso de CSV!")
        except Exception as e:
            self.show_popup("Erro", f"Erro ao importar o arquivo CSV: {e}")

    def select_folder(self):
        folder_path = filedialog.askdirectory()
        if folder_path:
            self.caminho_pasta = Path(folder_path)
            self.folder_path_label.config(text=str(self.caminho_pasta))
            self.log("📂 Pasta selecionada. Mapeando arquivos...", 'green')
            self.anexos_map = self.pre_process_anexos(self.caminho_pasta)
            
            nomes_arquivos = sorted(list(self.anexos_map.keys()))
            self.files_combobox['values'] = nomes_arquivos
            self.files_combobox.set('')
            
            if self.anexos_map:
                self.log(f"✅ {len(self.anexos_map)} arquivos encontrados e mapeados.", 'green')
            else:
                self.log(f"❌ Nenhum arquivo suportado encontrado na pasta.", 'red')

    def add_selected_file(self):
        selected_file = self.files_combobox.get()
        if selected_file and selected_file not in self.selected_files:
            self.selected_files.append(selected_file)
            self.selected_files_label.config(text=f"Arquivos selecionados para este item: {', '.join(self.selected_files)}")
            
    def add_item(self):
        destinatario_input = self.destinatario_entry.get().strip()
        anexo_nomes_str = ", ".join(self.selected_files)
        assunto_str = self.assunto_entry.get().strip()
        corpo_html_str = self.corpo_email_text.get("1.0", tk.END).strip()
        
        agendamento_str = self.schedule_entry.get().strip()
        agendamento_valido = False
        
        if self.agendar_var.get():
            if agendamento_str and agendamento_str.lower() != "dd/mm/aaaa hh:mm":
                try:
                    agendamento_dt = datetime.strptime(agendamento_str, "%d/%m/%Y %H:%M")
                    if agendamento_dt > datetime.now():
                        agendamento_valido = True
                    else:
                        self.show_popup("Erro de Agendamento", "A data e hora de agendamento devem ser no futuro.")
                        return
                except ValueError:
                    self.show_popup("Erro de Agendamento", "Formato de agendamento inválido. Use 'dd/mm/aaaa hh:mm'.")
                    return
            else:
                self.show_popup("Erro de Agendamento", "Marquei 'Agendar Envio', mas o campo está vazio.")
                return

        if not destinatario_input or not anexo_nomes_str:
            self.show_popup("Atenção", "Por favor, preencha o e-mail e selecione os arquivos para adicionar à lista.")
            return

        emails_list = [email.strip() for email in destinatario_input.split(',') if email.strip()]

        for nome in self.selected_files:
            if not self.anexos_map.get(normalizar_texto(nome)):
                self.show_popup("Erro de Validação", f"O arquivo '{nome}' não foi encontrado na pasta selecionada.")
                return

        for email in emails_list:
            if not is_valid_email(email):
                self.show_popup("Erro de Validação", f"O e-mail '{email}' não é válido.")
                return

        if self.editing_item:
            self.tree.delete(self.editing_item)
            self.editing_item = None
            self.add_file_button.config(text="Adicionar Item")
            self.log(f"✅ Item atualizado com sucesso.", 'green')
        
        for email in emails_list:
            self.tree.insert("", tk.END, values=(email, anexo_nomes_str, assunto_str, corpo_html_str, agendamento_str if agendamento_valido else "Não agendado"))
            
        self.destinatario_entry.delete(0, tk.END)
        self.selected_files = []
        self.selected_files_label.config(text="Arquivos selecionados para este item: Nenhum")
        self.schedule_entry.delete(0, tk.END)
        self.schedule_entry.insert(0, "dd/mm/aaaa hh:mm")
        self.agendar_var.set(False)

    def remove_item(self):
        selected_items = self.tree.selection()
        if not selected_items:
            self.show_popup("Atenção", "Nenhum item selecionado para remover.")
            return
        
        for item in selected_items:
            self.tree.delete(item)
            if item == self.editing_item:
                self.editing_item = None
                self.add_file_button.config(text="Adicionar Item")

    def save_list_db(self):
        if not self.current_user_uid or not self.token:
            self.show_popup("Erro", "Você precisa estar logado para salvar a lista.")
            return

        data = [self.tree.item(item, 'values') for item in self.tree.get_children()]
        
        try:
            db.child("lists").child(self.current_user_uid).set(data, self.token)
            self.show_popup("Sucesso", "Lista salva no banco de dados com sucesso!")
        except Exception as e:
            self.show_popup("Erro", f"Erro ao salvar a lista: {e}")

    def load_list_db(self):
        if not self.current_user_uid or not self.token:
            self.show_popup("Erro", "Você precisa estar logado para carregar a lista.")
            return

        try:
            list_data = db.child("lists").child(self.current_user_uid).get(self.token)
            
            if list_data and list_data.val():
                self.tree.delete(*self.tree.get_children())
                for item in list_data.val():
                    self.tree.insert("", tk.END, values=item)
                self.show_popup("Sucesso", "Última lista carregada com sucesso!")
            else:
                self.show_popup("Info", "Nenhuma lista encontrada para seu usuário.")
        except Exception as e:
            self.show_popup("Erro", f"Erro ao carregar a lista: {e}")

    def export_list_csv(self):
        file_path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
        )
        if not file_path:
            return

        with open(file_path, 'w', newline='', encoding='utf-8') as file:
            writer = csv.writer(file)
            writer.writerow(["Destinatário", "Arquivos", "Assunto", "Corpo", "Agendamento"])
            for item in self.tree.get_children():
                writer.writerow(self.tree.item(item)['values'])
        
        self.show_popup("Exportar", "Lista exportada com sucesso para CSV!")

    def import_list_csv(self):
        file_path = filedialog.askopenfilename(
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
        )
        if not file_path:
            return

        try:
            with open(file_path, 'r', newline='', encoding='utf-8') as file:
                reader = csv.reader(file)
                next(reader, None)
                
                self.tree.delete(*self.tree.get_children())
                for row in reader:
                    if len(row) == 5:
                        self.tree.insert("", tk.END, values=row)
            self.show_popup("Importar", "Lista importada com sucesso de CSV!")
        except Exception as e:
            self.show_popup("Erro", f"Erro ao importar o arquivo CSV: {e}")

    def select_folder(self):
        folder_path = filedialog.askdirectory()
        if folder_path:
            self.caminho_pasta = Path(folder_path)
            self.folder_path_label.config(text=str(self.caminho_pasta))
            self.log("📂 Pasta selecionada. Mapeando arquivos...", 'green')
            self.anexos_map = self.pre_process_anexos(self.caminho_pasta)
            
            nomes_arquivos = sorted(list(self.anexos_map.keys()))
            self.files_combobox['values'] = nomes_arquivos
            self.files_combobox.set('')
            
            if self.anexos_map:
                self.log(f"✅ {len(self.anexos_map)} arquivos encontrados e mapeados.", 'green')
            else:
                self.log(f"❌ Nenhum arquivo suportado encontrado na pasta.", 'red')

    def add_selected_file(self):
        selected_file = self.files_combobox.get()
        if selected_file and selected_file not in self.selected_files:
            self.selected_files.append(selected_file)
            self.selected_files_label.config(text=f"Arquivos selecionados para este item: {', '.join(self.selected_files)}")
            
    def add_item(self):
        destinatario_input = self.destinatario_entry.get().strip()
        anexo_nomes_str = ", ".join(self.selected_files)
        assunto_str = self.assunto_entry.get().strip()
        corpo_html_str = self.corpo_email_text.get("1.0", tk.END).strip()
        
        if not destinatario_input or not anexo_nomes_str:
            self.show_popup("Atenção", "Por favor, preencha o e-mail e selecione os arquivos para adicionar à lista.")
            return

        emails_list = [email.strip() for email in destinatario_input.split(',') if email.strip()]

        for nome in self.selected_files:
            if not self.anexos_map.get(normalizar_texto(nome)):
                self.show_popup("Erro de Validação", f"O arquivo '{nome}' não foi encontrado na pasta selecionada.")
                return

        for email in emails_list:
            if not is_valid_email(email):
                self.show_popup("Erro de Validação", f"O e-mail '{email}' não é válido.")
                return

        # Se estiver em modo de edição, remove o item antigo e insere o novo
        if self.editing_item:
            self.tree.delete(self.editing_item)
            self.editing_item = None
            self.add_file_button.config(text="Adicionar Item")
            self.log(f"✅ Item atualizado com sucesso.", 'green')
        
        for email in emails_list:
            self.tree.insert("", tk.END, values=(email, anexo_nomes_str, assunto_str, corpo_html_str))
            
        self.destinatario_entry.delete(0, tk.END)
        self.selected_files = []
        self.selected_files_label.config(text="Arquivos selecionados para este item: Nenhum")
    
    def remove_item(self):
        selected_items = self.tree.selection()
        if not selected_items:
            self.show_popup("Atenção", "Nenhum item selecionado para remover.")
            return
        
        for item in selected_items:
            self.tree.delete(item)

    def log(self, message, tag=None):
        self.log_text.config(state="normal")
        self.log_text.insert(tk.END, f"{message}\n", tag)
        self.log_text.see(tk.END)
        self.log_text.config(state="disabled")

    def pre_process_anexos(self, caminho_pasta):
        anexos_map = {}
        try:
            arquivos_na_pasta = list(caminho_pasta.glob("*"))
            
            for arquivo in arquivos_na_pasta:
                if arquivo.is_file() and arquivo.suffix.lower() in ('.pdf', '.xlsx', '.docx', '.jpg', '.jpeg', '.png'):
                    nome_base = arquivo.name
                    anexos_map[normalizar_texto(nome_base)] = str(arquivo)
        except Exception as e:
            self.log(f"{datetime.now().strftime('%H:%M:%S')} ❌ Erro ao processar anexos: {e}", 'red')
        return anexos_map

    def start_email_thread(self):
        self.progress_bar["value"] = 0
        self.send_button["state"] = "disabled"
        threading.Thread(target=self.send_multiple_emails).start()

    def send_multiple_emails(self):
        cc_list = [email.strip() for email in self.cc_entry.get().strip().split(',') if email.strip()]

        lista_envios = [self.tree.item(item, 'values') for item in self.tree.get_children()]
        
        if not lista_envios or not self.caminho_pasta:
            self.log(f"{datetime.now().strftime('%H:%M:%S')} ❌ Por favor, adicione itens à lista e selecione uma pasta.", 'red')
            self.show_popup("Atenção", "Por favor, adicione itens à lista e selecione uma pasta.")
            self.send_button["state"] = "normal"
            return

        # Validação de e-mails antes de iniciar o processo de envio
        for item_data in lista_envios:
            destinatario = item_data[0]
            if not is_valid_email(destinatario):
                self.log(f"{datetime.now().strftime('%H:%M:%S')} ❌ Erro de Validação: O e-mail '{destinatario}' na lista não é válido. Processo de envio cancelado.", 'red')
                self.show_popup("Erro de Validação", f"O e-mail '{destinatario}' na lista não é válido. O processo de envio foi cancelado.")
                self.send_button["state"] = "normal"
                return

        # Lógica de agendamento
        agendamento_total = None
        for item_data in lista_envios:
            if len(item_data) > 4 and item_data[4] != "Não agendado":
                try:
                    agendamento_total = datetime.strptime(item_data[4], "%d/%m/%Y %H:%M")
                    break
                except ValueError:
                    pass

        if agendamento_total:
            self.log(f"{datetime.now().strftime('%H:%M:%S')} ⏰ Agendamento detectado. Envio programado para {agendamento_total.strftime('%d/%m/%Y %H:%M')}.", 'green')
            while datetime.now() < agendamento_total:
                time.sleep(1)
            self.log(f"{datetime.now().strftime('%H:%M:%S')} 🚀 Horário do agendamento atingido. Iniciando envio...", 'green')
        else:
            self.log(f"{datetime.now().strftime('%H:%M:%S')} 🚀 Iniciando envio de e-mails em lote.", 'green')
            time.sleep(3)
        
        self.progress_bar["maximum"] = len(lista_envios)
        total_sent = 0
        
        # O loop percorre toda a lista de envios, sem limite de quantidade.
        for i, item_data in enumerate(lista_envios):
            destinatario, nomes_base_arquivos_str, assunto, corpo_email_html, agendamento_status = item_data
            nomes_base_arquivos = [nome.strip() for nome in nomes_base_arquivos_str.split(',') if nome.strip()]
            
            anexos_paths = []
            for nome in nomes_base_arquivos:
                caminho = self.anexos_map.get(normalizar_texto(nome))
                if caminho:
                    anexos_paths.append(caminho)
            
            if not anexos_paths:
                self.log(f"{datetime.now().strftime('%H:%M:%S')} ❌ Erro: Nenhum arquivo válido encontrado para '{nomes_base_arquivos_str}'. Pulando este e-mail.", 'red')
                continue

            corpo_email_final = corpo_email_html.replace("{{nomes_arquivos}}", nomes_base_arquivos_str).replace("{{assunto}}", assunto)
            
            self.log(f"{datetime.now().strftime('%H:%M:%S')} --- Processando E-mail {i+1} de {len(lista_envios)} ---", 'green')
            self.log(f"{datetime.now().strftime('%H:%M:%S')} Destinatário: {destinatario}", 'green')
            self.log(f"{datetime.now().strftime('%H:%M:%S')} Assunto: {assunto}", 'green')
            self.log(f"{datetime.now().strftime('%H:%M:%S')} Arquivos: {', '.join([Path(p).name for p in anexos_paths])}", 'green')

            success = enviar_email_com_anexos(
                destinatario=destinatario,
                assunto=assunto,
                corpo_html=corpo_email_final,
                anexos_paths=anexos_paths,
                cc_list=cc_list,
                log_text=self.log_text
            )

            if success:
                total_sent += 1
            
            self.progress_bar["value"] = i + 1
            time.sleep(self.send_interval)

        self.log(f"\n{datetime.now().strftime('%H:%M:%S')} ✅ Processo de envio finalizado. {total_sent} e-mails enviados com sucesso.", 'green')
        
        if total_sent > 0:
            self.show_popup("Envio Concluído", f"O processo de envio foi finalizado.\n{total_sent} e-mail(s) enviados com sucesso.")
        else:
            self.show_popup("Envio Falhou", f"Nenhum e-mail foi enviado com sucesso.")
        
        self.send_button["state"] = "normal"
    
    # --- MÉTODOS AUXILIARES PARA POP-UPS CENTRALIZADOS ---
    def show_popup(self, title, message):
        """Exibe um messagebox centralizado."""
        # Cria uma janela temporária transparente
        root = tk.Tk()
        root.withdraw()

        # Calcula a posição para centralizar o pop-up
        x = self.winfo_x() + (self.winfo_width() - 300) // 2
        y = self.winfo_y() + (self.winfo_height() - 150) // 2
        root.geometry(f'+{x}+{y}')

        messagebox.showinfo(title, message, master=root)
        root.destroy()
    
    def ask_yes_no(self, title, message):
        """Exibe um messagebox de confirmação centralizado."""
        root = tk.Tk()
        root.withdraw()
        x = self.winfo_x() + (self.winfo_width() - 300) // 2
        y = self.winfo_y() + (self.winfo_height() - 150) // 2
        root.geometry(f'+{x}+{y}')
        result = messagebox.askyesno(title, message, master=root)
        root.destroy()
        return result

if __name__ == "__main__":
    app = EmailApp()
    app.mainloop()