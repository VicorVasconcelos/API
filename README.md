# Sistema de Automação de E-mails

## 📋 Descrição

Sistema desenvolvido em Python para automatizar o envio de e-mails em massa com anexos múltiplos. Inclui autenticação Firebase, interface gráfica moderna, agendamento de envios e gerenciamento de listas de destinatários.

## 🚀 Principais Funcionalidades

- **🔐 Autenticação Firebase** - Login/cadastro seguro com recuperação de senha
- **📧 Múltiplos anexos** por destinatário com templates HTML personalizáveis
- **⏰ Agendamento** de envios com data/hora específica
- **📊 Gerenciamento de listas** - CRUD completo com exportação/importação CSV
- **☁️ Salvamento na nuvem** - Sincronização automática via Firebase
- **🎯 Interface moderna** - tkinter/ttk com modo de edição e log em tempo real

## 🛠️ Tecnologias

- **Python 3.x** - Linguagem principal
- **Firebase** - Autenticação e banco de dados (Pyrebase)
- **tkinter/ttk** - Interface gráfica moderna
- **pywin32** - Integração com Microsoft Outlook

## 📦 Requisitos e Instalação

### Pré-requisitos
- Windows + Microsoft Outlook configurado
- Python 3.6+ e conexão com internet

### Instalação
```bash
git clone https://github.com/VicorVasconcelos/API.git
cd API
pip install pyrebase4 pywin32
python "# --- IMPORTS ---.py"
```

## ⚙️ Configuração Básica

Configure suas credenciais Firebase no código e edite a lista de e-mails de cópia:

```python
# Firebase Config
firebaseConfig = { /* suas credenciais */ }

# E-mails de cópia
copia_emails = ["email1@exemplo.com", "email2@exemplo.com"]
```

**Formatos suportados**: PDF, XLSX, DOCX, JPG, PNG

## 🎯 Como Usar

1. **Login**: Cadastre-se ou faça login na tela inicial
2. **Pasta**: Selecione a pasta com os arquivos
3. **Destinatários**: Adicione e-mails e selecione arquivos específicos
4. **Configuração**: Ajuste assunto, template e agendamento se necessário
5. **Envio**: Clique em "Enviar E-mails" e acompanhe o progresso

**Funcionalidades extras**: Templates salvos, exportação CSV, edição com clique duplo

## � Versão: v9.3

**Principais novidades:**
- ✅ Autenticação Firebase + múltiplos anexos
- ✅ Agendamento + templates personalizáveis  
- ✅ Exportação CSV + interface moderna

## 🔧 Problemas Comuns

- **Firebase**: Verifique credenciais e conexão
- **Outlook**: Execute como administrador se necessário
- **Arquivos**: Confirme formatos (.pdf, .xlsx, .docx, .jpg, .png)
- **Agendamento**: Use formato dd/mm/aaaa hh:mm

## 👨‍💻 Autores

**Victor Vasconcelos** e **Samuel Almeida**
- GitHub: [@VicorVasconcelos](https://github.com/VicorVasconcelos) | [@oak-samc](https://github.com/oak-samc)
- Issues: [GitHub Issues](https://github.com/VicorVasconcelos/API/issues)

---
**⚠️ Nota**: Requer Windows + Outlook + internet para funcionalidades completas.
