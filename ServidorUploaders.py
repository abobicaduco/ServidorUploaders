import os
import sys
import socket
import random
import string
import subprocess
import logging
from datetime import datetime, timedelta
from pathlib import Path

# --- AUTO-INSTALLER DE DEPENDÊNCIAS ---
def check_and_install_dependencies():
    required_packages = {
        'flask': 'flask',
        'pandas': 'pandas',
        'win32com': 'pywin32',
        'werkzeug': 'werkzeug',
        'openpyxl': 'openpyxl',  # Necessário para ler .xlsx no pandas
        'pytz': 'pytz',
        'dotenv': 'python-dotenv'
    }
    
    # Não tenta instalar se estiver rodando como executável compilado (.exe)
    if getattr(sys, 'frozen', False):
        return
        
    for module_name, install_name in required_packages.items():
        try:
            __import__(module_name)
        except ImportError:
            print(f"[*] Instalando dependência requerida: {install_name}...")
            # Pip install
            subprocess.check_call([sys.executable, "-m", "pip", "install", install_name])

check_and_install_dependencies()

import pandas as pd
import pytz
from flask import Flask, request, render_template, session, redirect, url_for, flash, abort, jsonify, Response
from werkzeug.utils import secure_filename
from dotenv import load_dotenv

# Carrega variáveis de ambiente do arquivo .env, se existir
load_dotenv()

try:
    import win32com.client as win32
    import pythoncom
    WIN32_AVAILABLE = True
except ImportError:
    WIN32_AVAILABLE = False

# --- CORREÇÃO DOS CAMINHOS PARA SUPORTAR .EXE (PYINSTALLER) ---
if getattr(sys, 'frozen', False):
    # Se estiver rodando como um .exe compilado, o _MEIPASS é onde os arquivos do app estão
    BUNDLE_DIR = sys._MEIPASS
    APPLICATION_PATH = os.path.dirname(sys.executable)
else:
    # Se estiver rodando como um script .py normal
    BUNDLE_DIR = os.path.dirname(os.path.abspath(__file__))
    APPLICATION_PATH = BUNDLE_DIR

app = Flask(__name__, template_folder=BUNDLE_DIR)
# 12-Factor: Configurações via variáveis de ambiente com fallback
app.secret_key = os.environ.get('FLASK_SECRET_KEY', 'chave_super_secreta_para_sessao_c6_monitoracao_fg78')

# Configurar Sessão para expirar em 24h
app.config['PERMANENT_SESSION_LIFETIME'] = timedelta(hours=24)

# --- CONFIGURAÇÕES DE SEGURANÇA E HEADERS ---
@app.after_request
def add_security_headers(response):
    response.headers['X-Content-Type-Options'] = 'nosniff'
    response.headers['X-Frame-Options'] = 'SAMEORIGIN'
    response.headers['X-XSS-Protection'] = '1; mode=block'
    response.headers['Strict-Transport-Security'] = 'max-age=31536000; includeSubDomains'
    return response

# --- CONFIGURAÇÕES DO SISTEMA E LOGS ---
# Para adicionar mais administradores, separar os nomes por vírgula no .env: ADMIN_USERS=carlos.lsilva,joao.silva
admin_env = os.environ.get('ADMIN_USERS', 'carlos.lsilva')
ADMIN_USERS = [u.strip() for u in admin_env.split(',')]

HOME = Path.home()
# Customização pelo ambiente (.env) para conseguir emular fora da rede da Empresa
ENV_PATH_CELULA = os.environ.get('PATH_CELULA')
if ENV_PATH_CELULA:
    PATH_CELULA = Path(ENV_PATH_CELULA)
else:
    PATH_CELULA = HOME / "C6 CTVM LTDA, BANCO C6 S.A. e C6 HOLDING S.A" / "Mensageria e Cargas Operacionais - 11.CelulaPython"

ENV_BASE_PATH = os.environ.get('BASE_PATH')
if ENV_BASE_PATH:
    BASE_PATH = ENV_BASE_PATH
else:
    # Utilizando PATH_CELULA para localizar o diretório dinamicamente
    BASE_PATH = str(PATH_CELULA / "graciliano" / "automacoes")

TIMEZONE = pytz.timezone("America/Sao_Paulo")
START_TIME = datetime.now(TIMEZONE).replace(microsecond=0)
try:
    SCRIPT_NAME = Path(__file__).stem.lower()
except NameError:
    SCRIPT_NAME = "servidor"
SCRIPT_DEPS_NAME = SCRIPT_NAME.upper() # Para logs e email

def setup_logger(area_name="servidor"):
    # Pasta raiz de logs
    log_dir = PATH_CELULA / "graciliano" / "automacoes" / area_name / "logs" / SCRIPT_NAME / START_TIME.strftime('%Y-%m-%d')
    TEMP_DIR = Path("C:/TEMP/logs_orquestrador")
    
    # Fallback para TEMP se rede indisponível
    try:
        log_dir.mkdir(parents=True, exist_ok=True)
    except Exception:
        log_dir = TEMP_DIR
        log_dir.mkdir(parents=True, exist_ok=True)
        
    log_file = log_dir / f"{SCRIPT_NAME}_{datetime.now().strftime('%H%M%S')}.log"
    
    logger = logging.getLogger(SCRIPT_NAME)
    logger.setLevel(logging.DEBUG)
    logger.propagate = False
    
    if not logger.handlers:
        formatter = logging.Formatter('%(asctime)s | %(levelname)s | %(message)s', datefmt='%Y-%m-%d %H:%M:%S')
        
        # File Handler
        fh = logging.FileHandler(log_file, encoding='utf-8')
        fh.setFormatter(formatter)
        logger.addHandler(fh)
        
        # Stream Handler
        ch = logging.StreamHandler(sys.stdout)
        ch.setFormatter(formatter)
        logger.addHandler(ch)
        
    return logger, log_dir

logger, log_dir = setup_logger()

# O Excel não deve ser embutido no .exe se quiser editá-lo. 
# Por isso, ele fica na mesma pasta do servidor.py ou do servidor.exe.
# Permite override via .env para uso de emulação (EXCEL_FILE=dummy.xlsx)
EXCEL_FILE = os.environ.get('EXCEL_FILE', os.path.join(APPLICATION_PATH, "UPLOADERS.xlsx"))
DOMAIN = "@c6bank.com"

auth_tokens = {}

def get_ip():
    s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
    s.settimeout(0)
    try:
        s.connect(('10.254.254.254', 1))
        ip = s.getsockname()[0]
    except Exception:
        ip = '127.0.0.1'
    finally:
        s.close()
    return ip

def enviar_email_outlook(destinatario, token):
    # Condição para desenvolvimento pessoal sem acesso ao Outlook
    if os.environ.get('MOCK_EMAIL', 'False').lower() == 'true':
        logger.info(f"[MOCK EMAIL] Simulando envio. O Token de Acesso para {destinatario} é: {token}")
        return True
        
    if not WIN32_AVAILABLE:
        logger.error("Erro: win32com não está instalado ou disponível. Não é possível enviar email.")
        return False

    pythoncom.CoInitialize() 
    try:
        outlook = win32.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)
        mail.To = destinatario
        mail.Subject = 'Seu Token de Acesso - Célula Python Monitoração'
        mail.HTMLBody = f'''
            <div style="font-family: Arial, sans-serif; color: #333; max-width: 600px; margin: 0 auto; border: 1px solid #e0e0e0; border-radius: 8px; padding: 20px;">
                <h2 style="color: #242424; border-bottom: 2px solid #d3ad65; padding-bottom: 10px;">Portal de Mensageria - C6 Bank</h2>
                <p>Olá,</p>
                <p>Seu código de autenticação seguro (Token) para upload de arquivos é:</p>
                <div style="text-align: center; margin: 30px 0;">
                    <h1 style="color: #242424; background-color: #f3f4f6; padding: 15px 30px; display: inline-block; border-radius: 8px; letter-spacing: 8px; margin: 0; font-size: 32px; border: 1px solid #ccc;">{token}</h1>
                </div>
                <p style="color: #d32f2f; font-size: 13px;"><b>Atenção:</b> Este código expira em 2 minutos. Não o compartilhe com ninguém.</p>
                <br>
                <p>Atenciosamente,<br><b>Célula Python Monitoração</b></p>
            </div>
        '''
        mail.Send()
        logger.info(f"Token de acesso enviado com sucesso para {destinatario}")
        return True
    except Exception as e:
        logger.error(f"Erro ao enviar email: {e}")
        return False
    finally:
        pythoncom.CoUninitialize()

def get_todas_pastas_raiz():
    if not os.path.exists(BASE_PATH):
        return []
    return [d for d in os.listdir(BASE_PATH) if os.path.isdir(os.path.join(BASE_PATH, d)) and not d.startswith('_')]

def ler_pastas_permitidas(username):
    """
    Lê a planilha UPLOADERS.xlsx e retorna a lista de pastas que o usuário tem permissão.
    Suporta múltiplos usuários por linha e múltiplas pastas por linha (separados por , ou ;).
    """
    pastas_permitidas = []
    username_limpo = username.strip().lower()

    if username_limpo in [u.lower() for u in ADMIN_USERS]:
        return ['ALL']

    if not os.path.exists(EXCEL_FILE):
        logger.warning(f"Planilha de permissões não encontrada: {EXCEL_FILE}")
        return pastas_permitidas
        
    try:
        import re
        df = pd.read_excel(EXCEL_FILE)
        # Garantir que as colunas são tratadas como string e remover NaNs
        df['PASTA'] = df['PASTA'].fillna('').astype(str)
        df['USERS'] = df['USERS'].fillna('').astype(str)
        
        for index, row in df.iterrows():
            # Suporta separadores , ou ;
            usuarios_na_linha = [u.strip().lower() for u in re.split(',|;', row['USERS']) if u.strip()]
            
            if username_limpo in usuarios_na_linha:
                pastas_na_linha = [p.strip() for p in re.split(',|;', row['PASTA']) if p.strip()]
                pastas_permitidas.extend(pastas_na_linha)
                
        # Se encontrou 'ALL' em qualquer lugar, libera tudo
        if any(p.upper() == 'ALL' for p in pastas_permitidas):
            return ['ALL']
            
        # Remover duplicatas mantendo a ordem (opcional, mas set resolve)
        pastas_permitidas = list(dict.fromkeys(pastas_permitidas))
        
    except Exception as e:
        logger.error(f"Erro crítico ao ler Excel {EXCEL_FILE}: {e}")
    
    return pastas_permitidas

def mapear_diretorios_arquivos_input(pastas_permitidas):
    """
    Com base nas pastas permitidas, varre o disco procurando as subpastas 'arquivos_input'.
    Retorna o dicionário agrupado para o frontend e o set de caminhos válidos para trava de segurança.
    """
    grouped_diretorios = {}
    valid_paths = set()
    
    # Se for ALL, pega todas as pastas reais no diretório automacoes
    if any(p.upper() == 'ALL' for p in pastas_permitidas):
        pastas_permitidas = get_todas_pastas_raiz()

    for pasta_raiz in pastas_permitidas:
        caminho_base_pasta = os.path.join(BASE_PATH, pasta_raiz)
        caminho_input = os.path.join(caminho_base_pasta, "arquivos_input")
        
        if os.path.exists(caminho_input):
            if pasta_raiz not in grouped_diretorios:
                grouped_diretorios[pasta_raiz] = {}
                
            grouped_diretorios[pasta_raiz][caminho_input] = "arquivos_input (Raiz)"
            valid_paths.add(os.path.abspath(caminho_input))
            
            # Varredura de subpastas dentro de arquivos_input
            for root, dirs, files in os.walk(caminho_input):
                for d in dirs:
                    caminho_completo = os.path.join(root, d)
                    caminho_relativo = os.path.relpath(caminho_completo, start=caminho_input)
                    nome_exibicao = f"arquivos_input / {caminho_relativo.replace(os.sep, ' / ')}"
                    grouped_diretorios[pasta_raiz][caminho_completo] = nome_exibicao
                    valid_paths.add(os.path.abspath(caminho_completo))
        else:
            # LOG DE APOIO AO ADMINISTRADOR
            if not os.path.exists(caminho_base_pasta):
                logger.warning(f"Permissão negada ou Pasta não encontrada no disco: '{pasta_raiz}' (Caminho esperado: {caminho_base_pasta})")
            else:
                logger.info(f"Pasta '{pasta_raiz}' existe, mas não possui subpasta 'arquivos_input'. Ignorada no portal.")
                    
    return grouped_diretorios, valid_paths

def localizar_script_automacao(target_path):
    """Analisa o caminho do upload e localiza o script Python correspondente na pasta metodos."""
    try:
        caminho_norm = os.path.normpath(target_path)
        partes = caminho_norm.split(os.sep)
        
        # Encontra onde está "arquivos_input" no caminho
        idx_input = partes.index("arquivos_input")
        
        # A pasta raiz da célula é a pasta imediatamente anterior a "arquivos_input"
        pasta_raiz = partes[idx_input - 1] 
        nome_da_pasta_alvo = partes[-1]
        
        # Se o usuário subir direto na raiz do arquivos_input (sem subpasta)
        if nome_da_pasta_alvo == "arquivos_input":
            return None, False, nome_da_pasta_alvo
            
        # Constrói o caminho para a pasta "metodos"
        caminho_ate_raiz = os.sep.join(partes[:idx_input])
        caminho_script = os.path.join(caminho_ate_raiz, "metodos", f"{nome_da_pasta_alvo}.py")
        
        return caminho_script, os.path.exists(caminho_script), nome_da_pasta_alvo
    except Exception as e:
        logger.error(f"Erro ao tentar localizar o script correspondente: {e}")
        return None, False, "Desconhecido"

@app.route('/', methods=['GET'])
def index():
    if 'username' not in session:
        return render_template('uploaders.html', state='login')
    
    username = session['username']
    pastas_permitidas = ler_pastas_permitidas(username)
    grouped_diretorios, valid_paths = mapear_diretorios_arquivos_input(pastas_permitidas)
    is_admin = username in ADMIN_USERS

    return render_template('uploaders.html', state='upload', diretorios=grouped_diretorios, username=username, is_admin=is_admin)

@app.route('/upload_ajax', methods=['POST'])
def upload_ajax():
    """Endpoint chamado via JavaScript (AJAX) para subir o arquivo e descobrir o script."""
    if 'username' not in session:
        return jsonify({"status": "error", "message": "Sessão expirada."}), 401
        
    username = session['username']
    pastas_permitidas = ler_pastas_permitidas(username)
    _, valid_paths = mapear_diretorios_arquivos_input(pastas_permitidas)
    
    if 'file' not in request.files:
        return jsonify({"status": "error", "message": "Nenhum arquivo enviado."})
        
    file = request.files['file']
    target_path = request.form.get('target_folder')
    execution_mode = request.form.get('execution_mode', 'upload_run')
    
    if file.filename == '':
        return jsonify({"status": "error", "message": "Nome do arquivo vazio."})
        
    if target_path not in valid_paths:
        return jsonify({"status": "error", "message": "Acesso negado a este diretório."})
        
    filename = secure_filename(file.filename)
    caminho_final = os.path.join(target_path, filename)
    
    try:
        # 1. Salva o arquivo
        file.save(caminho_final)
        logger.info(f"Arquivo '{filename}' salvo por '{username}' com sucesso em: {caminho_final}")
        
        if execution_mode == 'upload_only':
            return jsonify({
                "status": "success",
                "execution_mode": "upload_only",
                "file_saved": filename,
                "directory": target_path
            })

        # 2. Procura o script correspondente
        caminho_script, existe, nome_alvo = localizar_script_automacao(target_path)
        
        if existe:
            # Salva na sessão qual script ele tem permissão para rodar agora
            session['pending_script'] = caminho_script
            return jsonify({
                "status": "success", 
                "script_exists": True, 
                "script_name": f"{nome_alvo}.py"
            })
        else:
            return jsonify({
                "status": "success", 
                "script_exists": False, 
                "folder_name": nome_alvo
            })
            
    except Exception as e:
        logger.error(f"Erro ao salvar arquivo: {e}")
        return jsonify({"status": "error", "message": f"Erro interno: {e}"})

@app.route('/stream_logs')
def stream_logs():
    """Endpoint SSE que executa o script e envia os logs em tempo real para a tela."""
    script_path = session.get('pending_script')
    username_logado = session.get('username', 'desconhecido')
    
    # Limpa o script da sessão imediatamente para evitar re-execução em refresh
    # e para evitar erros de contexto dentro do gerador
    session.pop('pending_script', None)
    
    if not script_path or not os.path.exists(script_path):
        def error_gen():
            yield "data: [ERRO_INTERNO] Script não autorizado ou não encontrado na sessão.\n\n"
        return Response(error_gen(), mimetype='text/event-stream')

    def generate():
        try:
            yield f"data: [*] Iniciando automação: {os.path.basename(script_path)}...\n\n"
            
            # Se for executável, sys.executable apontará pro servidor.exe, e não pro python.
            # Portanto, precisamos usar 'python' pra chamar o script python real da automação.
            python_cmd = sys.executable if not getattr(sys, 'frozen', False) else 'python'
            
            # --- INJEÇÃO DE CONTEXTO (AUDITORIA C6 BANK) ---
            # Clona o ambiente do sistema atual
            current_env = os.environ.copy()
            # Força o Python a não "segurar" as linhas de print no buffer
            current_env['PYTHONUNBUFFERED'] = '1'
            # Identifica que a execução foi iniciada via botão pelo usuário na tela
            current_env['ENV_EXEC_MODE'] = 'SOLICITACAO'
            # Passa o nome do usuário que apertou o botão (email do C6)
            current_env['ENV_EXEC_USER'] = f"{username_logado}@c6bank.com"
            
            # Executa passando o novo "env" clonado e injetado com as métricas
            process = subprocess.Popen(
                [python_cmd, '-u', script_path],
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT, # Junta erro e saida normal no mesmo lugar
                text=True,
                env=current_env,
                bufsize=1
            )
            
            # Lê cada linha do console do script gerado e envia pro frontend
            for line in iter(process.stdout.readline, ''):
                yield f"data: {line}\n\n"
                
            process.stdout.close()
            return_code = process.wait()
            
            if return_code == 0:
                logger.info(f"Automação '{os.path.basename(script_path)}' finalizada com SUCESSO.")
                yield f"data: [CONCLUIDO_SUCESSO]\n\n"
            else:
                logger.warning(f"Automação '{os.path.basename(script_path)}' finalizou com ERRO (Código {return_code}).")
                yield f"data: [CONCLUIDO_ERRO] Automação finalizou com código {return_code}.\n\n"
                
        except Exception as e:
            logger.error(f"Falha crítica ao iniciar processo de automação: {str(e)}")
            yield f"data: [ERRO_INTERNO] Falha ao iniciar processo: {str(e)}\n\n"

    return Response(generate(), mimetype='text/event-stream')

@app.route('/request_token', methods=['POST'])
def request_token():
    username = request.form.get('username', '').strip().lower()
    
    if not username:
        flash('Digite um usuário válido.', 'error')
        return redirect(url_for('index'))

    if username in ADMIN_USERS:
        session['username'] = username
        flash(f'Bem-vindo(a) Administrador(a) {username}!', 'success')
        return redirect(url_for('index'))
        
    pastas_permitidas = ler_pastas_permitidas(username)
    
    if not pastas_permitidas:
        flash('Usuário não encontrado ou sem acesso a pastas.', 'error')
        return redirect(url_for('index'))
        
    token = ''.join(random.choices(string.digits, k=6))
    expire_time = datetime.now() + timedelta(minutes=2)
    auth_tokens[username] = {'token': token, 'expires': expire_time}
    
    email = f"{username}{DOMAIN}"
    if enviar_email_outlook(email, token):
        flash(f'Token enviado para {email}.', 'info')
        return render_template('uploaders.html', state='verify', username=username)
    else:
        flash('Erro ao enviar e-mail pelo Outlook.', 'error')
        return redirect(url_for('index'))

@app.route('/verify_token', methods=['POST'])
def verify_token():
    username = request.form.get('username')
    token_inserido = request.form.get('token')
    
    if username in auth_tokens:
        dados_token = auth_tokens[username]
        if datetime.now() > dados_token['expires']:
            flash('Token expirado.', 'error')
            del auth_tokens[username]
            return redirect(url_for('index'))
            
        if token_inserido == dados_token['token']:
            session['username'] = username
            session.permanent = True  # Ativa o limite de 24 horas definido em PERMANENT_SESSION_LIFETIME
            del auth_tokens[username]
            return redirect(url_for('index'))
        else:
            flash('Token inválido.', 'error')
            return render_template('uploaders.html', state='verify', username=username)
            
    flash('Sessão expirada.', 'error')
    return redirect(url_for('index'))

@app.route('/logout')
def logout():
    session.pop('username', None)
    return redirect(url_for('index'))

if __name__ == '__main__':
    my_ip = get_ip()
    logger.info("="*70)
    logger.info(f"SERVIDORES INICIADOS (IP: {my_ip}) - C6 BANK")
    logger.info(f"Log do Servidor configurado em: {log_dir}")
    logger.info(f"Acesse na rede: http://{my_ip}:5000")
    logger.info("="*70)
    app.run(host='0.0.0.0', port=5000, debug=True)
