# Servidor Uploaders (Portal de Mensageria)

![Python](https://img.shields.io/badge/python-3.x-blue.svg)
![Flask](https://img.shields.io/badge/flask-latest-green.svg)

Um servidor web seguro, desenvolvido em Python usando Flask, para gerenciamento de uploads de arquivos e execução de scripts de automação (RPA) de forma controlada a partir de uma interface amigável.

## 🚀 Como Funciona

O Servidor Uploaders atua como um portal self-service onde os usuários podem:
1. **Fazer login de forma segura** através de um Token de Acesso temporário encaminhado para o e-mail corporativo.
2. **Fazer upload de arquivos** para diretórios específicos (controlados por permissões).
3. **Acionar automações Python** que processam os arquivos enviados, visualizando em tempo real o log de execução no próprio navegador.
4. **Gerenciamento de permissões** via planilha Excel, distribuindo acessos granulares por pastas e rotinas a diferentes colaboradores.

### Arquitetura de Execução, Sessões e Threads
- **Sessões e Tokens**: Ao solicitar o acesso, um código é gerado e expira em **2 minutos** caso não seja utilizado. Após logado, a sessão permanece ativa por **24 horas**, suportando uso contínuo (mesmo usuário em máquinas diferentes).
- **Multithreading**: A navegação web é gerenciada em threads do Flask, suportando alta concorrência.
- **Isolamento de Processos (PIDs)**: O acionamento de um script de automação gera um novo Processo em nível de SO (subprocesso/PID isolado). Dessa forma, a memória, logs e fluxos de dados de diferentes automações rodando ao mesmo tempo não se misturam nem corrompem os dados.

## ⚙️ Instalação e Configuração

O projeto conta com um sistema de auto-restauração/instalação de dependências automáticas no momento da execução inicial.

### 1. Clonando e Instalando Dependências

```bash
git clone https://github.com/abobicaduco/ServidorUploaders.git
cd ServidorUploaders
```
As bibliotecas serão instaladas automaticamente no primeiro acesso (`Flask`, `pandas`, `pywin32`, `python-dotenv`, etc.), desde que o `pip` esteja configurado no sistema.

### 2. Configurando o Ambiente (`.env`)

Para não codificar dados sensíveis (Hardcode) no código e para permitir o desenvolvimento contínuo (em seu computador pessoal ou fora do ambiente restrito da empresa), usamos variáveis de ambiente (`.env`).

Crie um arquivo `.env` na raiz do projeto baseado no `.env.example`:

```env
# MOCK_EMAIL=True imprime o token no terminal ao invés de tentar enviar via Microsoft Outlook. Ótimo para testar localmente.
MOCK_EMAIL=True

# Lista de administradores com acesso full a todas as rotinas (separados por vírgula)
ADMIN_USERS=seu.usuario,admin

# Caminho para sobrepor o arquivo de permissões (opcional)
EXCEL_FILE=dummy_uploaders.xlsx

# Emulação de caminhos de rede (Substitua por pastas locais para poder testar)
PATH_CELULA=./sua_pasta_local
BASE_PATH=./sua_pasta_local/arquivos_input
```

### 3. Executando o Servidor

```bash
python ServidorUploaders.py
```
Acesse no seu navegador local via `http://127.0.0.1:5000` ou pelo IP da sua rede local (o IP será printado no terminal ao iniciar!).

## 🔐 Gestão de Acessos
Os acessos são mapeados em um arquivo Excel (`UPLOADERS.xlsx`) localizado na raiz. 
Ele deve possuir as colunas `PASTA` (Pastas base permitidas para os usuários) e `USERS` (Emails/Logins separados por vírgula). Para conceder acesso a todos os diretórios, utiliza-se a keyword `ALL`.

---
*Desenvolvido para orquestração de rotinas e democratização do uso de scripts internos.*
