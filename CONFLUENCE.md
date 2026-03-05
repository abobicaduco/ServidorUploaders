# Documentação Interna: Mensageria e Cargas Operacionais (Python)

A automação e a execução de scripts em Python da área de Mensageria e Cargas Operacionais agora podem ser acessados via Intranet/Navegador utilizando o **Servidor Uploaders**.
A plataforma permite às áreas demandarem orquestração de robôs, realizar upload de arquivos de carga, e ter retornos transparentes dos servidores automatizados.

## 🔐 1. Acesso ao Sistema e Segurança (Token)
Ao acessar a página do Servidor Uploaders:
1. Insira seu Login da rede (Exemplo: `carlos.lsilva`). Não é necessário adicionar `@c6bank.com`.
2. O sistema verificará se seu usuário tem acesso mapeado. Caso positivo, enviará um Token numérico de 6 dígitos via Microsoft Outlook para seu email (O seu PC no qual roda o servidor deve ter sessão do Outlook iniciada).
3. Insira o Token no portal.
4. **Validade e Duração:**
   * O Token enviado no email expira em **2 minutos** por questões de segurança (LGPD). 
   * A sessão do portal dura **24 horas**. Isso significa que múltiplos robôs logados com o mesmo usuário não irão se desconectar e você pode trabalhar normalmente. Caso esteja logado em sua máquina e precise logar em outra simultaneamente, um novo token será disparado e permitirá duplo acesso sem interromper a sessão anterior (Multithreading).

## 📂 2. Layout, Pastas e Envios
Ao se conectar, você verá uma lista de **Diretórios Permitidos**, conforme a hierarquia da sua equipe preestabelecida na planilha matriz de controle (`UPLOADERS.xlsx`).

### Passos de Execução
1. Faça o "Upload" na respectiva pasta (ou apenas arraste os arquivos Excel/CSV pra janela). O sistema alocará o arquivo nas pastas compartilhadas da Rede (`C6 CTVM LTDA/Mensageria [...]`).
2. Assim que o Uploader receber o envio, ele exibirá o nome do script (se encontrar o modelo referente).
3. Aperte `Executar Automação`.
4. Uma janela pop-up se abrirá exibindo o log e console em **Tempo Real**, idêntico à execução do Visual Studio Code. Todo log que rola neste momento trata-se de um processamento físico do servidor sendo roteado em formato SSE para seu navegador.
5. Aguarde o final da execução, que emitirá em verde "Concluído com Sucesso" (ou o aviso de erro, com relatório). Todos os logs gerados são guardados fisicamente na respectiva pasta gerencial do Servidor.

## 🗃️ 3. Informações Técnicas a Processos
* Todo processo roda de maneira **Isolada (Novo PID)** gerado pelo Windows Server subjacente. Arquivos sendo lidos por robô X em sua tela **jamais se misturam** e não causam corrupção de variáveis de processos alheios ou de outro usuário operando simultaneamente o mesmo servidor.
* Contudo, duas execuções *da mesma rotina* num mesmo *arquivo físico editável (Excel)* resultará no Lock do Windows por conta de restrição de File System local.

*Contato Suporte / Manutenção: Célula Python*
