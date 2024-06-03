Este script Python é uma solução automatizada para gerenciar e enviar mensagens via Google Messages. Ele foi projetado para operar com uma integração direta ao navegador via Selenium, gerenciando diretórios e arquivos Excel, e executando operações de envio de mensagens em massa baseadas nos números de telefone extraídos de uma planilha do Excel.

Funcionalidades
Gerenciamento de Diretórios: Cria, copia e exclui diretórios para organização de dados e backups.
Integração com Excel: Lê números de telefone de um arquivo Excel para filtrar e utilizar no envio de mensagens.
Envio Automatizado de Mensagens: Utiliza o Selenium para automatizar o envio de mensagens via Google Messages, incluindo operações como escrever, buscar contatos e enviar mensagens.
Gestão de Sessão do Navegador: Utiliza perfis do usuário para persistir e restaurar sessões do navegador, otimizando o processo de login e autenticação.
Componentes Principais
Criação e gerenciamento de diretórios: Funções para criar, copiar e excluir diretórios para gerenciamento eficiente dos dados necessários e backups.

Inicialização e manipulação do driver Selenium: Configura e inicializa uma instância do navegador com opções específicas para interagir com o Google Messages.

Operações com Excel: Lê e processa dados de números de telefone de arquivos Excel, verificando sua validade e preparando listas para o envio de mensagens.

Envio de mensagens: Automatiza o processo de envio de mensagens, incluindo a interação com a interface do Google Messages para selecionar contatos e enviar mensagens personalizadas.

Tecnologias Utilizadas
Python
Pandas para manipulação de dados em Excel
Selenium para automação web
webdriver-manager para gerenciamento do driver do navegador
Como Usar
Certifique-se de que todas as dependências estão instaladas: pandas, selenium, shutil, os, sys, time, glob, webdriver-manager.
Coloque o script no mesmo diretório que o arquivo Excel com os números de telefone.
Execute o script diretamente a partir da linha de comando ou através de um agendador de tarefas para automatização.
