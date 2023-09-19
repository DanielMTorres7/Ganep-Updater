# Briefing do Projeto de Automação em Python

## Objetivos do Projeto
- Automatizar o download de várias planilhas de um sistema.
- Atualizar planilhas do Google Sheets com dados baixados.
- Gerenciar arquivos e pastas no sistema operacional.
- Enviar e-mails com anexos e informações atualizadas.

## Principais Bibliotecas Utilizadas
- pyautogui: Para automação da interface do usuário.
- time: Para pausas.
- pyperclip: Para manipulação da área de transferência.
- datetime, dtTime: Para lidar com datas e horários.
- subprocess: Para executar comandos do sistema.

## Principais Funcionalidades
### Download de Planilhas
- Baixa várias planilhas de um sistema chamado "IW-Ganep".

### Atualização de Planilhas do Google
- Atualiza planilhas do Google Sheets com os dados das planilhas baixadas.

### Gestão de Arquivos
- Exclui, copia e move arquivos no sistema operacional.

### Envio de E-mails
- Envia e-mails com anexos e informações atualizadas.

## Fluxo de Trabalho
1. Faz login em um sistema para baixar várias planilhas.
2. Abre o Google Sheets e atualiza planilhas específicas com os dados baixados.
3. Baixa planilhas do Google Sheets para o sistema local.
4. Gerencia arquivos, como excluir mapas anteriores e mover novos arquivos para locais específicos.
5. Envia e-mails com as planilhas atualizadas como anexos.
