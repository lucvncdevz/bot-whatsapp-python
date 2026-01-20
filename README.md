# Bot de Envio de Mensagens no WhatsApp Web (Python)

## Visão Geral
Este projeto automatiza o envio de mensagens no **WhatsApp Web** a partir de uma **planilha Excel**, utilizando **Python**, abertura de URLs e automação de interface gráfica com `pyautogui`.

Ele lê contatos (nome e telefone) de uma planilha `.xlsx`, gera mensagens personalizadas e envia automaticamente via WhatsApp Web.

>  **Aviso**: Este projeto **não utiliza a API oficial do WhatsApp**. Ele depende de automação visual, o que o torna sensível a mudanças de layout, resolução de tela e foco da janela. 

---

## Como Funciona
1. Abre o WhatsApp Web no navegador padrão
2. Aguarda o usuário estar logado
3. Lê uma planilha Excel com nomes e telefones
4. Monta uma mensagem personalizada
5. Abre a conversa via URL (`/send?phone=...&text=...`)
6. Localiza o botão de envio na tela
7. Clica no botão e envia a mensagem
8. Fecha a aba e passa para o próximo contato

##  Tecnologias Utilizadas
- Python 3
- `openpyxl` — leitura de planilhas Excel
- `pyautogui` — automação de mouse e teclado
- `webbrowser` — abertura de links no navegador
- `urllib.parse` — codificação da mensagem na URL

## Como Executar
1. Clone o repositório  
2. Crie e ative um ambiente virtual Python  
3. Instale as dependências:
   ```bash
   pip install -r requirements.txt

## Licença

Projeto de uso educacional e experimental.

