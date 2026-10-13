# Planilha Flash ⚡

**PT:** Aplicativo desktop em Python (Tkinter) que busca PDFs de DANFE (nota fiscal) em pastas do computador, extrai os dados e lança em planilhas do Google Sheets.

**EN:** Python (Tkinter) desktop app that finds DANFE (invoice) PDFs in local folders, extracts the data, and adds it to Google Sheets spreadsheets.

## Por que existe / Why it exists
Criei para o meu trabalho no faturamento, para evitar digitar manualmente os dados de cada nota fiscal nas planilhas. Uso no dia a dia.

I built it for my billing job, to avoid typing each invoice's data into spreadsheets by hand. I use it daily.

## Funcionalidades / Features
- Você digita as notas no formato `número-série` (ex.: `21660-1`) / You type invoices as `number-series` (e.g. `21660-1`)
- Busca o PDF da DANFE em uma ou mais pastas / Searches for the DANFE PDF in one or more folders
- Extrai número da NF, cliente, representante e itens (código e quantidade) / Extracts invoice number, customer, sales rep, and items (code and quantity)
- Lança os dados na planilha escolhida do Google Sheets / Adds the data to the chosen Google Sheets file
- Evita lançar uma NF que já existe na planilha / Skips invoices already in the spreadsheet
- Faz cópia de backup dos PDFs lançados / Saves a backup copy of the posted PDFs
- Relatório na tela: lançadas, duplicadas, não encontradas e com erro / On-screen report: posted, duplicated, not found, and errors
- Interface que não trava durante o processamento / UI stays responsive while processing

## Tecnologias / Tech
Python, Tkinter, pdfplumber, gspread, google-auth

## Como usar / How to use
1. Instale o Python e as bibliotecas / Install Python and the libraries:
   `pip install -r requirements.txt`
2. Crie uma conta de serviço no Google Cloud e baixe o `credentials.json`. Compartilhe suas planilhas com o e-mail dessa conta. **Esse arquivo não faz parte deste repositório e nunca deve ser publicado.**
   Create a Google Cloud service account and download its `credentials.json`. Share your spreadsheets with that account's email. **This file is not part of this repository and must never be published.**
3. Execute / Run: `python Planilha_Flash.py`
4. Na aba **Configurações**, informe as pastas dos PDFs, a pasta de backup, os nomes das planilhas, o nome da aba e o `credentials.json`.
   In the **Configurações** tab, set the PDF folders, backup folder, spreadsheet names, tab name, and the `credentials.json`.

## Sobre o uso de IA / About AI use
Este projeto foi desenvolvido por mim com ajuda de inteligência artificial (Claude). Eu defini o que o programa precisa fazer, testei e uso no trabalho; a IA ajudou a escrever e melhorar o código.

I developed this project with the help of AI (Claude). I defined what the program needs to do, tested it, and use it at work; the AI helped write and improve the code.

## Autor / Author
Richardson — [LinkedIn](https://www.linkedin.com/in/richardsongp)
