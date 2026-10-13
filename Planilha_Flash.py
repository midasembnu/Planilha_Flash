import os
import re
import json
import shutil
import queue
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog

import pdfplumber
import gspread
from google.oauth2.service_account import Credentials


# ============================================================
# CONFIGURAÇÃO PERSISTENTE
# ============================================================

APP_NAME = "LancadorNF"
CONFIG_DIR = os.path.join(os.getenv("APPDATA", os.path.expanduser("~")), APP_NAME)
CONFIG_PATH = os.path.join(CONFIG_DIR, "config.json")

CONFIG_PADRAO = {
    "credentials_path": "",
    "nome_aba": "",
    "pastas_origem": [],
    "pasta_backup": "",
    "planilhas": []
}


def carregar_config():
    if not os.path.exists(CONFIG_PATH):
        return dict(CONFIG_PADRAO)
    try:
        with open(CONFIG_PATH, "r", encoding="utf-8") as f:
            dados = json.load(f)
        config = dict(CONFIG_PADRAO)
        config.update(dados)
        return config
    except Exception:
        return dict(CONFIG_PADRAO)


def salvar_config(config):
    os.makedirs(CONFIG_DIR, exist_ok=True)
    with open(CONFIG_PATH, "w", encoding="utf-8") as f:
        json.dump(config, f, ensure_ascii=False, indent=2)


# ============================================================
# LÓGICA DE EXTRAÇÃO DE PDF (adaptado de aut.py)
# ============================================================

def extrair_dados_pdf(caminho_pdf):
    with pdfplumber.open(caminho_pdf) as pdf:
        texto_completo = ""
        todas_as_linhas = []
        for pagina in pdf.pages:
            texto_bruto = pagina.extract_text()
            if texto_bruto:
                texto_completo += texto_bruto + "\n"
                todas_as_linhas.extend(texto_bruto.split("\n"))

        # Número da NF
        nf_match = re.search(r"(?:N[°º\s]*|NÚMERO\s*)(\d+)", texto_completo, re.IGNORECASE)
        numero_nf = str(int(nf_match.group(1))) if nf_match else "N/A"

        # Nome do cliente
        cliente = "N/A"
        for i, linha in enumerate(todas_as_linhas):
            if "NOME" in linha.upper() and "SOCIAL" in linha.upper():
                if i + 1 < len(todas_as_linhas):
                    raw_cliente = todas_as_linhas[i + 1].strip()
                    nome_limpo = re.split(r"\d{11,}", raw_cliente)[0].strip()
                    cliente = re.split(r"\(COD:", nome_limpo)[0].strip()
                    cliente = re.sub(r"[-.]$", "", cliente).strip()
                break

        # Representante
        repres_match = re.search(r"REPRES\.:\s*([^\n]*)", texto_completo)
        representante = repres_match.group(1).strip() if repres_match else "N/A"

        # Itens (código e quantidade)
        itens_extraidos = []
        for linha in todas_as_linhas:
            partes = linha.split()
            if partes and partes[0].isdigit() and len(partes) >= 6:
                codigo = partes[0]
                for idx, p in enumerate(partes):
                    if p.upper() in ["UN", "UNI", "UNID", "CX", "CAIXA", "PC", "PT", "KG", "DP"] and (idx + 1) < len(partes):
                        qtd_str = partes[idx + 1].replace(".", "").replace(",", ".")
                        try:
                            qtd = str(int(float(qtd_str)))
                            itens_extraidos.append({"codigo": codigo, "qtd": qtd})
                            break
                        except ValueError:
                            continue

        return numero_nf, cliente, representante, itens_extraidos


# ============================================================
# LÓGICA DE BUSCA DO PDF (adaptado de s.py)
# ============================================================

def buscar_pdf(numero, serie, pastas_origem, log_callback):
    padrao_procurado = f"danfe_{serie}_{numero}.pdf".lower()

    for pasta in pastas_origem:
        if not os.path.exists(pasta):
            log_callback(f"Aviso: pasta de origem inacessível: {pasta}")
            continue
        try:
            arquivos_na_pasta = os.listdir(pasta)
        except Exception:
            log_callback(f"Aviso: erro ao ler a pasta: {pasta}")
            continue

        for arquivo in arquivos_na_pasta:
            if arquivo.lower() == padrao_procurado:
                return os.path.join(pasta, arquivo)

    return None


# ============================================================
# LÓGICA DE CONEXÃO COM GOOGLE SHEETS
# ============================================================

def conectar_planilha(credentials_path, nome_planilha, nome_aba):
    scope = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
    creds = Credentials.from_service_account_file(credentials_path, scopes=scope)
    client = gspread.authorize(creds)
    return client.open(nome_planilha).worksheet(nome_aba)


# ============================================================
# PARSE DA ENTRADA DE NFs
# ============================================================

def parse_entrada(texto):
    itens_brutos = texto.replace(",", " ").split()
    validos = []
    invalidos = []
    for item in itens_brutos:
        if "-" not in item:
            invalidos.append(item)
            continue
        numero, serie = item.split("-", 1)
        numero = numero.strip()
        serie = serie.strip()
        if not numero or not serie:
            invalidos.append(item)
            continue
        validos.append((numero, serie, item))
    return validos, invalidos


# ============================================================
# PROCESSAMENTO PRINCIPAL (roda em thread separada)
# ============================================================

def processar_lote(itens_texto, planilha_nome, config, log_callback, done_callback):
    try:
        validos, invalidos = parse_entrada(itens_texto)

        relatorio = {
            "lancadas": [],
            "nao_encontradas": [],
            "duplicadas": [],
            "erros": [],
            "formato_invalido": invalidos,
        }

        if not validos and not invalidos:
            log_callback("Nenhuma NF foi digitada.")
            done_callback()
            return

        for item in invalidos:
            log_callback(f"Formato inválido: '{item}' (use numero-serie, ex: 21660-1)")

        if not validos:
            done_callback()
            return

        log_callback(f"Conectando na planilha '{planilha_nome}'...")
        try:
            aba = conectar_planilha(config["credentials_path"], planilha_nome, config["nome_aba"])
            nfs_na_planilha = aba.col_values(1)
        except Exception as e:
            log_callback(f"Erro ao conectar na planilha: {e}")
            done_callback()
            return

        bloco_total = []
        copias_pendentes = []  # (caminho_pdf, nome_arquivo) para copiar após lançar com sucesso

        for numero, serie, item_original in validos:
            log_callback(f"Buscando NF {item_original}...")
            caminho_pdf = buscar_pdf(numero, serie, config["pastas_origem"], log_callback)

            if not caminho_pdf:
                relatorio["nao_encontradas"].append(item_original)
                log_callback(f"  -> Não encontrada: {item_original}")
                continue

            try:
                nf, cliente, repres, itens = extrair_dados_pdf(caminho_pdf)
            except Exception as e:
                relatorio["erros"].append(f"{item_original} (falha ao ler PDF: {e})")
                log_callback(f"  -> Erro ao extrair dados de {item_original}: {e}")
                continue

            if nf == "N/A":
                relatorio["erros"].append(f"{item_original} (não foi possível extrair o número da NF)")
                log_callback(f"  -> Erro: não foi possível extrair o número da NF de {item_original}")
                continue

            if nf in nfs_na_planilha:
                relatorio["duplicadas"].append(item_original)
                log_callback(f"  -> Já existente na planilha: NF {nf}")
                continue

            # Monta o bloco de linhas dessa NF
            if bloco_total or nfs_na_planilha:
                bloco_total.append(["", "", ""])
            bloco_total.append([nf, cliente, repres])
            for it in itens:
                bloco_total.append([it["codigo"], None, it["qtd"]])

            nfs_na_planilha.append(nf)
            copias_pendentes.append((caminho_pdf, os.path.basename(caminho_pdf)))
            relatorio["lancadas"].append(item_original)
            log_callback(f"  -> OK: NF {nf} | {cliente} | {len(itens)} itens")

        if bloco_total:
            log_callback("Enviando dados para a planilha...")
            try:
                aba.append_rows(bloco_total, value_input_option="USER_ENTERED")
            except Exception as e:
                log_callback(f"Erro ao enviar dados para a planilha: {e}")
                done_callback()
                return

            pasta_backup = config.get("pasta_backup", "")
            if pasta_backup:
                os.makedirs(pasta_backup, exist_ok=True)
                for caminho_pdf, nome_arquivo in copias_pendentes:
                    try:
                        shutil.copy2(caminho_pdf, os.path.join(pasta_backup, nome_arquivo))
                    except Exception as e:
                        log_callback(f"Aviso: falha ao copiar backup de {nome_arquivo}: {e}")

        # Relatório final
        log_callback("-" * 60)
        log_callback("Processo finalizado!")
        log_callback(f"NFs lançadas: {len(relatorio['lancadas'])}")
        if relatorio["duplicadas"]:
            log_callback(f"Já existentes: {', '.join(relatorio['duplicadas'])}")
        if relatorio["nao_encontradas"]:
            log_callback(f"Não encontradas: {', '.join(relatorio['nao_encontradas'])}")
        if relatorio["erros"]:
            log_callback(f"Erros: {', '.join(relatorio['erros'])}")
        log_callback("-" * 60)

    except Exception as e:
        log_callback(f"Erro inesperado: {e}")
    finally:
        done_callback()


# ============================================================
# INTERFACE GRÁFICA
# ============================================================

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Lançador de NFs")
        self.geometry("720x560")
        self.minsize(640, 480)

        self.config_dados = carregar_config()
        self.log_queue = queue.Queue()
        self.processando = False

        notebook = ttk.Notebook(self)
        notebook.pack(fill="both", expand=True, padx=10, pady=10)

        self.aba_lancamento = ttk.Frame(notebook)
        self.aba_config = ttk.Frame(notebook)
        notebook.add(self.aba_lancamento, text="Lançamento")
        notebook.add(self.aba_config, text="Configurações")
        notebook.bind("<<NotebookTabChanged>>", self._on_tab_changed)

        self._montar_aba_lancamento()
        self._montar_aba_config()

        self.after(100, self._poll_log_queue)

    # --------------------------------------------------------
    # ABA 1 - LANÇAMENTO
    # --------------------------------------------------------
    def _montar_aba_lancamento(self):
        frame = self.aba_lancamento

        ttk.Label(frame, text="NFs (número-série), separadas por espaço ou vírgula:").pack(
            anchor="w", padx=10, pady=(10, 0)
        )
        self.txt_nfs = tk.Text(frame, height=4)
        self.txt_nfs.pack(fill="x", padx=10, pady=5)

        linha_planilha = ttk.Frame(frame)
        linha_planilha.pack(fill="x", padx=10, pady=5)
        ttk.Label(linha_planilha, text="Planilha:").pack(side="left")
        self.combo_planilha = ttk.Combobox(linha_planilha, state="readonly")
        self.combo_planilha.pack(side="left", padx=10, fill="x", expand=True)
        self._atualizar_combo_planilhas()

        self.btn_ok = ttk.Button(frame, text="OK", command=self._on_clicar_ok)
        self.btn_ok.pack(pady=10)

        ttk.Label(frame, text="Relatório:").pack(anchor="w", padx=10)
        log_frame = ttk.Frame(frame)
        log_frame.pack(fill="both", expand=True, padx=10, pady=(0, 10))

        scrollbar = ttk.Scrollbar(log_frame)
        scrollbar.pack(side="right", fill="y")
        self.txt_log = tk.Text(log_frame, state="disabled", yscrollcommand=scrollbar.set)
        self.txt_log.pack(fill="both", expand=True)
        scrollbar.config(command=self.txt_log.yview)

    def _atualizar_combo_planilhas(self):
        planilhas = self.config_dados.get("planilhas", [])
        self.combo_planilha["values"] = planilhas
        if planilhas and not self.combo_planilha.get():
            self.combo_planilha.current(0)

    def _on_tab_changed(self, event):
        self._atualizar_combo_planilhas()

    def _log(self, mensagem):
        self.log_queue.put(mensagem)

    def _poll_log_queue(self):
        try:
            while True:
                mensagem = self.log_queue.get_nowait()
                self.txt_log.config(state="normal")
                self.txt_log.insert("end", mensagem + "\n")
                self.txt_log.see("end")
                self.txt_log.config(state="disabled")
        except queue.Empty:
            pass
        self.after(100, self._poll_log_queue)

    def _on_clicar_ok(self):
        if self.processando:
            return

        texto = self.txt_nfs.get("1.0", "end").strip()
        planilha_nome = self.combo_planilha.get()

        if not texto:
            messagebox.showwarning("Atenção", "Digite ao menos uma NF.")
            return
        if not planilha_nome:
            messagebox.showwarning("Atenção", "Selecione uma planilha.")
            return
        if not self.config_dados.get("credentials_path") or not self.config_dados.get("nome_aba"):
            messagebox.showwarning("Atenção", "Configure o credentials.json e o nome da aba na aba Configurações.")
            return
        if not self.config_dados.get("pastas_origem"):
            messagebox.showwarning("Atenção", "Configure ao menos uma pasta de origem na aba Configurações.")
            return

        self.processando = True
        self.btn_ok.config(state="disabled", text="Processando...")

        def worker():
            processar_lote(
                texto,
                planilha_nome,
                self.config_dados,
                self._log,
                self._on_processamento_concluido,
            )

        threading.Thread(target=worker, daemon=True).start()

    def _on_processamento_concluido(self):
        def atualizar_ui():
            self.processando = False
            self.btn_ok.config(state="normal", text="OK")
        self.after(0, atualizar_ui)

    # --------------------------------------------------------
    # ABA 2 - CONFIGURAÇÕES
    # --------------------------------------------------------
    def _montar_aba_config(self):
        frame = self.aba_config

        # Pastas de origem
        ttk.Label(frame, text="Pastas de origem (busca dos PDFs):").pack(anchor="w", padx=10, pady=(10, 0))
        frame_pastas = ttk.Frame(frame)
        frame_pastas.pack(fill="x", padx=10, pady=5)
        self.lista_pastas = tk.Listbox(frame_pastas, height=4)
        self.lista_pastas.pack(side="left", fill="x", expand=True)
        for pasta in self.config_dados.get("pastas_origem", []):
            self.lista_pastas.insert("end", pasta)

        frame_botoes_pastas = ttk.Frame(frame_pastas)
        frame_botoes_pastas.pack(side="left", padx=5)
        ttk.Button(frame_botoes_pastas, text="Adicionar", command=self._adicionar_pasta).pack(fill="x")
        ttk.Button(frame_botoes_pastas, text="Remover", command=self._remover_pasta).pack(fill="x", pady=(5, 0))

        # Pasta de backup
        ttk.Label(frame, text="Pasta de backup local (cópia dos PDFs lançados):").pack(
            anchor="w", padx=10, pady=(10, 0)
        )
        frame_backup = ttk.Frame(frame)
        frame_backup.pack(fill="x", padx=10, pady=5)
        self.entry_backup = ttk.Entry(frame_backup)
        self.entry_backup.insert(0, self.config_dados.get("pasta_backup", ""))
        self.entry_backup.pack(side="left", fill="x", expand=True)
        ttk.Button(frame_backup, text="Procurar...", command=self._escolher_pasta_backup).pack(side="left", padx=5)

        # Planilhas
        ttk.Label(frame, text="Planilhas cadastradas:").pack(anchor="w", padx=10, pady=(10, 0))
        frame_planilhas = ttk.Frame(frame)
        frame_planilhas.pack(fill="x", padx=10, pady=5)
        self.lista_planilhas = tk.Listbox(frame_planilhas, height=4)
        self.lista_planilhas.pack(side="left", fill="x", expand=True)
        for planilha in self.config_dados.get("planilhas", []):
            self.lista_planilhas.insert("end", planilha)

        frame_botoes_planilhas = ttk.Frame(frame_planilhas)
        frame_botoes_planilhas.pack(side="left", padx=5)
        ttk.Button(frame_botoes_planilhas, text="Adicionar", command=self._adicionar_planilha).pack(fill="x")
        ttk.Button(frame_botoes_planilhas, text="Remover", command=self._remover_planilha).pack(fill="x", pady=(5, 0))

        # Nome da aba
        ttk.Label(frame, text="Nome da aba (igual em todas as planilhas):").pack(anchor="w", padx=10, pady=(10, 0))
        self.entry_aba = ttk.Entry(frame)
        self.entry_aba.insert(0, self.config_dados.get("nome_aba", ""))
        self.entry_aba.pack(fill="x", padx=10, pady=5)

        # Credentials.json
        ttk.Label(frame, text="Arquivo credentials.json:").pack(anchor="w", padx=10, pady=(10, 0))
        frame_cred = ttk.Frame(frame)
        frame_cred.pack(fill="x", padx=10, pady=5)
        self.entry_credentials = ttk.Entry(frame_cred)
        self.entry_credentials.insert(0, self.config_dados.get("credentials_path", ""))
        self.entry_credentials.pack(side="left", fill="x", expand=True)
        ttk.Button(frame_cred, text="Procurar...", command=self._escolher_credentials).pack(side="left", padx=5)

        # Salvar
        ttk.Button(frame, text="Salvar configurações", command=self._salvar_configuracoes).pack(pady=15)

    def _adicionar_pasta(self):
        pasta = filedialog.askdirectory(title="Selecione a pasta de origem")
        if pasta:
            self.lista_pastas.insert("end", pasta)

    def _remover_pasta(self):
        selecao = self.lista_pastas.curselection()
        if selecao:
            self.lista_pastas.delete(selecao[0])

    def _escolher_pasta_backup(self):
        pasta = filedialog.askdirectory(title="Selecione a pasta de backup")
        if pasta:
            self.entry_backup.delete(0, "end")
            self.entry_backup.insert(0, pasta)

    def _adicionar_planilha(self):
        nome = simpledialog.askstring("Nova planilha", "Nome exato da planilha no Google Sheets:")
        if nome:
            nome = nome.strip()
            if nome:
                self.lista_planilhas.insert("end", nome)

    def _remover_planilha(self):
        selecao = self.lista_planilhas.curselection()
        if selecao:
            self.lista_planilhas.delete(selecao[0])

    def _escolher_credentials(self):
        caminho = filedialog.askopenfilename(
            title="Selecione o credentials.json",
            filetypes=[("Arquivo JSON", "*.json")],
        )
        if caminho:
            self.entry_credentials.delete(0, "end")
            self.entry_credentials.insert(0, caminho)

    def _salvar_configuracoes(self):
        self.config_dados["pastas_origem"] = list(self.lista_pastas.get(0, "end"))
        self.config_dados["pasta_backup"] = self.entry_backup.get().strip()
        self.config_dados["planilhas"] = list(self.lista_planilhas.get(0, "end"))
        self.config_dados["nome_aba"] = self.entry_aba.get().strip()
        self.config_dados["credentials_path"] = self.entry_credentials.get().strip()

        salvar_config(self.config_dados)
        self._atualizar_combo_planilhas()
        messagebox.showinfo("Configurações", "Configurações salvas com sucesso.")


if __name__ == "__main__":
    app = App()
    app.mainloop()
