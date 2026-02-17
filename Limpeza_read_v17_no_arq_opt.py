import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
import csv
import sys
import re
import threading
from datetime import datetime

# --- CONFIGURAÇÃO DE OTIMIZAÇÃO (GLOBAL) ---
pd = None

# --- CONFIGURAÇÕES VISUAIS ---
COLORS = {
    "primary": "#2C3E50", "secondary": "#ECF0F1", "card_bg": "#FFFFFF",
    "text_dark": "#2C3E50", "text_light": "#FFFFFF", "accent_blue": "#3498DB",
    "accent_orange": "#E67E22", "accent_green": "#27AE60", "accent_red": "#E74C3C",
    "accent_purple": "#8E44AD", "accent_teal": "#16A085", "accent_yellow": "#F1C40F"
}


if getattr(sys, 'frozen', False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

ICON_PATH = os.path.join(BASE_DIR, "doc.ico")
CONFIG_FILE = os.path.join(os.path.expanduser("~"), "logistica_seq_config.txt")


class LogicApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Organizador Logístico Pro v26.00")
        self.root.geometry("1100x750")
        self.root.configure(bg=COLORS["secondary"])

        # Ícone do App
        if os.path.exists(ICON_PATH):
            try:
                self.root.iconbitmap(ICON_PATH)
            except Exception as e:
                print(f"Erro ao carregar ícone: {e}")

        self.libs_carregadas = False
        self.file_path = None
        self.df_preview = None
        self.layout_detectado = None

        # --- ESTILOS ---
        style = ttk.Style()
        style.theme_use('clam')
        style.configure("Treeview", background="white", foreground="black", rowheight=30, fieldbackground="white",
                        font=("Segoe UI", 10))
        style.configure("Treeview.Heading", font=("Segoe UI", 10, "bold"), background="#DFE6E9", foreground="#2D3436")
        style.map("Treeview", background=[('selected', COLORS['accent_blue'])])

        # --- LAYOUT DO CABEÇALHO ---
        header_frame = tk.Frame(root, bg=COLORS["primary"], height=80)
        header_frame.pack(fill="x", side="top")
        header_frame.pack_propagate(False)
        tk.Label(header_frame, text="ORGANIZADOR DE TRANSPORTES", bg=COLORS["primary"], fg=COLORS["text_light"],
                 font=("Segoe UI", 18, "bold")).pack(side="left", padx=20, pady=20)
        tk.Label(header_frame, text="v26.00 Final", bg=COLORS["primary"], fg="#95A5A6",
                 font=("Segoe UI", 10)).pack(side="right", padx=20, pady=25)

        # --- ÁREA DE CONTROLE ---
        control_frame = tk.Frame(root, bg=COLORS["card_bg"], bd=1, relief="solid")
        control_frame.pack(fill="x", padx=20, pady=20)

        # Linha 1: Seleção
        row1 = tk.Frame(control_frame, bg=COLORS["card_bg"])
        row1.pack(fill="x", padx=20, pady=15)
        self.btn_select = tk.Button(row1, text="📂 Selecionar Arquivo", bg=COLORS["accent_blue"], fg="white",
                                    font=("Segoe UI", 10, "bold"), relief="flat", padx=15, pady=5, cursor="hand2",
                                    command=self.selecionar_arquivo)
        self.btn_select.pack(side="left")
        self.lbl_filename = tk.Label(row1, text="Nenhum arquivo selecionado", bg=COLORS["card_bg"], fg="#7F8C8D",
                                     font=("Segoe UI", 10, "italic"))
        self.lbl_filename.pack(side="left", padx=15)
        ttk.Separator(control_frame, orient='horizontal').pack(fill='x', padx=20)

        # Linha 2: Ações e Status
        row2 = tk.Frame(control_frame, bg=COLORS["card_bg"])
        row2.pack(fill="x", padx=20, pady=15)
        self.lbl_detect_icon = tk.Label(row2, text="⚪", bg=COLORS["card_bg"], font=("Segoe UI", 14))
        self.lbl_detect_icon.pack(side="left")
        self.lbl_detect_text = tk.Label(row2, text="Aguardando...", bg=COLORS["card_bg"], fg="#7F8C8D",
                                        font=("Segoe UI", 10, "bold"))
        self.lbl_detect_text.pack(side="left", padx=5)

        tk.Frame(row2, bg=COLORS["card_bg"], width=30).pack(side="left")  # Espaçador

        self.btn_process = tk.Button(row2, text="⚙️ Processar", bg=COLORS["accent_orange"], fg="white",
                                     font=("Segoe UI", 10, "bold"), relief="flat", padx=15, pady=5, cursor="hand2",
                                     state="disabled", command=self.processar_dados)
        self.btn_process.pack(side="left", padx=5)

        self.btn_save_text = tk.StringVar()
        self.btn_save_text.set("💾 3. Salvar (Aguardando arquivo...)")
        self.btn_save = tk.Button(row2, textvariable=self.btn_save_text, bg=COLORS["accent_green"], fg="white",
                                  font=("Segoe UI", 10, "bold"), relief="flat", padx=15, pady=5, cursor="hand2",
                                  state="disabled", command=self.salvar_sequencial)
        self.btn_save.pack(side="left", padx=5)

        self.btn_reset = tk.Button(row2, text="↻ Reset", bg=COLORS["card_bg"], fg=COLORS["accent_red"],
                                   font=("Segoe UI", 9), relief="flat", bd=0, cursor="hand2",
                                   command=self.resetar_contador_manual)
        self.btn_reset.pack(side="right")

        # --- TABELA DE PREVIEW ---
        data_frame = tk.Frame(root, bg=COLORS["secondary"])
        data_frame.pack(fill="both", expand=True, padx=20, pady=(0, 20))
        tk.Label(data_frame, text="Pré-visualização:", bg=COLORS["secondary"], fg="#7F8C8D",
                 font=("Segoe UI", 10, "bold")).pack(anchor="w", pady=(0, 5))

        tree_scroll_y = ttk.Scrollbar(data_frame)
        tree_scroll_y.pack(side="right", fill="y")
        tree_scroll_x = ttk.Scrollbar(data_frame, orient="horizontal")
        tree_scroll_x.pack(side="bottom", fill="x")

        self.tree = ttk.Treeview(data_frame, columns=("col1"), show="headings", yscrollcommand=tree_scroll_y.set,
                                 xscrollcommand=tree_scroll_x.set)
        tree_scroll_y.config(command=self.tree.yview)
        tree_scroll_x.config(command=self.tree.xview)

        self.tree.pack(fill="both", expand=True)
        self.tree.tag_configure('oddrow', background="white")
        self.tree.tag_configure('evenrow', background="#F7F9F9")

        # --- RODAPÉ ---
        status_frame = tk.Frame(root, bg="#BDC3C7", height=25)
        status_frame.pack(fill="x", side="bottom")
        self.lbl_status = tk.Label(status_frame, text=" Iniciando interface...", bg="#BDC3C7", fg="#2C3E50",
                                   font=("Segoe UI", 9))
        self.lbl_status.pack(side="left", padx=10)

        # --- THREAD DE CARREGAMENTO ---
        threading.Thread(target=self.carregar_libs_pesadas, daemon=True).start()

    def carregar_libs_pesadas(self):
        global pd
        try:
            self.root.after(0, lambda: self.lbl_status.config(text=" Carregando núcleo de dados..."))
            import pandas as pandas_lib
            pd = pandas_lib
            self.libs_carregadas = True
            self.root.after(0, lambda: self.lbl_status.config(text=" Pronto."))
        except Exception as e:
            self.root.after(0, lambda: messagebox.showerror("Erro", f"Falha libs: {e}"))

    def verificar_libs(self):
        if not self.libs_carregadas:
            messagebox.showinfo("Carregando", "O sistema está otimizando a inicialização.\nAguarde...")
            return False
        return True

    # >>> HELPERS (NOVO) <<<
    def _clean_nf(self, val, split_hifen=False):
        """Limpa valores de Nota Fiscal/Documento para manter apenas os últimos 6 dígitos."""
        s = str(val).strip()
        if not s or s.lower() in ['nan', 'none', 'nat']: return ""
        if s.endswith('.0'): s = s[:-2]
        
        # Tratamento especial para TNT (hífens) - Só aplica se solicitado
        if split_hifen and '-' in s:
             s = s.split('-')[0].strip()

        s_numeros = re.sub(r'\D', '', s)
        if s_numeros:
            return s_numeros.zfill(6)[-6:]
        return ""


    def _fmt_dt_safe(self, val):
        """Formata datas de forma robusta usando Pandas."""
        if pd.isna(val) or str(val).strip() == '' or str(val).lower() in ['nan', 'none', 'nat']: return ""
        try:
            s = str(val).strip()
            # Tenta converter direto (dayfirst=True para formato BR)
            dt = pd.to_datetime(s, dayfirst=True, errors='coerce')
            
            if pd.notna(dt):
                return dt.strftime("%d/%m/%Y 00:00")
            return ""
        except:
            return ""


    # >>> FILTROS GLOBAIS <<<
    def _filtrar_valores_zerados(self, df_origem, df_destino):
        try:
            col_valor = next(
                (c for c in df_origem.columns if
                 any(x in str(c).upper() for x in ["VALOR", "VLR", "VR", "TOTAL", "MERCADORIA", "AMOUNT"])),
                None
            )
            if not col_valor: return df_destino

            def clean_money(val):
                if pd.isna(val) or str(val).strip() == '': return 0.0
                try:
                    s = str(val).strip()
                    s = re.sub(r'[^\d,]', '', s)
                    s = s.replace(',', '.')
                    return float(s)
                except:
                    return 0.0

            valores_limpos = df_origem[col_valor].apply(clean_money)
            if len(valores_limpos) == len(df_destino):
                return df_destino[valores_limpos.values > 0]
            else:
                return df_destino
        except Exception as e:
            print(f"Aviso filtro valor: {e}")
            return df_destino

    def _filtrar_linhas_sem_data(self, df):
        try:
            df = df.fillna("")
            mask_tem_prev = df["Data de Previsão de Entrega"].astype(str).str.strip() != ""
            mask_tem_ent = df["Data Entrega"].astype(str).str.strip() != ""
            df_final = df[mask_tem_prev | mask_tem_ent].copy()
            return df_final
        except Exception as e:
            print(f"Aviso filtro datas: {e}")
            return df

    # >>> LEITURA INTELIGENTE (CORRIGIDA: HEADER=NONE) <<<
    def ler_arquivo_inteligente(self):
        try:
            try:
                xls = pd.ExcelFile(self.file_path)
            except:
                # Se não for Excel, tenta CSV
                return pd.read_csv(self.file_path, sep=None, encoding='latin1', engine='python', header=None)

            # Scanner de Abas
            for sheet in xls.sheet_names:
                try:
                    # Lê 20 linhas para verificar se é a aba certa
                    df_check = pd.read_excel(self.file_path, sheet_name=sheet, nrows=20, header=None)

                    texto_completo = df_check.to_string().upper()

                    # Procura palavras-chave de cabeçalho
                    if "N.FISCAL" in texto_completo or "NFISCAL" in texto_completo or "NOTA FISCAL" in texto_completo or "DESTINATARIO" in texto_completo:
                        # ACHOU! Retorna com header=None para não perder a primeira linha
                        return pd.read_excel(self.file_path, sheet_name=sheet, header=None)
                except:
                    continue

            # Fallback
            return pd.read_excel(self.file_path, sheet_name=0, header=None)

        except Exception as e:
            try:
                return pd.read_csv(self.file_path, sep=None, encoding='latin1', engine='python', header=None)
            except:
                raise Exception(f"Erro Crítico na leitura: {e}")

    def identificar_layout(self, path):
        nome_arq = os.path.basename(path).upper()

        if "LISTA" in nome_arq and "CARGAS" in nome_arq: return "LISTA_CARGAS"
        if "EXCELLENCE" in nome_arq: return "TXT_EXCELLENCE"
        if "LT" in nome_arq or "DONIZETE" in nome_arq: return "LT"

        if "SOLISTICA" in nome_arq: return "SOLISTICA"

        if "ALFA" in nome_arq: return "ALFA"
        if "AGE" in nome_arq or "MH" in nome_arq: return "AGE"
        if "TNT" in nome_arq: return "TNT"

        if not self.verificar_libs(): return "AGUARDANDO_LIBS"

        content_upper = ""
        try:
            if path.lower().endswith(('.xls', '.xlsx')):
                try:
                    xls = pd.ExcelFile(path)
                    for sheet in xls.sheet_names:
                        df_temp = pd.read_excel(xls, sheet_name=sheet, nrows=20, header=None)
                        content_upper += df_temp.to_string().upper() + " "
                        if "N.FISCAL" in content_upper or "CTRC" in content_upper: break
                except:
                    pass

            if not content_upper:
                try:
                    with open(path, 'r', encoding='latin1', errors='ignore') as f:
                        content_upper = f.read(5000).upper()
                except:
                    pass

            if not content_upper: return "ERRO_LEITURA"

            if "EXCELLENCE" in content_upper and "NFISCAL" in content_upper: return "TXT_EXCELLENCE"
            if "NRO.DOC" in content_upper or "NRO DOC" in content_upper: return "ALFA"

            if "DATA PREVISÃO RECALCULADA" in content_upper or ("SOLISTICA" in content_upper): return "SOLISTICA"

            tem_tnt = "NOTA" in content_upper and ("SERIE" in content_upper or "SÉRIE" in content_upper)
            if tem_tnt or "FIL. ORIGEM" in content_upper: return "TNT"

            tem_ctrc = "CTRC" in content_upper
            tem_nf = (
                        "N.FISCAL" in content_upper or "NFISCAL" in content_upper or "NOTAFISCAL" in content_upper or "NR_NFE" in content_upper)

            if "DON" in content_upper and tem_nf: return "LT"
            if (tem_ctrc and tem_nf): return "AGE"

            if (tem_ctrc and tem_nf): return "AGE"

            return "GENERICO"
        except Exception as e:
            print(f"Erro ao identificar: {e}")
            return "GENERICO"


    def selecionar_arquivo(self):
        filename = filedialog.askopenfilename(title="Selecione o arquivo",
                                              filetypes=[("Arquivos", "*.xls *.xlsx *.csv *.txt"), ("Todos", "*.*")])
        if filename:
            self.file_path = filename
            self.lbl_filename.config(text=os.path.basename(filename), fg=COLORS["text_dark"],
                                     font=("Segoe UI", 10, "bold"))
            self.layout_detectado = self.identificar_layout(self.file_path)
            if self.layout_detectado == "AGUARDANDO_LIBS":
                self.lbl_detect_text.config(text="Carregando sistema...", fg="orange")
                self.root.after(1000, lambda: self.selecionar_arquivo_retry(filename))
                return
            self._aplicar_layout_config()

    def selecionar_arquivo_retry(self, filename):
        if not self.libs_carregadas:
            self.root.after(1000, lambda: self.selecionar_arquivo_retry(filename))
            return
        self.layout_detectado = self.identificar_layout(filename)
        self._aplicar_layout_config()

    def _aplicar_layout_config(self):
        # Lista de layouts VÁLIDOS que o sistema sabe processar
        layouts_validos = ["ALFA", "TNT", "LT", "AGE", "TXT_EXCELLENCE", "LISTA_CARGAS", "SOLISTICA", "GENERICO"]

        if self.layout_detectado in layouts_validos:
            # Mensagem genérica para todos os layouts reconhecidos - Preserva lógica interna, mas esconde nome
            self.configurar_status("Detectado", "✅", COLORS["accent_yellow"])
            self.btn_process.config(state="normal", bg=COLORS["accent_yellow"])
            self.btn_save.config(state="disabled", bg="#95A5A6")

        else:
            # Qualquer outra coisa que não seja um layout válido (Erro)
            self.lbl_detect_text.config(text=f"Erro ({self.layout_detectado})", fg="red")
            self.lbl_detect_icon.config(text="❌", fg="red")
            self.btn_process.config(state="disabled", bg="#95A5A6");
            self.btn_save.config(state="disabled", bg="#95A5A6")



    def configurar_status(self, texto, icone, cor):
        self.lbl_detect_text.config(text=f"Layout: {texto}", fg=cor)
        self.lbl_detect_icon.config(text=icone, fg=cor)
        self.btn_process.config(bg=cor)

    def processar_dados(self):
        if not self.verificar_libs(): return
        self.lbl_status.config(text=f"Processando {self.layout_detectado}...")
        self.root.update_idletasks()
        try:
            df_limpo = None

            if self.layout_detectado == "ALFA":
                df_limpo = self._limpar_alfa()
            elif self.layout_detectado == "TNT":
                df_limpo = self._limpar_tnt_smart()
            elif self.layout_detectado in ["LT", "AGE", "MH", "GENERICO"]:
                df_limpo = self._limpar_generico()
            elif self.layout_detectado == "TXT_EXCELLENCE":
                df_limpo = self._limpar_txt_excellence()
            elif self.layout_detectado == "LISTA_CARGAS":
                df_limpo = self._limpar_lista_cargas()
            elif self.layout_detectado == "SOLISTICA":
                df_limpo = self._limpar_solistica()


            if df_limpo is not None and not df_limpo.empty:
                # Aplica o filtro de datas vazias no final
                df_limpo = self._filtrar_linhas_sem_data(df_limpo)

                if df_limpo.empty:
                    self.lbl_status.config(text="Vazio pós-filtro.")
                    messagebox.showwarning("Aviso", "Notas encontradas, mas nenhuma possui data válida.")
                    return

                df_limpo.reset_index(drop=True, inplace=True)
                df_limpo.insert(0, "ITEM", range(1, len(df_limpo) + 1))
                self.df_preview = df_limpo
                self.atualizar_tabela(df_limpo)

                self.btn_save_text.set(self.get_texto_botao_salvar())
                self.btn_save.config(state="normal", bg=COLORS["accent_green"])
                self.lbl_status.config(text=f"Sucesso! {len(df_limpo)} linhas prontas.")
                messagebox.showinfo("Processado", f"{len(df_limpo)} linhas extraídas com sucesso.")
            else:
                self.lbl_status.config(text="Vazio.")
                messagebox.showwarning("Aviso", "Nenhum dado válido encontrado (verifique filtros de valor/data).")
        except Exception as ex:
            self.lbl_status.config(text="Erro.")
            messagebox.showerror("Erro Detalhado", f"Ocorreu um erro no processamento:\n{str(ex)}")

    def atualizar_tabela(self, df):
        self.tree.delete(*self.tree.get_children())
        cols = list(df.columns)
        self.tree["columns"] = cols
        for col in cols:
            self.tree.heading(col, text=col.upper())
            if col == "ITEM":
                self.tree.column(col, width=50, minwidth=50, stretch=False, anchor="center")
            elif col == "DATA DE PREVISÃO DE ENTREGA":
                self.tree.column(col, width=250, minwidth=200, stretch=True, anchor="center")
            elif col == "NR. DOC.":
                self.tree.column(col, width=150, minwidth=100, stretch=True, anchor="center")
            else:
                self.tree.column(col, width=180, minwidth=120, stretch=True, anchor="center")
        for i, row in enumerate(df.iterrows()):
            tag = 'evenrow' if i % 2 == 0 else 'oddrow'
            self.tree.insert("", "end", values=list(row[1]), tags=(tag,))
        self.lbl_status.config(text=f" Visualizando {len(df)} linhas.")

    # Removed old fmt_dt


    # --- FUNÇÕES DE LIMPEZA ---

    def _limpar_lista_cargas(self):
        try:
            df = pd.read_csv(self.file_path, sep=None, engine='python', encoding='latin1')
        except:
            try:
                df = pd.read_excel(self.file_path)
            except Exception as e:
                raise Exception(f"Não foi possível ler o arquivo Lista Cargas: {e}")

        df.columns = [str(c).upper().strip() for c in df.columns]

        col_nota = next(
            (c for c in df.columns if any(x in c for x in ['NOTA', 'NOTA_FISCAL', 'NF', 'DOCUMENTO', 'NR_NOTA'])), None)
        col_prev = next((c for c in df.columns if 'PREV' in c), None)
        col_ent = next((c for c in df.columns if 'ENTREGA' in c and 'PREV' not in c), None)

        if not col_ent:
            col_ent = next((c for c in df.columns if any(x in c for x in ['REALIZ', 'BAIXA'])), None)

        if not col_nota:
            raise Exception("Não encontrei a coluna de Nota Fiscal/Documento no arquivo.")

        df_final = pd.DataFrame()

        df_final["Nr. Doc."] = df[col_nota].apply(self._clean_nf)
        df_final["Data de Previsão de Entrega"] = df[col_prev].apply(self._fmt_dt_safe) if col_prev else ""
        df_final["Data Entrega"] = df[col_ent].apply(self._fmt_dt_safe) if col_ent else ""


        df_final = df_final[df_final["Nr. Doc."].astype(bool)]
        df_final = df_final[df_final["Nr. Doc."] != "000000"]

        return self._filtrar_valores_zerados(df, df_final[["Nr. Doc.", "Data de Previsão de Entrega", "Data Entrega"]])

    # >>> FUNÇÃO GENÉRICA/MH COM RESGATE DE LINHA <<<
    def _limpar_generico(self):

        df = self.ler_arquivo_inteligente()

        if len(df.columns) < 2:
            try:
                df = pd.read_csv(self.file_path, sep=None, encoding='latin1', engine='python', header=None)
            except:
                pass

        header_idx = None
        keywords_header = [
            "N.FISCAL", "NFISCAL", "NOTA FISCAL", "NR.NOTA", "N. NOTA",
            "NR_NFE", "NOTA FISCAL", "DOC.", "DOCUMENTO",
            "NOTAFISCAL", "CTRC", "REMETENTE", "DESTINATARIO",
            "NRO.DOC", "NRO DOC", "DT.ENTREGA", "DT.EMTREGA",
            "NUMERO DOCUMENTO", "NÚMERO DOCUMENTO"
        ]

        # Busca header na tabela bruta
        for i, row in df.head(50).iterrows():
            row_str = " ".join([str(val).upper() for val in row.values])
            if any(k in row_str for k in keywords_header):
                header_idx = i
                break

        if header_idx is not None:
            # Pega a linha bruta do cabeçalho
            raw_header = df.iloc[header_idx]

            clean_headers = []
            hidden_row_data = []
            has_hidden_data = False

            # >>> DESMESCLAR: Separa "Titulo" de "Dado" que estão na mesma celula <<<
            for val in raw_header:
                s = str(val).strip()

                # Se tiver ENTER, corta!
                if '\n' in s:
                    parts = s.rsplit('\n', 1)  # Separa pelo último Enter
                    clean_headers.append(parts[0].strip())  # Título
                    hidden_row_data.append(parts[1].strip())  # Dado da linha 1
                    has_hidden_data = True
                elif '\r' in s:
                    parts = s.rsplit('\r', 1)
                    clean_headers.append(parts[0].strip())
                    hidden_row_data.append(parts[1].strip())
                    has_hidden_data = True
                else:
                    clean_headers.append(s)
                    hidden_row_data.append(None)

            # Aplica nomes
            df.columns = clean_headers

            # Pega o resto dos dados
            df_rest = df.iloc[header_idx + 1:].copy()

            # Se achou dados escondidos, adiciona eles de volta!
            if has_hidden_data:
                df_hidden = pd.DataFrame([hidden_row_data], columns=clean_headers)
                df = pd.concat([df_hidden, df_rest], ignore_index=True)
            else:
                df = df_rest.reset_index(drop=True)

        new_cols = []
        for c in df.columns:
            s = str(c).upper().strip()
            # Limpeza de segurança extra
            s = s.replace('"', '').replace("'", "").replace('.', '')
            s = s.replace('Ï»¿', '').replace('ï»¿', '')
            new_cols.append(s)
        df.columns = new_cols

        # Sistema de Prioridade para encontrar a Nota Fiscal
        priority_groups = [
            ["NOTA FISCAL", "NOTAFISCAL", "NFISCAL", "DANFE", "NFE", " NR NF", "NRNF", "NNOTA"],
            ["NRNOTA", "N NOTA", "NUMERO NOTA", "NRNFE"],
            ["DOCUMENTO", "DOC", "NRODOC", "NUMERODOCUMENTO", "NÚMERODOCUMENTO"]
        ]


        col_nf = None
        for keywords in priority_groups:
            # Tenta encontrar uma coluna que contenha algum termo deste grupo
            # A verificação 'if k in c' checa se a palavra chave está contida no nome da coluna
            found = next((c for c in df.columns if any(k in c for k in keywords)), None)
            if found:
                col_nf = found
                break


        col_prev = next((c for c in df.columns if "PREV" in c), None)
        col_data = next((c for c in df.columns if ("ENTREGA" in c or "EMTREGA" in c) and "PREV" not in c), None)

        if not col_data:
            col_data = next((c for c in df.columns if "DATA" in c and (
                        "BAIXA" in c or "REALIZ" in c or "ULT" in c or "ÚLT" in c or "OCORR" in c)), None)

        if not col_nf:
            cols_encontradas = ", ".join(list(df.columns))
            raise Exception(f"Coluna de Nota Fiscal não encontrada.\nColunas limpas: [{cols_encontradas}]")

        df_final = pd.DataFrame()

        df_final["Nr. Doc."] = df[col_nf].apply(self._clean_nf)

        series_prev = df[col_prev].apply(self._fmt_dt_safe) if col_prev else pd.Series([""] * len(df))
        series_ent = df[col_data].apply(self._fmt_dt_safe) if col_data else pd.Series([""] * len(df))

        df_final["Data Entrega"] = series_ent
        df_final["Data de Previsão de Entrega"] = series_prev


        mask_sem_prev = df_final["Data de Previsão de Entrega"] == ""
        mask_com_ent = df_final["Data Entrega"] != ""
        df_final.loc[mask_sem_prev & mask_com_ent, "Data de Previsão de Entrega"] = df_final.loc[
            mask_sem_prev & mask_com_ent, "Data Entrega"]

        df_final = df_final[df_final["Nr. Doc."].astype(bool)]
        df_final = df_final[df_final["Nr. Doc."] != "000000"]
        df_final = df_final[~df_final["Nr. Doc."].str.upper().isin(["NRODOC", "NUMERODOCUMENTO"])]

        return self._filtrar_valores_zerados(df, df_final[["Nr. Doc.", "Data de Previsão de Entrega", "Data Entrega"]])

    def _limpar_tnt_smart(self):
        try:
            try:
                df_raw = pd.read_csv(self.file_path, sep=None, engine='python', header=None, nrows=20)
            except:
                df_raw = pd.read_excel(self.file_path, header=None, nrows=20)
        except Exception as e:
            raise Exception(f"Erro ao ler TNT: {e}")
        header_row_idx = None
        for idx, row in df_raw.iterrows():
            row_str = " ".join([str(val).upper() for val in row.values])
            if "NOTA" in row_str and ("SERIE" in row_str or "SÉRIE" in row_str):
                header_row_idx = idx
                break
        if header_row_idx is None:
            raise Exception("Não encontrei a linha de cabeçalho 'NOTA/SERIE'.")
        try:
            try:
                df = pd.read_csv(self.file_path, sep=None, engine='python', skiprows=header_row_idx)
            except:
                df = pd.read_excel(self.file_path, skiprows=header_row_idx)
        except:
            raise Exception("Erro ao recarregar TNT.")
        df.columns = df.columns.str.strip().str.upper()
        col_nota = next((c for c in df.columns if "NOTA" in c and ("SERIE" in c or "SÉRIE" in c)), None)
        if not col_nota:
            raise Exception(f"Coluna NOTA/SERIE não encontrada.")
        df_final = pd.DataFrame()

        df_final["Nr. Doc."] = df[col_nota].apply(lambda x: self._clean_nf(x, split_hifen=True))

        col_ent = next((c for c in df.columns if "DATA" in c and "FINALIZA" in c), None)
        col_prev = next((c for c in df.columns if "PREVIS" in c), None)
        df_final["Data Entrega"] = df[col_ent].apply(self._fmt_dt_safe) if col_ent else ""
        df_final["Data de Previsão de Entrega"] = df[col_prev].apply(self._fmt_dt_safe) if col_prev else ""

        df_final = df_final[df_final["Nr. Doc."].astype(bool)]

        return self._filtrar_valores_zerados(df, df_final[
            ["Nr. Doc.", "Data de Previsão de Entrega", "Data Entrega"]].fillna(""))

    def _limpar_txt_excellence(self):
        with open(self.file_path, 'r', encoding='latin1') as f:
            lines = f.readlines()
        dados = [];
        ano_atual = datetime.now().year
        for linha in lines:
            if "NFISCAL" in linha or "EXCELLENCE" in linha: continue
            match = re.search(r'\s+(\d{4,9})\s+.*(\d{2}/\d{2})\s+(\d{2}/\d{2})', linha)
            if match:
                nota, prev, ent = match.groups();
                nota_final = nota.zfill(6)[-6:]
                try:
                    dt_prev = self._fmt_dt_safe(f"{prev}/{ano_atual}")
                    dt_ent = self._fmt_dt_safe(f"{ent}/{ano_atual}")
                    dados.append([nota_final, dt_prev, dt_ent])

                except:
                    continue
        return pd.DataFrame(dados, columns=["Nr. Doc.", "Data de Previsão de Entrega", "Data Entrega"])

    def _limpar_alfa(self):
        try:
            df = pd.read_csv(self.file_path, header=None, sep=',', encoding='latin1', engine='python')
        except:
            df = pd.read_excel(self.file_path, header=None)
        cabecalho_idx = None;
        colunas_map = {};
        colunas_busca = {"Nro.Doc": "Nro.Doc", "Dt.Emtrega": "Dt.Emtrega"}
        for idx, row in df.iterrows():
            row_str = [str(v).strip() for v in row.values]
            if "Nro.Doc" in row_str:
                cabecalho_idx = idx
                for k, v in colunas_busca.items():
                    if k in row_str: colunas_map[v] = row_str.index(k)
                break
        if cabecalho_idx is None: raise Exception("Layout ALFA inválido.")
        df_dados = df.iloc[cabecalho_idx + 1:].copy();
        df_final = pd.DataFrame()
        if "Nro.Doc" in colunas_map: df_final["Nr. Doc."] = df_dados.iloc[:, colunas_map["Nro.Doc"]]
        if "Dt.Emtrega" in colunas_map: df_final["Data Entrega"] = df_dados.iloc[:, colunas_map["Dt.Emtrega"]]
        df_final = df_final[df_final["Nr. Doc."].notna()]
        df_final = df_final[~df_final["Nr. Doc."].astype(str).str.contains("Nro.Doc")]

        df_final["Nr. Doc."] = df_final["Nr. Doc."].apply(self._clean_nf)
        df_final["Data de Previsão de Entrega"] = df_final["Data Entrega"].apply(self._fmt_dt_safe)
        df_final["Data Entrega"] = df_final["Data Entrega"].apply(self._fmt_dt_safe)


        return self._filtrar_valores_zerados(df_dados,
                                             df_final[["Nr. Doc.", "Data de Previsão de Entrega", "Data Entrega"]])

    def _limpar_solistica(self):
        try:
            # Tenta ler Excel ou CSV
            try:
                df = pd.read_excel(self.file_path)
            except:
                df = pd.read_csv(self.file_path, sep=None, engine='python', encoding='latin1')
        except Exception as e:
            raise Exception(f"Erro ao ler Solistica: {e}")

        # Padroniza colunas para maiúsculo
        df.columns = [str(c).upper().strip() for c in df.columns]

        # Mapeamento direto (O arquivo Solistica é bem organizado)
        # Nota Fiscal -> "NOTA FISCAL"
        # Previsão -> "DATA PREVISÃO"
        # Entrega -> "DATA ENTREGA"

        col_nf = "NOTA FISCAL"
        col_prev = "DATA PREVISÃO"
        col_ent = "DATA ENTREGA"

        # Verificação básica se as colunas existem
        if col_nf not in df.columns:
            raise Exception("Layout Solistica mudou? Não achei a coluna 'NOTA FISCAL'.")

        df_final = pd.DataFrame()

        # Limpeza do número da nota
        df_final["Nr. Doc."] = df[col_nf].apply(self._clean_nf)

        # Datas
        df_final["Data de Previsão de Entrega"] = df[col_prev].apply(self._fmt_dt_safe)
        df_final["Data Entrega"] = df[col_ent].apply(self._fmt_dt_safe)


        # Remove vazios
        df_final = df_final[df_final["Nr. Doc."].astype(bool)]
        df_final = df_final[df_final["Nr. Doc."] != "000000"]

        # Aplica o filtro global de valores zerados e datas vazias
        df_final = self._filtrar_valores_zerados(df,
                                                 df_final[["Nr. Doc.", "Data de Previsão de Entrega", "Data Entrega"]])
        return df_final

    def get_proximo_numero(self):
        try:
            if os.path.exists(CONFIG_FILE):
                with open(CONFIG_FILE, "r") as f: return int(f.read().strip()) + 1
            return 1
        except:
            return 1

    def salvar_numero_atual(self, numero):
        try:
            with open(CONFIG_FILE, "w") as f:
                f.write(str(numero))
        except:
            pass

    def get_texto_botao_salvar(self):
        agora = datetime.now().strftime("%d-%m-%Y_%Hh%M")

        if self.file_path:
            nome_original = os.path.splitext(os.path.basename(self.file_path))[0]
            nome_limpo = re.sub(r'[^\w\-]', '_', nome_original)
            if len(nome_limpo) > 40: nome_limpo = nome_limpo[:40]
            nome_final = f"Logistica_{nome_limpo}_{agora}.csv"
        else:
            nome_final = "Logistica_Geral.csv"

        return f"💾 3. Salvar '{nome_final}'"

    def salvar_sequencial(self):
        if self.df_preview is None: return
        numero = self.get_proximo_numero()

        agora = datetime.now().strftime("%d-%m-%Y_%Hh%M")

        if self.file_path:
            nome_original = os.path.splitext(os.path.basename(self.file_path))[0]
            nome_limpo = re.sub(r'[^\w\-]', '_', nome_original)
            if len(nome_limpo) > 40: nome_limpo = nome_limpo[:40]
            nome_arq = f"Logistica_{nome_limpo}_{agora}.csv"
        else:
            nome_arq = f"Logistica_Geral_{agora}.csv"

        pasta = os.path.join(os.path.expanduser("~"), "Downloads")
        caminho = os.path.join(pasta, nome_arq)
        try:
            df_para_salvar = self.df_preview.copy()
            if "ITEM" in df_para_salvar.columns: df_para_salvar = df_para_salvar.drop(columns=["ITEM"])

            if "Nr. Doc." in df_para_salvar.columns:
                df_para_salvar["Nr. Doc."] = df_para_salvar["Nr. Doc."].astype(str)

            df_para_salvar.to_csv(caminho, index=False, sep=';', encoding='utf-8-sig', header=False,
                                  quoting=csv.QUOTE_ALL)

            self.salvar_numero_atual(numero)

            self.btn_save_text.set(self.get_texto_botao_salvar())

            self.lbl_status.config(text=f"Salvo: {nome_arq}")
            self.reset_tela_pos_salvamento(caminho)
        except Exception as ex:
            messagebox.showerror("Erro", str(ex))

    def reset_tela_pos_salvamento(self, caminho):
        messagebox.showinfo("Sucesso", f"Salvo em:\n{caminho}")
        self.file_path = None;
        self.df_preview = None
        self.lbl_filename.config(text="Nenhum arquivo", fg="#7F8C8D", font=("Segoe UI", 10, "italic"))
        self.lbl_detect_text.config(text="Aguardando...", fg="#7F8C8D");
        self.lbl_detect_icon.config(text="⚪", fg="#7F8C8D")
        self.btn_process.config(state="disabled", bg="#95A5A6");
        self.btn_save.config(state="disabled", bg="#95A5A6")
        self.tree.delete(*self.tree.get_children());
        self.lbl_status.config(text="Pronto.")

    def resetar_contador_manual(self):
        if messagebox.askyesno("Reset", "Zerar contador?"):
            try:
                with open(CONFIG_FILE, "w") as f:
                    f.write("0")
                self.btn_save_text.set(self.get_texto_botao_salvar())
            except:
                pass


if __name__ == "__main__":
    root = tk.Tk()
    app = LogicApp(root)
    root.mainloop()