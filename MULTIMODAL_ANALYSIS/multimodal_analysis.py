"""

conda install tkinter pandas folium pycities

"""

import os
import webbrowser
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pandas as pd
import folium

try:
    from pycities.cities import Cities
except ModuleNotFoundError:
    Cities = None

def gerar_mapa_pontos(caminho_planilha, caminho_salvamento):
    """
    Lê a planilha do usuário, localiza as coordenadas via pycities,
    cria marcadores visuais e abre o mapa no navegador do computador.
    """
    if Cities is None:
        raise ImportError("O módulo 'pycities' não está instalado no seu ambiente Python.")
        
    if caminho_planilha.endswith('.xlsx') or caminho_planilha.endswith('.xls'):
        df = pd.read_excel(caminho_planilha)
    elif caminho_planilha.endswith('.csv'):
        df = pd.read_csv(caminho_planilha, sep=None, engine='python')
    else:
        raise ValueError("Formato de arquivo inválido. Selecione uma planilha .xlsx ou .csv.")

    # force para identificar a coluna que contém o nome das cidades
    coluna_cidade = None
    for col in df.columns:
        if 'cidade' in col.lower() or 'municipio' in col.lower() or 'local' in col.lower():
            coluna_cidade = col
            break
            
    if not coluna_cidade:
        raise KeyError("Não foi encontrada nenhuma coluna com o nome 'Cidade' ou 'Municipio' na planilha.")

    # inicialização do Banco de Dados pycities
    base_cidades = Cities()
    mapa_interativo = folium.Map(location=[-15.7801, -47.9292], zoom_start=4)
    
    pontos_mapeados = 0

    for _, linha in df.iterrows():
        nome_busca = str(linha[coluna_cidade]).strip()
        
        # coordenadas exatas na base do pycities
        resultado = base_cidades.search(nome_busca)
        
        if resultado:
            if isinstance(resultado, list) and len(resultado) > 0: 
                resultado = resultado[0]
                
            # extract de atributos reais do objeto geográfico
            lat = getattr(resultado, 'latitude', None)
            lon = getattr(resultado, 'longitude', None)
            nome_oficial = getattr(resultado, 'name', nome_busca)
            
            if lat is None or lon is None:
                continue
                
            detalhes_popup = f"<b>{nome_oficial}</b><br>"
            for col in df.columns:
                if col != coluna_cidade:
                    detalhes_popup += f"{col}: {linha[col]}<br>"

            # adc alfinete/marcador ao mapa folium
            folium.Marker(
                location=[lat, lon],
                popup=folium.Popup(detalhes_popup, max_width=300),
                tooltip=nome_oficial,
                icon=folium.Icon(color='black', icon='info-sign')
            ).add_to(mapa_interativo)
            
            pontos_mapeados += 1

    if pontos_mapeados == 0:
        raise ValueError("Nenhuma cidade da planilha foi reconhecida ou localizada na base do pycities.")

    # save arq no HTML do Mapa
    caminho_html_mapa = os.path.join(caminho_salvamento, "visualizacao_pontos_pycities.html")
    mapa_interativo.save(caminho_html_mapa)
    
    # força a abertura do mapa no navegador padrão
    webbrowser.open(caminho_html_mapa)
    return caminho_html_mapa

class AplicativoMapaMapeamento:
    def __init__(self, root):
        self.root = root
        self.root.title("Mapeamento Geográfico Inteligente")
        self.root.geometry("520x420")
        self.root.configure(bg="#121212")
        self.root.resizable(False, False)

        self.style = ttk.Style()
        self.style.theme_use('default')
        self.style.configure('.', background='#121212', foreground='#FFFFFF')
        self.style.configure('TLabel', background='#121212', foreground='#E0E0E0', font=('Arial', 10))
        
        self.lbl_titulo = tk.Label(root, text="SISTEMA DE MAPEAMENTO DE TABELAS", bg="#121212", fg="#FFFFFF", font=("Arial", 13, "bold"))
        self.lbl_titulo.pack(pady=25)

        self.frame = tk.Frame(root, bg="#121212")
        self.frame.pack(padx=30, fill='x')

        self.lbl_planilha = ttk.Label(self.frame, text="Selecione sua Planilha (.xlsx ou .csv):")
        self.lbl_planilha.pack(anchor='w', pady=(0, 5))
        
        self.sheet_frame = tk.Frame(self.frame, bg="#121212")
        self.sheet_frame.pack(fill='x', pady=(0, 15))
        self.entry_planilha = tk.Entry(self.sheet_frame, bg="#1E1E1E", fg="#FFFFFF", insertbackground="white", bd=0, font=("Arial", 11))
        self.entry_planilha.pack(side='left', fill='x', expand=True, ipady=8)
        self.btn_procurar_sheet = tk.Button(self.sheet_frame, text="...", bg="#2A2A2A", fg="#FFFFFF", bd=0, width=4, command=self.selecionar_planilha)
        self.btn_procurar_sheet.pack(side='right', padx=(10, 0), ipady=5)

        self.lbl_caminho = ttk.Label(self.frame, text="Pasta para salvar o Mapa de Saída:")
        self.lbl_caminho.pack(anchor='w', pady=(0, 5))
        
        self.path_frame = tk.Frame(self.frame, bg="#121212")
        self.path_frame.pack(fill='x', pady=(0, 30))
        self.entry_caminho = tk.Entry(self.path_frame, bg="#1E1E1E", fg="#FFFFFF", insertbackground="white", bd=0, font=("Arial", 11))
        self.entry_caminho.pack(side='left', fill='x', expand=True, ipady=8)
        self.btn_procurar_dir = tk.Button(self.path_frame, text="...", bg="#2A2A2A", fg="#FFFFFF", bd=0, width=4, command=self.selecionar_pasta)
        self.btn_procurar_dir.pack(side='right', padx=(10, 0), ipady=5)

        self.btn_gerar = tk.Button(root, text="ABRIR MAPA INTERATIVO NO NAVEGADOR", bg="#FFFFFF", fg="#000000", activebackground="#E0E0E0", bd=0, font=("Arial", 10, "bold"), cursor="hand2", command=self.executar_mapeamento)
        self.btn_gerar.pack(fill='x', padx=30, ipady=12)

    def selecionar_planilha(self):
        arquivo = filedialog.askopenfilename(filetypes=[("Arquivos de Tabela", "*.xlsx *.xls *.csv")])
        if arquivo:
            self.entry_planilha.delete(0, tk.END)
            self.entry_planilha.insert(0, arquivo)

    def selecionar_pasta(self):
        pasta = filedialog.askdirectory()
        if pasta:
            self.entry_caminho.delete(0, tk.END)
            self.entry_caminho.insert(0, pasta)

    def executar_mapeamento(self):
        planilha = self.entry_planilha.get().strip()
        pasta_destino = self.entry_caminho.get().strip()

        if not planilha or not pasta_destino:
            messagebox.showwarning("Campos Vazios", "Por favor, selecione tanto o arquivo da planilha quanto a pasta de destino.")
            return

        try:
            self.root.config(cursor="watch")
            self.root.update()
            
            mapa_html = gerar_mapa_pontos(planilha, pasta_destino)
            
            messagebox.showinfo("Sucesso!", f"Mapa gerado e aberto no navegador!\n\nArquivo salvo em:\n{os.path.basename(mapa_html)}")
        except Exception as e:
            messagebox.showerror("Erro de Processamento", str(e))
        finally:
            self.root.config(cursor="")

if __name__ == "__main__":
    app_root = tk.Tk()
    app = AplicativoMapaMapeamento(app_root)
    app_root.mainloop()
