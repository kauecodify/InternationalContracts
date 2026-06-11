"""
conda install pandas folium customtkinter openpyxl geopy thefuzz python-Levenshtein

GeoMapper Inteligente

Aplicação desktop para geração automática de mapas interativos a partir de planilhas Excel ou CSV contendo endereços.

Leitura automática de arquivos XLSX, XLS e CSV.
Detecção inteligente da coluna de endereço.
Geocodificação de endereços completos utilizando OpenStreetMap/Nominatim.
Conversão automática de endereços em coordenadas geográficas.
Geração de mapa interativo HTML com marcadores.
Exibição de dados complementares da planilha em popups.
Interface gráfica moderna desenvolvida com CustomTkinter.
Barra de progresso durante o processamento.
Abertura automática do mapa gerado no navegador.

Fluxo:
Planilha -> Geocodificação -> Latitude/Longitude -> Mapa Interativo -> Exportação HTML

Aplicações:

Logística e roteirização.
Mapeamento de clientes.
Inteligência comercial.
Projetos ESG.
Planejamento territorial.
Monitoramento de ativos e operações.

alterações: em "def detectar_coluna_endereco" > linha 67

"""

import os
import webbrowser
import unicodedata
import pandas as pd
import folium
import customtkinter as ctk
import geopy

from tkinter import filedialog, messagebox
from thefuzz import process, fuzz
from geopy.geocoders import Nominatim


def normalizar_texto(texto):
    if not isinstance(texto, str):
        texto = str(texto)

    texto = unicodedata.normalize(
        'NFKD',
        texto
    ).encode(
        'ASCII',
        'ignore'
    ).decode(
        'utf-8'
    )

    texto = ''.join(
        e for e in texto
        if e.isalnum() or e.isspace()
    )

    return texto.strip().lower()


def detectar_coluna_endereco(df):
    palavras = [
        "endereco",
        "endereço",
        "logradouro",
        "rua",
        "avenida",
        "cep",
        "localizacao",
        "localização",
        "address"
    ]

    colunas = [str(col).lower() for col in df.columns]

    melhor, score = process.extractOne(
        " ".join(palavras),
        colunas,
        scorer=fuzz.token_sort_ratio
    )

    if score > 40:
        return df.columns[colunas.index(melhor)]

    return None


def geocodificar(endereco, geolocator):
    try:
        resultado = geolocator.geocode(
            endereco,
            timeout=15
        )

        if resultado:
            return (
                resultado.latitude,
                resultado.longitude,
                resultado.address
            )

    except:
        pass

    return None


def gerar_mapa(
    arquivo,
    destino,
    callback
):
    if arquivo.endswith((".xlsx", ".xls")):
        df = pd.read_excel(arquivo)

    elif arquivo.endswith(".csv"):
        df = pd.read_csv(
            arquivo,
            sep=None,
            engine="python"
        )

    else:
        raise ValueError(
            "Formato inválido."
        )

    coluna = detectar_coluna_endereco(df)

    if not coluna:
        raise ValueError(
            "Coluna de endereço não encontrada."
        )

    geolocator = Nominatim(
        user_agent="geomapper_ai"
    )

    mapa = folium.Map(
        location=[-14.2350, -51.9253],
        zoom_start=4,
        tiles="CartoDB positron"
    )

    total = len(df)
    sucesso = 0
    falha = 0

    for i, linha in df.iterrows():

        endereco = str(
            linha[coluna]
        ).strip()

        if not endereco:
            continue

        resultado = geocodificar(
            endereco,
            geolocator
        )

        if resultado:

            lat, lon, endereco_real = resultado

            popup = f"<b>{endereco_real}</b><hr>"

            for c in df.columns:
                if pd.notna(linha[c]):
                    popup += (
                        f"<b>{c}</b>: "
                        f"{linha[c]}<br>"
                    )

            folium.Marker(
                location=[lat, lon],
                popup=folium.Popup(
                    popup,
                    max_width=400
                ),
                tooltip=endereco_real
            ).add_to(mapa)

            sucesso += 1

        else:
            falha += 1

        progresso = int(
            ((i + 1) / total) * 100
        )

        callback(progresso)

    arquivo_html = os.path.join(
        destino,
        "mapa_interativo.html"
    )

    mapa.save(arquivo_html)

    return (
        arquivo_html,
        sucesso,
        falha
    )


class GeoMapper:

    def __init__(self, root):

        self.root = root

        self.root.title(
            "GeoMapper"
        )

        self.root.geometry(
            "700x500"
        )

        ctk.set_appearance_mode(
            "dark"
        )

        ctk.set_default_color_theme(
            "blue"
        )

        frame = ctk.CTkFrame(root)

        frame.pack(
            fill="both",
            expand=True,
            padx=20,
            pady=20
        )

        self.planilha = ctk.CTkEntry(
            frame,
            height=40
        )

        self.planilha.pack(
            fill="x",
            padx=20,
            pady=10
        )

        ctk.CTkButton(
            frame,
            text="Selecionar Arquivo",
            command=self.escolher_arquivo
        ).pack(
            padx=20,
            pady=5
        )

        self.destino = ctk.CTkEntry(
            frame,
            height=40
        )

        self.destino.pack(
            fill="x",
            padx=20,
            pady=10
        )

        ctk.CTkButton(
            frame,
            text="Selecionar Pasta",
            command=self.escolher_pasta
        ).pack(
            padx=20,
            pady=5
        )

        self.barra = ctk.CTkProgressBar(
            frame
        )

        self.barra.pack(
            fill="x",
            padx=20,
            pady=20
        )

        self.barra.set(0)

        self.status = ctk.CTkLabel(
            frame,
            text=""
        )

        self.status.pack()

        self.botao = ctk.CTkButton(
            frame,
            text="Gerar Mapa",
            height=50,
            command=self.processar
        )

        self.botao.pack(
            fill="x",
            padx=20,
            pady=20
        )

    def escolher_arquivo(self):

        arquivo = filedialog.askopenfilename(
            filetypes=[
                (
                    "Planilhas",
                    "*.xlsx *.xls *.csv"
                )
            ]
        )

        if arquivo:

            self.planilha.delete(
                0,
                ctk.END
            )

            self.planilha.insert(
                0,
                arquivo
            )

            pasta = os.path.dirname(
                arquivo
            )

            self.destino.delete(
                0,
                ctk.END
            )

            self.destino.insert(
                0,
                pasta
            )

    def escolher_pasta(self):

        pasta = filedialog.askdirectory()

        if pasta:

            self.destino.delete(
                0,
                ctk.END
            )

            self.destino.insert(
                0,
                pasta
            )

    def atualizar(self, valor):

        self.barra.set(
            valor / 100
        )

        self.status.configure(
            text=f"{valor}%"
        )

        self.root.update()

    def processar(self):

        arquivo = self.planilha.get()

        pasta = self.destino.get()

        if not arquivo or not pasta:

            messagebox.showerror(
                "Erro",
                "Selecione arquivo e pasta."
            )

            return

        try:

            html, ok, erro = gerar_mapa(
                arquivo,
                pasta,
                self.atualizar
            )

            webbrowser.open(
                html
            )

            messagebox.showinfo(
                "Concluído",
                f"Mapeados: {ok}\nFalhas: {erro}"
            )

        except Exception as e:

            messagebox.showerror(
                "Erro",
                str(e)
            )


if __name__ == "__main__":

    root = ctk.CTk()

    app = GeoMapper(root)

    root.mainloop()
