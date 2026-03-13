import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.patches as patches
from matplotlib.path import Path
import altair as alt
from fpdf import FPDF
import tempfile
import os
import base64
import requests
from datetime import datetime
import numpy as np

# ==============================================================================
# CONFIGURAÇÃO GERAL
# ==============================================================================
st.set_page_config(page_title="Resumo Executivo SPDA", page_icon="⚡", layout="wide")

st.markdown("""
<style>
    div[data-testid="stMetricValue"] { font-size: 20px; }
    .stSelectbox { margin-bottom: 20px; }
    .dataframe { font-size: 11px !important; }
    button[data-baseweb="tab"] { font-size: 16px; font-weight: bold; }
</style>
""", unsafe_allow_html=True)

CORES = {'OK': '#2ecc71', 'Alto': '#f1c40f', 'Parcialmente Aberto': '#e67e22', 'Aberto': '#e74c3c', 'Indefinido': '#95a5a6'}

# ==============================================================================
# UTILITÁRIOS
# ==============================================================================
LOGO_BASE64 = "" 
def salvar_logo_temporario():
    if LOGO_BASE64:
        try:
            with open("logo_temp.png", "wb") as f: f.write(base64.b64decode(LOGO_BASE64))
            return "logo_temp.png"
        except: pass
    if os.path.exists("logo.png"): return "logo.png"
    return None

def salvar_fig_temp(fig):
    with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as t:
        fig.savefig(t.name, bbox_inches='tight', dpi=150, transparent=True)
        return t.name

def baixar_foto_da_url(url):
    if not isinstance(url, str) or len(url) < 5: return None
    try:
        response = requests.get(url, timeout=4)
        if response.status_code == 200:
            suffix = '.jpg'
            if '.png' in url.lower(): suffix = '.png'
            with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as t:
                t.write(response.content)
                return t.name
    except: return None
    return None

# ==============================================================================
# 1. PROCESSAMENTO
# ==============================================================================
@st.cache_data
def carregar_dados(uploaded_file):
    if uploaded_file is None: return None
    try:
        if uploaded_file.name.endswith('.xlsx'): df = pd.read_excel(uploaded_file)
        else:
            try: df = pd.read_csv(uploaded_file)
            except: df = pd.read_csv(uploaded_file, sep=';')
        df.columns = [c.strip() for c in df.columns]
        for col in ['Result', 'Receptor', 'Side']:
            if col in df.columns: df[col] = df[col].astype(str).str.strip()
        return df
    except Exception as e: st.error(f"Erro: {e}"); return None

def processar_dataframe(df):
    df['Resistance_Num'] = pd.to_numeric(df['Resistance'], errors='coerce')
    df['Resistance_Num'] = df['Resistance_Num'].fillna(999.0)
    
    if 'Date' in df.columns and not df['Date'].dropna().empty:
        # Garante que a coluna seja datetime
        df['Date'] = pd.to_datetime(df['Date'])
        data_inicio = df['Date'].min().strftime('%d/%m/%Y')
        data_fim = df['Date'].max().strftime('%d/%m/%Y')
        if data_inicio == data_fim:
            data_periodo = data_inicio
        else:
            data_periodo = f"{data_inicio} a {data_fim}"
    else:
        data_periodo = datetime.today().strftime('%d/%m/%Y')

    if 'Side' not in df.columns: df['Side'] = ''
    df['Side'] = df['Side'].fillna('').astype(str).str.upper()
    
    mask_tip = df['Receptor'].astype(str).str.contains('Tip', case=False, na=False)
    df.loc[mask_tip, 'Side'] = 'PS'

    def classificar_receptor(row):
        res_txt = str(row['Result']).lower()
        val = row['Resistance_Num']
        if 'open' in res_txt or 'no con' in res_txt: return 'Aberto'
        if val > 250: return 'Alto'
        return 'OK'

    df['Status_Calc'] = df.apply(classificar_receptor, axis=1)
    df['Display_Val'] = df.apply(lambda r: "OPEN" if r['Status_Calc'] == 'Aberto' else f"{int(r['Resistance_Num'])}", axis=1)
    df['Receptor_Label'] = df['Receptor'] + " " + df['Side'] + " (" + df['Location'].astype(int).astype(str) + "m)"
    
    df_pivot = df.pivot_table(index=['Turbine', 'Blade Index', 'Blade Model'], columns='Receptor_Label', values='Display_Val', aggfunc='first').reset_index()
    
    def contar_anomalias(grupo):
        lista = grupo['Status_Calc'].tolist()
        return lista.count('Aberto') * 10 + lista.count('Alto')

    def calcular_status_pa(grupo):
        lista = grupo['Status_Calc'].tolist()
        if not lista: return 'Indefinido'
        total = len(lista); abertos = lista.count('Aberto'); altos = lista.count('Alto')
        if total > 0 and abertos == total: return 'Aberto'
        if abertos > 0: return 'Parcialmente Aberto'
        if altos > 0: return 'Alto'
        return 'OK'

    stt_pa = df.groupby(['Turbine', 'Blade Index', 'Blade Model']).apply(
        lambda x: pd.Series({
            'Status_Final': calcular_status_pa(x),
            'Qtd_Anomalias': contar_anomalias(x)
        })
    ).reset_index()
    
    df_final = pd.merge(df_pivot, stt_pa, on=['Turbine', 'Blade Index', 'Blade Model'], how='left')

    def calcular_status_turbina(grupo):
        lista = grupo['Status_Final'].tolist()
        if lista.count('Aberto') == len(lista) and len(lista)>0: return 'Aberto'
        if 'Aberto' in lista or 'Parcialmente Aberto' in lista: return 'Parcialmente Aberto'
        if 'Alto' in lista: return 'Alto'
        return 'OK'

    stt_tb = df_final.groupby('Turbine').apply(calcular_status_turbina).reset_index(name='Status_Turbina')
    df_final = pd.merge(df_final, stt_tb, on='Turbine', how='left')
    return df, df_final

def gerar_conclusao_texto(df_piv, df_raw):
    total_pas = len(df_piv)
    criticas = len(df_piv[df_piv['Status_Final'] == 'Aberto'])
    parciais = len(df_piv[df_piv['Status_Final'] == 'Parcialmente Aberto'])
    altas = len(df_piv[df_piv['Status_Final'] == 'Alto'])
    oks = len(df_piv[df_piv['Status_Final'] == 'OK'])
    
    texto = f"A campanha de inspeção analisou um total de {total_pas} pás.\n\n" \
            f"DIAGNÓSTICO GERAL:\n" \
            f"- {oks} pás ({oks/total_pas:.1%}) encontram-se em conformidade (OK).\n" \
            f"- {criticas} pás ({criticas/total_pas:.1%}) apresentam PERDA TOTAL DE CONTINUIDADE (Aberto).\n" \
            f"- {parciais} pás ({parciais/total_pas:.1%}) apresentam FALHA PARCIAL em um ou mais receptores.\n" \
            f"- {altas} pás ({altas/total_pas:.1%}) apresentam RESISTÊNCIA ALTA (> 250 mΩ).\n\n" \
            f"RECOMENDAÇÕES TÉCNICAS:\n"

    recs = []
    if criticas > 0: recs.append("1. PARA PÁS COM CIRCUITO ABERTO (CRÍTICO): Recomenda-se inspeção interna imediata.")
    if parciais > 0: recs.append("2. PARA FALHAS PARCIAIS: Verificar conexão entre receptor e cabo principal.")
    if altas > 0: recs.append("3. PARA RESISTÊNCIA ALTA: Realizar limpeza e reaperto das conexões.")
    
    if not recs: texto += "Nenhuma anomalia grave detectada."
    else: texto += "\n".join(recs)
    return texto

# ==============================================================================
# 2. VISUAIS
# ==============================================================================
def criar_shape_suave(comp, y_offset=0):
    raiz_width = 0.8; raiz_len = comp * 0.06
    max_chord_x = comp * 0.20; max_chord_top = 2.0; max_chord_bot = 1.5; tip_w = 0.05
    verts = [
        (0, y_offset + raiz_width), (raiz_len, y_offset + raiz_width),
        (max_chord_x, y_offset + max_chord_top), (comp * 0.6, y_offset + 0.9), (comp, y_offset + tip_w),
        (comp, y_offset - tip_w), (max_chord_x, y_offset - max_chord_bot),
        (raiz_len, y_offset - raiz_width), (0, y_offset - raiz_width), (0, y_offset + raiz_width),
    ]
    codes = [Path.MOVETO, Path.LINETO, Path.CURVE4, Path.CURVE4, Path.CURVE4, Path.LINETO, Path.LINETO, Path.LINETO, Path.LINETO, Path.CLOSEPOLY]
    return patches.PathPatch(Path(verts, codes), facecolor='#ecf0f1', edgecolor='#34495e', lw=1.5, zorder=1)

def desenhar_pa_individual(df_p, mod, idx, comp_input):
    fig, ax = plt.subplots(figsize=(10, 4), facecolor='none')
    if df_p.empty: ax.text(0.5,0.5,"N/A",ha='center'); ax.axis('off'); return fig
    comp = comp_input
    patch = criar_shape_suave(comp, y_offset=0); ax.add_patch(patch)
    ax.set_xlim(-2, comp+5); ax.set_ylim(-4, 4); ax.axis('off')
    ax.text(comp/2, 3.5, f"Pá {idx}", ha='center', fontweight='bold', color='#2c3e50')
    for _,r in df_p.iterrows():
        loc=r['Location']; st=r['Status_Calc']; res=str(r['Resistance']); side=str(r.get('Side',''))
        c=CORES.get(st,'gray'); y_pos = 0.6 if 'PS' in side else (-0.6 if 'SS' in side else 0)
        ax.scatter([loc], [y_pos], c=c, s=180, zorder=5, edgecolors='k', lw=0.5)
        ax.plot([loc,loc], [y_pos, y_pos*2.5], ls=':', color='gray', alpha=0.5)
        txt = "ABERTO" if st=='Aberto' else f"{res} mΩ"
        text_y = 2.5 if y_pos > 0 else -3.0
        bb=dict(boxstyle="round,pad=0.2",fc="white",ec=c,lw=1)
        ax.text(loc, text_y, f"{side}\n{txt}", ha='center', va='center', fontsize=7, bbox=bb, fontweight='bold')
    plt.close(fig); return fig

def desenhar_pa_estatistica(df_m, mod, comp_input):
    fig, ax = plt.subplots(figsize=(12, 9), facecolor='none')
    comp = comp_input
    y_ps = 3.5; y_ss = -3.5  
    
    patch_ps = criar_shape_suave(comp, y_offset=y_ps); patch_ps.set_facecolor('#f8f9fa'); ax.add_patch(patch_ps)
    ax.text(0, y_ps + 2.8, "Lado PS (Pressão)", fontsize=11, color='#34495e', fontweight='bold')
    
    patch_ss = criar_shape_suave(comp, y_offset=y_ss); patch_ss.set_facecolor('#f8f9fa'); ax.add_patch(patch_ss)
    ax.text(0, y_ss + 2.8, "Lado SS (Sucção)", fontsize=11, color='#34495e', fontweight='bold')
    
    ax.set_xlim(-5, comp+10); ax.set_ylim(-10, 10); ax.axis('off')
    ax.text(comp/2, 8.5, f"Distribuição de Status por Receptor: {mod}", ha='center', fontweight='bold', fontsize=14)
    
    status_order = ['OK', 'Alto', 'Parcialmente Aberto', 'Aberto']
    
    for loc in sorted(df_m['Location'].unique()):
        for y_off, side_label in [(y_ps, 'PS'), (y_ss, 'SS')]:
            subset = df_m[(df_m['Location'] == loc) & (df_m['Side'].str.contains(side_label))]
            if not subset.empty:
                counts = subset['Status_Calc'].value_counts(normalize=True)
                sizes = [counts.get(s, 0) for s in status_order if s in counts]
                cl = [CORES[s] for s in status_order if s in counts]
                
                if sizes:
                    ax_pie = ax.inset_axes([loc-2.5, y_off-2.5, 5, 5], transform=ax.transData)
                    wedges, _ = ax_pie.pie(sizes, colors=cl, radius=1, wedgeprops=dict(width=0.4, edgecolor='w'), startangle=90)
                    
                    for i, p in enumerate(wedges):
                        ang = (p.theta2 - p.theta1)/2. + p.theta1
                        y = np.sin(np.deg2rad(ang))
                        x = np.cos(np.deg2rad(ang))
                        if sizes[i] > 0:
                            percentage = f"{sizes[i]*100:.0f}%"
                            halign = "center" if abs(x) < 0.5 else ("left" if x > 0 else "right")
                            ax_pie.annotate(percentage, xy=(x, y), xytext=(1.35*x, 1.35*y),
                                            fontsize=7, fontweight='bold', arrowprops=dict(arrowstyle="-", color='gray', lw=0.7),
                                            ha=halign, va="center")
                    ax.text(loc, y_off-3.2, f"{loc}m", ha='center', fontsize=9, fontweight='bold', bbox=dict(facecolor='white', alpha=0.5, edgecolor='none'))
                
    plt.close(fig); return fig

def desenhar_grafico_pizza_pdf(series_dados, titulo):
    fig, ax = plt.subplots(figsize=(6, 4))
    counts = series_dados.value_counts()
    if counts.empty: ax.text(0.5,0.5,"N/A"); ax.axis('off'); return fig
    cores = [CORES.get(s, '#95a5a6') for s in counts.index]
    wedges, texts, autotexts = ax.pie(counts.values, colors=cores, autopct=lambda p: f'{p:.1f}%\n({int(round(p*sum(counts.values)/100))})', pctdistance=0.85, wedgeprops=dict(width=0.5, edgecolor='white', linewidth=2), startangle=90)
    for t in autotexts: t.set_fontsize(9); t.set_weight('bold'); t.set_color('black')
    ax.legend(wedges, counts.index, title="Status", loc="center left", bbox_to_anchor=(1, 0, 0.5, 1))
    ax.set_title(titulo, fontweight='bold', color='#2c3e50')
    plt.tight_layout(); plt.close(fig); return fig

# ==============================================================================
# 3. PDF REPORT
# ==============================================================================
class PDFReport(FPDF):
    def __init__(self, cliente, parque, data, turbina_focada=None):
        super().__init__()
        self.cliente = cliente; self.parque = parque; self.data_insp = data; self.turbina_focada = turbina_focada
        self.logo_path = salvar_logo_temporario(); self.sketch_path = "capa_sketch.png"

    def header(self):
        if self.page_no() > 1:
            self.set_font('Arial', 'B', 9); self.set_text_color(100)
            ttl = f"Resumo Executivo - {self.parque}"
            if self.turbina_focada: ttl += f" ({self.turbina_focada})"
            self.set_xy(10, 10); self.cell(0, 10, ttl.encode('latin-1','replace').decode('latin-1'), 0, 1, 'L')
            self.set_draw_color(200); self.line(10, 20, 200, 20); self.ln(10)

    def footer(self):
        self.set_y(-15); self.set_font('Arial','I',8); self.set_text_color(128); self.cell(0,10,f'Pagina {self.page_no()}',0,0,'C')
        if self.page_no() > 1 and self.logo_path and os.path.exists(self.logo_path): self.image(self.logo_path, 185, 282, 20)

    def criar_capa(self):
        self.add_page(); y_image_end = 170
        self.set_fill_color(224, 242, 247); self.rect(0, y_image_end, 210, 297-y_image_end, 'F')
        if os.path.exists(self.sketch_path): self.image(self.sketch_path, x=0, y=0, w=210, h=y_image_end)
        if self.logo_path and os.path.exists(self.logo_path): self.image(self.logo_path, 10, 10, 40)
        self.set_y(185); self.set_font('Arial', 'B', 24); self.set_text_color(44, 62, 80)
        self.cell(0, 15, "RESUMO EXECUTIVO DE SPDA", 0, 1, 'C')
        self.set_font('Arial', '', 14); self.cell(0, 10, "Inspeção de Pás Eólicas com Drone", 0, 1, 'C'); self.ln(20)
        self.set_font('Arial', '', 13); self.set_text_color(60, 60, 60)
        self.cell(0, 9, f"Cliente: {self.cliente}".encode('latin-1','replace').decode('latin-1'), 0, 1, 'C')
        self.cell(0, 9, f"Parque: {self.parque}".encode('latin-1','replace').decode('latin-1'), 0, 1, 'C')
        self.cell(0, 9, f"Data da Inspeção: {self.data_insp}".encode('latin-1','replace').decode('latin-1'), 0, 1, 'C')

    def adicionar_secao(self, ttl):
        self.set_font('Arial','B',12); self.set_fill_color(236,240,241); self.set_text_color(0)
        self.cell(0,8,ttl.encode('latin-1','replace').decode('latin-1'),0,1,'L',True); self.ln(5)

    def adicionar_conclusao_texto(self, texto):
        self.adicionar_secao("Conclusão")
        self.ln(5); self.set_font('Arial', '', 11); self.set_text_color(0)
        self.multi_cell(0, 6, texto.encode('latin-1', 'replace').decode('latin-1')); self.ln(10)

    def criar_tabela_criterios(self):
        self.add_page(); self.adicionar_secao("Critérios de Classificação")
        self.ln(5); self.set_font('Arial', 'B', 9); self.set_fill_color(52, 73, 94); self.set_text_color(255)
        w = [60, 130]; self.cell(w[0], 8, "Status", 1, 0, 'C', True); self.cell(w[1], 8, "Criterio", 1, 1, 'C', True)
        criterios = [("OK", "Resist. <= 250 mOhms", "verde"), ("Alto", "> 250 mOhms", "amarelo"), ("Parcialmente Aberto", "Pelo menos um receptor aberto", "laranja"), ("Aberto", "Todos os receptores abertos", "vermelho")]
        self.set_font('Arial', '', 9)
        for s, e, k in criterios:
            if k=="verde": self.set_fill_color(46,204,113)
            elif k=="amarelo": self.set_fill_color(241,196,15)
            elif k=="laranja": self.set_fill_color(230,126,34)
            else: self.set_fill_color(231,76,60)
            self.set_font('Arial','B',9); self.set_text_color(255) if k!='amarelo' else self.set_text_color(0)
            self.cell(w[0], 8, s.encode('latin-1','replace').decode('latin-1'), 1, 0, 'C', True)
            self.set_font('Arial','',9); self.set_text_color(0); self.cell(w[1], 8, e.encode('latin-1','replace').decode('latin-1'), 1, 1, 'L')
        self.ln(10)

    def criar_tabela_prioridades_tecnica(self, df_prio):
        self.add_page(); self.adicionar_secao("Lista de Prioridades (Por Volume de Falhas)")
        if df_prio.empty: return
        df_prio = df_prio.sort_values(by='Qtd_Anomalias', ascending=False)
        self.set_font('Arial','B',8); self.set_fill_color(240); cols = [40, 30, 40, 80]; headers = ['Turbina', 'Pá', 'Modelo', 'Status']
        for i,h in enumerate(headers): self.cell(cols[i],8,h,1,0,'C',True)
        self.ln(); self.set_font('Arial','',8)
        for _,r in df_prio.iterrows():
            self.set_text_color(0); self.cell(cols[0],8,str(r['Turbine']),1,0,'C'); self.cell(cols[1],8,str(r['Blade Index']),1,0,'C'); self.cell(cols[2],8,str(r['Blade Model']),1,0,'C')
            self.cell(cols[3],8, str(r['Status_Final']), 1, 0, 'C'); self.ln()

    def criar_galeria_fotos(self, df_raw):
        criticos = df_raw[(df_raw['Status_Calc'] == 'Aberto') & (df_raw['Image URL'].notna())].head(3)
        if criticos.empty: return
        self.add_page(); self.adicionar_secao("Galeria de Evidências")
        for _, row in criticos.iterrows():
            self.set_text_color(0) # TEXTO EM PRETO
            y_s = self.get_y(); self.set_fill_color(250); self.rect(10, y_s, 190, 60, 'DF')
            img_p = baixar_foto_da_url(str(row['Image URL']))
            if img_p: self.image(img_p, x=15, y=y_s+5, h=50); os.remove(img_p)
            self.set_xy(95, y_s+8); self.set_font('Arial','B',12); self.cell(0,8,f"Turbina: {row['Turbine']} | Pa {row['Blade Index']}",0,1)
            self.set_font('Arial','',10); self.set_x(95); self.cell(0,6,f"Receptor: {row['Receptor']}",0,1); self.set_x(95); self.cell(0,6,f"Lado: {row['Side']} | Pos: {row['Location']}m",0,1)
            self.set_x(95); self.set_font('Arial','B',10); self.cell(0,6,f"Status: {row['Status_Calc']}",0,1); self.set_y(y_s+65)
        self.set_text_color(0)

def gerar_relatorio_pdf(df, df_piv, cli, parq, data, comp_pa, turb_sel=None):
    pdf = PDFReport(cli, parq, data, turb_sel); pdf.criar_capa()
    if turb_sel:
        pdf.criar_tabela_criterios(); pdf.add_page(); pdf.adicionar_secao(f"Detalhes da Turbina: {turb_sel}")
        df_t = df[df['Turbine'] == turb_sel]
        for idx in ['A', 'B', 'C']:
            df_pa = df_t[df_t['Blade Index'] == idx]
            if not df_pa.empty:
                st_fin = df_piv[(df_piv['Turbine'] == turb_sel) & (df_piv['Blade Index'] == idx)]['Status_Final'].iloc[0]
                modelo = df_pa['Blade Model'].iloc[0]
                pdf.set_font('Arial', 'B', 10); pdf.set_fill_color(220)
                pdf.cell(0, 8, f"Pa {idx} ({modelo}) - {st_fin}".encode('latin-1', 'replace').decode('latin-1'), 0, 1, 'L', True); pdf.ln(2)
                img = salvar_fig_temp(desenhar_pa_individual(df_pa, modelo, idx, comp_pa))
                pdf.image(img, x=10, w=190); os.remove(img); pdf.ln(2)
                pdf.set_font('Arial', 'B', 8); pdf.set_fill_color(240); h = ['Receptor', 'Lado', 'Pos', 'Valor', 'Status']; w = [50, 20, 20, 30, 40]
                pdf.set_x(25); [pdf.cell(w[i], 8, v, 1, 0, 'C', True) for i, v in enumerate(h)]; pdf.ln(); pdf.set_font('Arial', '', 8)
                for _, r in df_pa.iterrows():
                    pdf.set_x(25); pdf.cell(w[0], 8, r['Receptor'][:25], 1, 0, 'C'); pdf.cell(w[1], 8, str(r.get('Side', '')), 1, 0, 'C'); pdf.cell(w[2], 8, str(r['Location']), 1, 0, 'C'); pdf.cell(w[3], 8, str(r['Resistance']), 1, 0, 'C'); pdf.cell(w[4], 8, r['Status_Calc'], 1, 0, 'C'); pdf.ln()
                pdf.ln(5)
    else:
        pdf.add_page(); pdf.adicionar_conclusao_texto(gerar_conclusao_texto(df_piv, df)); pdf.criar_tabela_criterios()
        pdf.add_page(); pdf.adicionar_secao("Visão Estatística da Frota")
        for m in df['Blade Model'].unique():
            img = salvar_fig_temp(desenhar_pa_estatistica(df[df['Blade Model'] == m], m, comp_pa))
            pdf.image(img, x=10, w=190); os.remove(img); pdf.ln(5)
        pdf.add_page(); pdf.adicionar_secao("Indicadores de Frota")
        f1 = desenhar_grafico_pizza_pdf(df['Status_Calc'], "Status Receptores"); i1 = salvar_fig_temp(f1); pdf.image(i1, x=10, w=90)
        f2 = desenhar_grafico_pizza_pdf(df_piv['Status_Final'], "Status Pás"); i2 = salvar_fig_temp(f2); pdf.image(i2, x=110, y=pdf.get_y() - 65, w=90)
        st_u = df_piv.groupby('Turbine')['Status_Turbina'].first()
        f3 = desenhar_grafico_pizza_pdf(st_u, "Status Turbinas"); i3 = salvar_fig_temp(f3); pdf.image(i3, x=60, w=90); os.remove(i1); os.remove(i2); os.remove(i3)

        for modelo in df_piv['Blade Model'].unique():
            pdf.add_page(); pdf.adicionar_secao(f"Dados Completos - Modelo: {modelo}")
            df_m = df_piv[df_piv['Blade Model'] == modelo].sort_values(by=['Turbine', 'Blade Index'])
            cols_r = [c for c in df_m.columns if 'Receptor' in c]; cols_v = df_m[cols_r].dropna(axis=1, how='all').columns.tolist()
            w_id = 20; w_st = 25; w_rec = (145 / len(cols_v)) if cols_v else 15
            pdf.set_font('Arial', 'B', 6); pdf.set_fill_color(230); y_top = pdf.get_y()
            pdf.cell(w_id, 10, "Turb/Pa", 1, 0, 'C', True); pdf.cell(w_st, 10, "Status Pa", 1, 0, 'C', True)
            for c in cols_v:
                x_curr = pdf.get_x(); lbl = c.replace('Receptor ', '').replace('(', '\n(')
                pdf.multi_cell(w_rec, 5, lbl, 1, 'C', fill=True); pdf.set_xy(x_curr + w_rec, y_top)
            pdf.ln(10); pdf.set_font('Arial', '', 6)
            for _, row in df_m.iterrows():
                pdf.set_text_color(0); pdf.cell(w_id, 6, f"{row['Turbine']}-{row['Blade Index']}", 1, 0, 'C'); pdf.cell(w_st, 6, str(row['Status_Final']), 1, 0, 'C')
                for c in cols_v:
                    val = str(row[c]); pdf.set_text_color(200, 0, 0) if val == 'OPEN' else pdf.set_text_color(0)
                    pdf.cell(w_rec, 6, val, 1, 0, 'C')
                pdf.ln()
            pdf.set_text_color(0)
        pdf.criar_tabela_prioridades_tecnica(df_piv[df_piv['Status_Final'] != 'OK'].copy())
        pdf.criar_galeria_fotos(df)
    return pdf.output(dest='S').encode('latin-1')

# ==============================================================================
# 5. INTERFACE
# ==============================================================================
st.sidebar.title("ArthWind App")
with st.sidebar.expander("ℹ️ Regras (250mΩ)", expanded=True):
    st.write("🔴 Aberto: Open / No Signal")
    st.write("🟠 Parcial: 1+ Aberto (não todos)")
    st.write("🟡 Alto: > 250 mΩ")
    st.write("🟢 OK: <= 250 mΩ")

comprimento_pa = st.sidebar.number_input("Comprimento da Pá (m):", min_value=1.0, value=72.0, step=1.0)
f = st.sidebar.file_uploader("Carregar Arquivo", type=['csv', 'xlsx'])

if f:
    df_raw = carregar_dados(f)
    if df_raw is not None:
        df, df_piv = processar_dataframe(df_raw)
        cli = df['Client'].iloc[0] if 'Client' in df.columns else "Cliente"
        pq = df['Windfarm'].iloc[0] if 'Windfarm' in df.columns else "Parque"
        data = str(df['Date'].iloc[0])[:10] if 'Date' in df.columns else datetime.today().strftime('%Y-%m-%d')
        md = st.sidebar.selectbox("Modelo", ['Todos']+list(df['Blade Model'].unique()))
        df_show = df if md=='Todos' else df[df['Blade Model']==md]
        df_piv_show = df_piv if md=='Todos' else df_piv[df_piv['Blade Model']==md]
        st.title("⚡ RESUMO EXECUTIVO SPDA"); c1,c2,c3,c4,c5=st.columns(5)
        c1.metric("Total Pás",len(df_piv_show)); c2.metric("OK",len(df_piv_show[df_piv_show['Status_Final']=='OK']))
        c3.metric("Alto",len(df_piv_show[df_piv_show['Status_Final']=='Alto']),delta_color='off')
        c4.metric("Parcial",len(df_piv_show[df_piv_show['Status_Final']=='Parcialmente Aberto']),delta_color='inverse')
        c5.metric("Aberto",len(df_piv_show[df_piv_show['Status_Final']=='Aberto']),delta_color='inverse')
        st.info(gerar_conclusao_texto(df_piv_show, df_show))
        t1,t2,t3 = st.tabs(["🔍 Turbina","📊 Frota","📋 Dados"])
        with t1:
            cs, cb = st.columns([3,1]); ts = cs.selectbox("Turbina:", sorted([str(t) for t in df_show['Turbine'].unique()]))
            cb.write(""); cb.download_button("📄 PDF Turbina", data=gerar_relatorio_pdf(df_show,df_piv_show,cli,pq,data,comprimento_pa,ts), file_name=f"Rel_{ts}.pdf", mime="application/pdf")
            df_t = df_show[df_show['Turbine']==ts]; ca,cb,cc = st.columns(3)
            for i,c in zip(['A','B','C'],[ca,cb,cc]):
                with c:
                    r=df_piv_show[(df_piv_show['Turbine']==ts)&(df_piv_show['Blade Index']==i)]
                    if not r.empty:
                        s=r['Status_Final'].iloc[0]; m=r['Blade Model'].iloc[0]
                        st.markdown(f"<div style='border:1px solid #ddd;border-radius:8px;text-align:center;padding:5px'><b>Pá {i}</b><br><span style='color:{CORES[s]}'>{s}</span></div>",unsafe_allow_html=True)
                        st.pyplot(desenhar_pa_individual(df_t[df_t['Blade Index']==i],m,i,comprimento_pa))
        with t2:
            ch,cb=st.columns([3,1]); ch.subheader("Visão Geral"); cb.write(""); cb.download_button("📄 PDF Frota",data=gerar_relatorio_pdf(df_show,df_piv_show,cli,pq,data,comprimento_pa),file_name="Rel_Frota.pdf",mime="application/pdf")
            for m in df_show['Blade Model'].unique(): st.pyplot(desenhar_pa_estatistica(df_show[df_show['Blade Model']==m],m,comprimento_pa))
            c_rec, c_pas, c_turb = st.columns(3); scale = alt.Scale(domain=['OK','Alto','Parcialmente Aberto','Aberto'], range=[CORES['OK'], CORES['Alto'], CORES['Parcialmente Aberto'], CORES['Aberto']])
            for label, series, col in [("Receptores", df_show['Status_Calc'], c_rec), ("Pás", df_piv_show['Status_Final'], c_pas), ("Turbinas", df_piv_show.groupby('Turbine')['Status_Turbina'].first(), c_turb)]:
                counts = series.value_counts().reset_index(); counts.columns=['Status','Qtd']
                col.markdown(f"**{label}**"); col.altair_chart(alt.Chart(counts).mark_arc(innerRadius=50).encode(theta='Qtd', color=alt.Color('Status', scale=scale)), use_container_width=True)
        with t3:
            def highlight(val):
                s = str(val)
                if s in ['OPEN', 'Aberto']: return f'color: {CORES["Aberto"]}; font-weight: bold'
                if s == 'Parcialmente Aberto': return f'color: {CORES["Parcialmente Aberto"]}; font-weight: bold'
                if s == 'Alto': return f'color: {CORES["Alto"]}; font-weight: bold'
                if s == 'OK': return f'color: {CORES["OK"]}; font-weight: bold'
                return 'color: black'
            for modelo in df_piv_show['Blade Model'].unique():
                st.subheader(f"Modelo: {modelo}"); df_mod = df_piv_show[df_piv_show['Blade Model'] == modelo]
                cols_r = [c for c in df_mod.columns if 'Receptor' in c]; cols_v = df_mod[cols_r].dropna(axis=1, how='all').columns.tolist()
                st.dataframe(df_mod[['Turbine','Blade Index', 'Status_Turbina', 'Status_Final'] + cols_v].style.map(highlight), use_container_width=True)

