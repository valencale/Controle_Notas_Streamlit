"""
analise_energia.py — Analise de consumo energetico do Palacio do Jaburu.
Gera 4 graficos (PNG) para o portfolio.
"""
import pandas as pd
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import plotly.express as px
import plotly.graph_objects as go
import os

DIR = os.path.dirname(os.path.abspath(__file__))
CSV = os.path.join(DIR, "consumo_energia_jaburu.csv")

# Ordem dos meses
MESES_ORDEM = ['Janeiro','Fevereiro','Marco','Abril','Maio','Junho',
               'Julho','Agosto','Setembro','Outubro','Novembro','Dezembro']

def carregar_dados():
    df = pd.read_csv(CSV, sep=';', encoding='latin-1')
    df.columns = ['ANO','MES','CONSUMO_KWH','VALOR_FATURA']
    # Limpar MES
    df['MES'] = df['MES'].str.strip()
    # Converter valor: "R$ 17.910,77" -> float
    df['VALOR'] = (df['VALOR_FATURA']
                   .str.replace('R$','',regex=False)
                   .str.replace('.','',regex=False)
                   .str.replace(',','.',regex=False)
                   .str.strip()
                   .astype(float))
    # Normalizar nomes de mes (remover acentos simples)
    replace_map = {'Marco':'Marco','Fevereiro':'Fevereiro'}
    df['MES_NORM'] = df['MES'].replace({'Março':'Marco','março':'Marco'})
    return df

def grafico_linha_anual(df):
    """Grafico de linha: tendencias anuais de consumo e valor."""
    anual = df.groupby('ANO').agg(
        consumo_total=('CONSUMO_KWH','sum'),
        valor_medio=('VALOR','mean')
    ).reset_index()
    
    fig, ax1 = plt.subplots(figsize=(10, 5))
    ax1.set_facecolor('#f8f9fa')
    fig.patch.set_facecolor('white')
    
    color1 = '#2563eb'
    ax1.plot(anual['ANO'], anual['consumo_total'], 'o-', color=color1, 
             linewidth=2, markersize=8, label='Consumo Total (kWh)')
    ax1.set_xlabel('Ano', fontsize=12)
    ax1.set_ylabel('Consumo Total (kWh)', color=color1, fontsize=12)
    ax1.tick_params(axis='y', labelcolor=color1)
    
    ax2 = ax1.twinx()
    color2 = '#dc2626'
    ax2.plot(anual['ANO'], anual['valor_medio'], 's--', color=color2,
             linewidth=2, markersize=8, label='Valor Medio Fatura (R$)')
    ax2.set_ylabel('Valor Medio Fatura (R$)', color=color2, fontsize=12)
    ax2.tick_params(axis='y', labelcolor=color2)
    
    plt.title('Tendencias Anuais - Consumo e Valor da Fatura', fontsize=14, fontweight='bold')
    lines1, labels1 = ax1.get_legend_handles_labels()
    lines2, labels2 = ax2.get_legend_handles_labels()
    ax1.legend(lines1+lines2, labels1+labels2, loc='upper right')
    plt.tight_layout()
    
    path = os.path.join(DIR, 'grafico_linha_anual.png')
    plt.savefig(path, dpi=150, bbox_inches='tight')
    plt.close()
    print(f"[OK] {path}")
    return anual

def grafico_barras_mensal(df):
    """Grafico de barras: consumo medio por mes."""
    # Mapear meses para ordem numerica
    mes_num = {'Janeiro':1,'Fevereiro':2,'Marco':3,'Abril':4,'Maio':5,
               'Junho':6,'Julho':7,'Agosto':8,'Setembro':9,'Outubro':10,
               'Novembro':11,'Dezembro':12}
    df2 = df.copy()
    df2['MES_CLEAN'] = df2['MES'].replace({'Março':'Marco','março':'Marco'})
    df2['MES_NUM'] = df2['MES_CLEAN'].map(mes_num)
    df2 = df2.dropna(subset=['MES_NUM'])
    
    mensal = df2.groupby(['MES_CLEAN','MES_NUM'])['CONSUMO_KWH'].mean().reset_index()
    mensal = mensal.sort_values('MES_NUM')
    
    fig, ax = plt.subplots(figsize=(10, 5))
    colors = ['#3b82f6' if v < mensal['CONSUMO_KWH'].mean() else '#ef4444' 
              for v in mensal['CONSUMO_KWH']]
    ax.bar(mensal['MES_CLEAN'], mensal['CONSUMO_KWH'], color=colors, edgecolor='white')
    ax.axhline(y=mensal['CONSUMO_KWH'].mean(), color='#6b7280', linestyle='--', 
               label=f'Media: {mensal["CONSUMO_KWH"].mean():,.0f} kWh')
    ax.set_xlabel('Mes', fontsize=12)
    ax.set_ylabel('Consumo Medio (kWh)', fontsize=12)
    ax.set_title('Consumo Medio de Energia por Mes', fontsize=14, fontweight='bold')
    ax.legend()
    plt.xticks(rotation=45, ha='right')
    plt.tight_layout()
    
    path = os.path.join(DIR, 'grafico_barras_mensal.png')
    plt.savefig(path, dpi=150, bbox_inches='tight')
    plt.close()
    print(f"[OK] {path}")

def grafico_correlacao(df):
    """Scatter plot: correlacao consumo vs valor."""
    fig, ax = plt.subplots(figsize=(8, 6))
    scatter = ax.scatter(df['CONSUMO_KWH'], df['VALOR'], c=df['ANO'], 
                        cmap='viridis', alpha=0.7, s=60, edgecolors='white')
    plt.colorbar(scatter, label='Ano')
    
    # Linha de tendencia
    z = pd.np.polyfit(df['CONSUMO_KWH'], df['VALOR'], 1) if hasattr(pd, 'np') else __import__('numpy').polyfit(df['CONSUMO_KWH'], df['VALOR'], 1)
    p = __import__('numpy').poly1d(z)
    x_line = sorted(df['CONSUMO_KWH'])
    ax.plot(x_line, p(x_line), '--', color='#ef4444', linewidth=2, label='Tendencia')
    
    corr = df['CONSUMO_KWH'].corr(df['VALOR'])
    ax.set_xlabel('Consumo (kWh)', fontsize=12)
    ax.set_ylabel('Valor da Fatura (R$)', fontsize=12)
    ax.set_title(f'Correlacao: Consumo x Valor (r = {corr:.2f})', fontsize=14, fontweight='bold')
    ax.legend()
    plt.tight_layout()
    
    path = os.path.join(DIR, 'grafico_correlacao.png')
    plt.savefig(path, dpi=150, bbox_inches='tight')
    plt.close()
    print(f"[OK] {path}")
    return corr

def grafico_interativo_plotly(df):
    """Grafico interativo com Plotly (salvo como PNG)."""
    fig = px.line(df, x=df.index, y='CONSUMO_KWH', color=df['ANO'].astype(str),
                  labels={'CONSUMO_KWH':'Consumo (kWh)','index':'Registro','color':'Ano'},
                  title='Consumo Energetico Mensal - Visao Interativa')
    fig.update_layout(template='plotly_white', font=dict(size=12))
    
    path = os.path.join(DIR, 'grafico_interativo.png')
    fig.write_image(path, width=1000, height=500, scale=2)
    print(f"[OK] {path}")


if __name__ == "__main__":
    print("=" * 60)
    print("  ANALISE DE ENERGIA — Palacio do Jaburu")
    print("=" * 60)
    
    df = carregar_dados()
    print(f"\nDataset: {df.shape[0]} registros, {df.shape[1]} colunas")
    print(f"Periodo: {df['ANO'].min()} a {df['ANO'].max()}\n")
    
    anual = grafico_linha_anual(df)
    grafico_barras_mensal(df)
    corr = grafico_correlacao(df)
    grafico_interativo_plotly(df)
    
    print(f"\nCorrelacao consumo x valor: {corr:.2f}")
    print("\n[OK] Todos os graficos gerados com sucesso!")
