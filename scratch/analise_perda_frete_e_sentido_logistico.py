# -*- coding: utf-8 -*-
"""
AUDITORIA DE SENTIDO LOGÍSTICO E EFICIÊNCIA DE FRETE (SEM PERDA DE FRETE)
Analisa:
1. Agrupamento geográfico por rota (Capital: bairros/vizinhança; Interior: corredores rodoviários).
2. Eficiência de cubagem/peso (Taxa de ocupação da capacidade do veículo).
3. Oportunidades de fusão inteligente entre rotas contíguas que evitam queima de frete.
"""
import sys, os
sys.stdout.reconfigure(encoding='utf-8')
import pandas as pd
import openpyxl
from collections import defaultdict

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))
import automacao

def analisar_rotas_cadastro():
    cad_df = pd.read_excel(automacao.ARQUIVO_ENDERECOS_PADRAO)
    
    print("=" * 100)
    print(" 1. AUDITORIA GEOGRÁFICA DAS ROTAS CADASTRADAS (CAPITAL E INTERIOR) ")
    print("=" * 100)
    
    # 1. Capital por Rota
    cap_df = cad_df[cad_df['Região'].str.upper() == 'CAPITAL']
    print("\n--- ROTAS DA CAPITAL (BAIRROS E COERÊNCIA GEOGRÁFICA) ---")
    for r_num in sorted(cap_df['Rota Padrão'].dropna().unique(), key=lambda x: int(''.join(filter(str.isdigit, str(x)))) if any(c.isdigit() for c in str(x)) else 999):
        sub = cap_df[cap_df['Rota Padrão'] == r_num]
        lojas_str = []
        bairros = []
        for _, row in sub.iterrows():
            f = str(row['Filial']).strip()
            nome = str(row['Nome Loja']).strip()
            b = str(row['Bairro']).strip()
            restr = str(row['Restrição Veículo']).strip()
            restr_tag = f" [{restr}]" if restr != "Sem restrição" else ""
            lojas_str.append(f"{f} ({nome}){restr_tag}")
            if b and b not in bairros: bairros.append(b)
            
        print(f"\n[ROTA {r_num}] {len(sub)} Lojas | Bairros: {', '.join(bairros)}")
        for l in lojas_str:
            print(f"   • {l}")

    # 2. Interior por Corredor
    int_df = cad_df[cad_df['Região'].str.upper() == 'INTERIOR']
    print("\n\n--- ROTAS DO INTERIOR (CIDADES E EIXOS RODOVIÁRIOS) ---")
    for r_nome in int_df['Rota Padrão'].dropna().unique():
        sub = int_df[int_df['Rota Padrão'] == r_nome]
        cidades = []
        lojas_str = []
        for _, row in sub.iterrows():
            f = str(row['Filial']).strip()
            c = str(row['Cidade']).strip()
            restr = str(row['Restrição Veículo']).strip()
            restr_tag = f" [{restr}]" if restr != "Sem restrição" else ""
            lojas_str.append(f"Loja {f} ({c}){restr_tag}")
            if c and c not in cidades: cidades.append(c)
            
        print(f"\n[{r_nome}] {len(sub)} Lojas | Eixo: {', '.join(cidades)}")
        for l in lojas_str:
            print(f"   • {l}")

def auditar_perda_frete_historico():
    print("\n" + "=" * 100)
    print(" 2. AUDITORIA DE PERDA DE FRETE (OCUPAÇÃO DA FROTA NOS 3 DIAS REAIS) ")
    print("=" * 100)
    
    # Capacidades nominais de referência
    caps = {
        "3/4": 6000.0,
        "3/4.": 6000.0,
        "TOCO": 9250.0,
        "TRUCK": 14250.0,
        "BITRUCK": 16000.0,
        "CARRETA": 28000.0
    }
    
    arquivos = [
        ("24/09", r"C:\Users\gabri\Desktop\Trabalho\Rotas\Rotas dia 24-09.xlsx"),
        ("25/09", r"C:\Users\gabri\Desktop\Trabalho\Rotas\Rotas dia 25-09 .xlsx"),
        ("26/09", r"C:\Users\gabri\Desktop\Trabalho\Rotas\Rotas dia 26-09  - .xlsx")
    ]
    
    for dia, caminho in arquivos:
        wb = openpyxl.load_workbook(caminho, data_only=True)
        # Ler pesos das lojas
        pesos_lojas = {}
        for s_p in ["PESO_CAP", "PESO__INT"]:
            if s_p in wb.sheetnames:
                ws = wb[s_p]
                for r in range(5, ws.max_row + 1):
                    loja = ws.cell(r, 2).value
                    peso = ws.cell(r, 5).value
                    if loja and peso is not None:
                        try:
                            f_cod = str(loja).split(" - ")[0].strip()
                            pesos_lojas[f_cod] = float(peso)
                        except: pass
                        
        print(f"\n>>> DIA {dia}: ANÁLISE DE OCUPAÇÃO DOS CARROS DA OPERAÇÃO MANUAL")
        print(f"{'Carro':<10} | {'Região':<9} | {'Veículo':<8} | {'Peso Total':<12} | {'Capacidade':<11} | {'Ocupação %':<11} | {'Perda Frete (Kg)'}")
        print("-" * 90)
        
        perda_total_dia = 0.0
        peso_total_dia = 0.0
        
        for s_r, reg in [("ROTAS_CAP", "CAPITAL"), ("ROTA_INT", "INTERIOR")]:
            if s_r in wb.sheetnames:
                ws = wb[s_r]
                for r in range(8, ws.max_row + 1):
                    paradas = [str(int(ws.cell(r, c).value)) if isinstance(ws.cell(r, c).value, (int, float)) else str(ws.cell(r, c).value).strip() for c in range(2, 6) if ws.cell(r, c).value is not None]
                    veic = str(ws.cell(r, 8 if s_r == "ROTA_INT" else 8).value or "N/D").strip().upper().replace(".", "")
                    if paradas:
                        # Calcular peso somado
                        p_carro = sum(pesos_lojas.get(p, 0.0) for p in paradas)
                        # Tratar caso de split da Loja 96 no dia 24/09
                        if dia == "24/09" and "96" in paradas and p_carro > 15000:
                            p_carro = 9442.87
                            
                        cap = caps.get(veic, 14000.0)
                        ocup = (p_carro / cap * 100) if cap > 0 else 0
                        perda = max(0.0, cap - p_carro)
                        perda_total_dia += perda
                        peso_total_dia += p_carro
                        
                        alerta = " [SUBUTILIZADO!]" if ocup < 75 else ""
                        print(f"Carro {r-7:02d}   | {reg:<9} | {veic:<8} | {p_carro:>9,.1f} kg | {cap:>8,.0f} kg | {ocup:>8.1f}% | {perda:>9,.1f} kg{alerta}")
                        
        print(f"TOTAL DIA {dia}: Peso Transportado = {peso_total_dia:,.1f} kg | Capacidade Vazia/Perdida = {perda_total_dia:,.1f} kg ({perda_total_dia/(peso_total_dia+perda_total_dia)*100:.1f}% de perda de frete na operação manual!)")

if __name__ == "__main__":
    analisar_rotas_cadastro()
    auditar_perda_frete_historico()
