import sys, os
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))
import pandas as pd
from automacao import obter_base_enderecos, carregar_dados_frios, localizar_aba_e_cabecalho, normalizar_nome_rota

aba, row = localizar_aba_e_cabecalho('Acomp_EvoluSep.xlsx')
df = carregar_dados_frios('Acomp_EvoluSep.xlsx', aba, row)
base = obter_base_enderecos(df, 'Cadastro_Lojas_Enderecos.xlsx')

for idx in df.index:
    f = df.at[idx, 'FILIAL_PADRAO']
    if f in base:
        r_cad = base[f].get('Rota Padrão')
        reg_cad = base[f].get('Região')
        if r_cad and str(r_cad).strip() not in ('', 'nan', 'None'):
            df.at[idx, 'ROTA_PADRAO'] = normalizar_nome_rota(r_cad)
        if reg_cad and str(reg_cad).strip() not in ('', 'nan', 'None'):
            df.at[idx, 'REGIAO_PADRAO'] = str(reg_cad).strip().upper()

print("=== DISTRIBUIÇÃO DAS LOJAS POR ROTA OFICIAL NA CAPITAL ===")
for (reg, rota), grp in df.groupby(['REGIAO_PADRAO', 'ROTA_PADRAO']):
    if reg == 'CAPITAL':
        p_tot = grp['PESOBRUTO'].sum()
        p_hoje = grp[grp['DIAS_CARREGAMENTO'].str.contains('HOJE', na=False)]['PESOBRUTO'].sum()
        p_sobra = p_tot - p_hoje
        print(f"\n{rota} ({reg}) | Total: {p_tot:>9.2f} kg (Hoje: {p_hoje:>9.2f} kg, Sobra: {p_sobra:>9.2f} kg)")
        lojas = grp.groupby('FILIAL_PADRAO').agg({'PESOBRUTO': 'sum', 'DIAS_CARREGAMENTO': 'first'}).reset_index()
        for _, l in lojas.iterrows():
            print(f"   • Loja {l['FILIAL_PADRAO']:>4}: {l['PESOBRUTO']:>9.2f} kg ({l['DIAS_CARREGAMENTO']})")
