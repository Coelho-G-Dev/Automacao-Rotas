import sys, os
sys.stdout.reconfigure(encoding='utf-8')
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))
import pandas as pd
from automacao import (
    localizar_aba_e_cabecalho, carregar_dados_frios, obter_base_enderecos, normalizar_nome_rota
)

aba, row = localizar_aba_e_cabecalho('Acomp_EvoluSep.xlsx')
df = carregar_dados_frios('Acomp_EvoluSep.xlsx', aba, row)
base = obter_base_enderecos(df, 'Cadastro_Lojas_Enderecos.xlsx')

col_d = "DIAS_CARREGAMENTO" if "DIAS_CARREGAMENTO" in df.columns else "Dias Carregamento"

# Preencher rotas e regioes da base
for idx in df.index:
    f = df.at[idx, 'FILIAL_PADRAO']
    if f in base:
        r_cad = base[f].get('Rota Padrão')
        reg_cad = base[f].get('Região')
        if r_cad and str(r_cad).strip() not in ('', 'nan', 'None'):
            df.at[idx, 'ROTA_PADRAO'] = normalizar_nome_rota(r_cad)
        if reg_cad and str(reg_cad).strip() not in ('', 'nan', 'None'):
            df.at[idx, 'REGIAO_PADRAO'] = str(reg_cad).strip().upper()

print("\n" + "=" * 90)
print(f"PANORAMA GERAL DO DIA 26/09 (TOTAL: {df['PESOBRUTO'].sum():,.2f} kg em {df['FILIAL_PADRAO'].nunique()} LOJAS)")
print("=" * 90)

for reg in ['CAPITAL', 'INTERIOR']:
    sub_reg = df[df['REGIAO_PADRAO'] == reg]
    print(f"\n>>> REGIÃO: {reg} (Total: {sub_reg['PESOBRUTO'].sum():,.2f} kg, {sub_reg['FILIAL_PADRAO'].nunique()} lojas)")
    
    rotas_unicas = sorted(sub_reg['ROTA_PADRAO'].dropna().unique(), key=lambda r: int(''.join(filter(str.isdigit, str(r)))) if any(c.isdigit() for c in str(r)) else 999)
    for r in rotas_unicas:
        sub_r = sub_reg[sub_reg['ROTA_PADRAO'] == r]
        p_tot = sub_r['PESOBRUTO'].sum()
        p_hoje = sub_r[sub_r[col_d].astype(str).str.contains("01-DE HOJE|HOJE", case=False, na=False)]['PESOBRUTO'].sum()
        p_prazo = sub_r[sub_r[col_d].astype(str).str.contains("02-NO PRAZO|PRAZO", case=False, na=False)]['PESOBRUTO'].sum()
        p_urg = sub_r[sub_r[col_d].astype(str).str.contains("03-URGENTE|URGENTE", case=False, na=False)]['PESOBRUTO'].sum()
        
        lojas_r = sub_r.groupby('FILIAL_PADRAO').agg({
            'PESOBRUTO': 'sum',
            col_d: lambda s: ', '.join(set(s.dropna().astype(str)))
        }).reset_index()
        
        print(f"\n   [ROTA] {r} | Total: {p_tot:>9.2f} kg (Hoje: {p_hoje:>9.2f} kg | Prazo: {p_prazo:>9.2f} kg | Urgente: {p_urg:>9.2f} kg)")
        for _, l in lojas_r.iterrows():
            print(f"      • Loja {l['FILIAL_PADRAO']:>4}: {l['PESOBRUTO']:>8.2f} kg [{l[col_d]}]")
