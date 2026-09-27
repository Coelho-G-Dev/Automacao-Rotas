import sys, os
sys.stdout.reconfigure(encoding='utf-8')
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))
import pandas as pd
from automacao import (
    localizar_aba_e_cabecalho, carregar_dados_frios, obter_base_enderecos, normalizar_nome_rota,
    otimizar_veiculos_rota_pura, selecionar_menor_veiculo
)

aba, row = localizar_aba_e_cabecalho('Acomp_EvoluSep.xlsx')
df = carregar_dados_frios('Acomp_EvoluSep.xlsx', aba, row)
base = obter_base_enderecos(df, 'Cadastro_Lojas_Enderecos.xlsx')

col_d = "DIAS_CARREGAMENTO" if "DIAS_CARREGAMENTO" in df.columns else "Dias Carregamento"

for idx in df.index:
    f = df.at[idx, 'FILIAL_PADRAO']
    if f in base:
        r_cad = base[f].get('Rota Padrão')
        reg_cad = base[f].get('Região')
        if r_cad and str(r_cad).strip() not in ('', 'nan', 'None'):
            df.at[idx, 'ROTA_PADRAO'] = normalizar_nome_rota(r_cad)
        if reg_cad and str(reg_cad).strip() not in ('', 'nan', 'None'):
            df.at[idx, 'REGIAO_PADRAO'] = str(reg_cad).strip().upper()

# Testar Chapadinha
sub_chap = df[df['ROTA_PADRAO'].str.contains('Chapadinha', case=False, na=False)]
lojas_chap = []
for f, grp in sub_chap.groupby('FILIAL_PADRAO'):
    p_tot = grp['PESOBRUTO'].sum()
    p_h = grp[grp[col_d].astype(str).str.contains("01-DE HOJE|HOJE", case=False, na=False)]['PESOBRUTO'].sum()
    lojas_chap.append({
        'regiao': 'INTERIOR', 'rota': 'Rota 2 - Chapadinha / Codó', 'filial': f, 'filial_orig': f,
        'peso': p_tot, 'peso_hoje': p_h, 'tem_hoje': (p_h > 0)
    })

veics_c, sobras_c = otimizar_veiculos_rota_pura('Rota 2 - Chapadinha / Codó', 'INTERIOR', lojas_chap, base)
print("=== CHAPADINHA ===")
print("Veículos:", veics_c)
print("Sobras:", [(s['filial'], s['peso']) for s in sobras_c])
for s in lojas_chap:
    print(f"Loja {s['filial']}: {s['peso']:.2f} kg (Hoje: {s['peso_hoje']:.2f} kg)")

# Testar se qualquer combinacao de lojas em Chapadinha atinge algum veiculo
import itertools
for k in [1, 2, 3]:
    for c in itertools.combinations(lojas_chap, k):
        p_c = sum(x['peso'] for x in c)
        f_list = [x['filial'] for x in c]
        v, exc, falta = selecionar_menor_veiculo(p_c, f_list, 'INTERIOR', base)
        print(f"Combo {[x['filial'] for x in c]} -> Peso: {p_c:.2f} kg -> Veículo: {v}")
