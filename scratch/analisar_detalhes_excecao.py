import sys
sys.path.insert(0, r"C:\Users\gabri\Desktop\Separacao")
import pandas as pd
from automacao import (
    carregar_dados_frios, 
    localizar_aba_e_cabecalho, 
    obter_base_enderecos,
    ARQUIVO_ENTRADA_PADRAO, 
    ARQUIVO_ENDERECOS_PADRAO
)

aba, lin = localizar_aba_e_cabecalho(ARQUIVO_ENTRADA_PADRAO)
df = carregar_dados_frios(ARQUIVO_ENTRADA_PADRAO, aba, lin)
cad = obter_base_enderecos(df, ARQUIVO_ENDERECOS_PADRAO)

print("\n=== RESUMO DETALHADO POR ROTA NA CAPITAL ===")
df_cap = df[df["REGIAO_PADRAO"] == "CAPITAL"]
col_dias = "DIAS_CARREGAMENTO" if "DIAS_CARREGAMENTO" in df.columns else "Dias Carregamento"

for rota, grp in df_cap.groupby("ROTA_PADRAO"):
    p_tot = grp["PESOBRUTO"].sum()
    lojas = grp["FILIAL_PADRAO"].unique()
    print(f"\n--- {rota} (Total: {p_tot:,.2f} kg) ---")
    for l in lojas:
        l_df = grp[grp["FILIAL_PADRAO"] == l]
        p_l = l_df["PESOBRUTO"].sum()
        p_hoje = l_df[l_df[col_dias].astype(str).str.contains("01-DE HOJE|HOJE", case=False, na=False)]["PESOBRUTO"].sum()
        p_prazo = l_df[l_df[col_dias].astype(str).str.contains("02-NO PRAZO|PRAZO", case=False, na=False)]["PESOBRUTO"].sum()
        p_urg = l_df[l_df[col_dias].astype(str).str.contains("03-URGENTE|URGENTE", case=False, na=False)]["PESOBRUTO"].sum()
        restr = cad.get(l, {}).get("Restrição Veículo", "Sem restrição")
        print(f"  Loja {l:5s}: Total={p_l:9.2f} kg | Hoje={p_hoje:9.2f} | Prazo={p_prazo:8.2f} | Urg={p_urg:8.2f} | Restr={restr}")
