import sys, os
sys.stdout.reconfigure(encoding='utf-8')
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))
import pandas as pd
from automacao import (
    localizar_aba_e_cabecalho, carregar_dados_frios, obter_base_enderecos,
    otimizar_veiculos_rota_pura, PARAMETROS_FROTA, obter_veiculos_compativeis, selecionar_menor_veiculo
)

aba, row = localizar_aba_e_cabecalho('Acomp_EvoluSep.xlsx')
df_frios = carregar_dados_frios('Acomp_EvoluSep.xlsx', aba, row)
base = obter_base_enderecos(df_frios, 'Cadastro_Lojas_Enderecos.xlsx')

col_d = "DIAS_CARREGAMENTO" if "DIAS_CARREGAMENTO" in df_frios.columns else "Dias Carregamento"

# Verificar Lojas 7, 19, 411, 16, 408, 30, 17, 217, 96, 20, 29
lojas_interesse = ['7', '19', '411', '16', '408', '30', '17', '217', '96', '20', '29', '61', '2061', '34']

print("=== VERIFICAÇÃO CADASTRO E PESOS DAS LOJAS DE INTERESSE ===")
for f in lojas_interesse:
    sub = df_frios[df_frios['FILIAL_PADRAO'] == f]
    cad = base.get(f, {})
    rota_cad = cad.get('Rota Padrão', 'NÃO CADASTRADO')
    reg_cad = cad.get('Região', 'NÃO CADASTRADO')
    restr = cad.get('Restrição Veículo', 'Sem restrição')
    
    p_tot = sub['PESOBRUTO'].sum()
    p_hoje = sub[sub[col_d].astype(str).str.contains("01-DE HOJE|HOJE", case=False, na=False)]['PESOBRUTO'].sum()
    p_prazo = sub[sub[col_d].astype(str).str.contains("02-NO PRAZO|PRAZO", case=False, na=False)]['PESOBRUTO'].sum()
    p_urg = sub[sub[col_d].astype(str).str.contains("03-URGENTE|URGENTE", case=False, na=False)]['PESOBRUTO'].sum()
    
    print(f"Loja {f:>4} | Rota Cad: {rota_cad:<8} | Reg: {reg_cad:<7} | Restr: {restr:<15} | Tot: {p_tot:>9.2f} kg | Hoje: {p_hoje:>9.2f} | Prazo: {p_prazo:>9.2f} | Urg: {p_urg:>9.2f}")
