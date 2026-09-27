# -*- coding: utf-8 -*-
import sys
sys.path.append(r"c:\Users\gabri\Desktop\Separacao")
import automacao
import pandas as pd
import itertools

aba, cab = automacao.localizar_aba_e_cabecalho(automacao.ARQUIVO_ENTRADA_PADRAO)
df_frios = automacao.carregar_dados_frios(automacao.ARQUIVO_ENTRADA_PADRAO, aba, cab)
cad = automacao.obter_base_enderecos(df_frios, automacao.ARQUIVO_ENDERECOS_PADRAO)

for idx in df_frios.index:
    f = df_frios.at[idx, 'FILIAL_PADRAO']
    if f in cad:
        r_cad = cad[f].get('Rota Padrão')
        reg_cad = cad[f].get('Região')
        if r_cad and str(r_cad).strip() not in ('', 'nan', 'None'):
            df_frios.at[idx, 'ROTA_PADRAO'] = automacao.normalizar_nome_rota(r_cad)
        if reg_cad and str(reg_cad).strip() not in ('', 'nan', 'None'):
            df_frios.at[idx, 'REGIAO_PADRAO'] = str(reg_cad).strip().upper()

col_dias = automacao.encontrar_coluna(df_frios.columns, ['Dias Carregamento', 'Dias Pedidos', 'CARREGAMENTO'])

# Agregação por loja analisando os status
dados_lojas = []
for (reg, rota, f), grp in df_frios.groupby(['REGIAO_PADRAO', 'ROTA_PADRAO', 'FILIAL_PADRAO']):
    p_total = float(grp['PESOBRUTO'].sum())
    p_hoje = float(grp[grp[col_dias] == '01-DE HOJE']['PESOBRUTO'].sum())
    p_prazo = float(grp[grp[col_dias] == '02-NO PRAZO']['PESOBRUTO'].sum())
    p_urgente = float(grp[grp[col_dias] == '03-URGENTE']['PESOBRUTO'].sum())
    
    tem_hoje = p_hoje > 0
    if tem_hoje:
        status_demanda = "01-DE HOJE"
        tipo_demanda = "Pedido do Dia (Obrigatório)"
    elif p_urgente > 0:
        status_demanda = "03-URGENTE"
        tipo_demanda = "Complemento (Urgente)"
    else:
        status_demanda = "02-NO PRAZO"
        tipo_demanda = "Complemento (No Prazo)"
        
    dados_lojas.append({
        "regiao": reg,
        "rota": rota,
        "filial": f,
        "filial_orig": f,
        "peso_total": p_total,
        "peso_hoje": p_hoje,
        "peso_prazo": p_prazo,
        "peso_urgente": p_urgente,
        "tem_hoje": tem_hoje,
        "status_demanda": status_demanda,
        "tipo_demanda": tipo_demanda
    })

df_perfil_lojas = pd.DataFrame(dados_lojas)
print("=== PERFIL DE DEMANDA DAS LOJAS ===")
print(df_perfil_lojas[['rota', 'filial', 'peso_total', 'tem_hoje', 'status_demanda', 'tipo_demanda']].head(30))
