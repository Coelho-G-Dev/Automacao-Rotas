# -*- coding: utf-8 -*-
"""
ATUALIZAR CADASTRO DE LOJAS COM:
1. Loja 251 (Mix Parnaíba) -> Rota 3 - BR-222 / Baixo Parnaíba (Ciclo TER-SEX).
2. Adicionar coluna 'Dias Geração' com os ciclos oficiais da aba CIclo.
3. Loja 211 (Piripiri) -> Restrição 'Somente Carreta' (Transferência Ceará).
"""
import sys, os
sys.stdout.reconfigure(encoding='utf-8')
import pandas as pd
import openpyxl

caminho_cad = r"C:\Users\gabri\Desktop\Separacao\Cadastro_Lojas_Enderecos.xlsx"
caminho_evolu = r"C:\Users\gabri\Desktop\Separacao\Acomp_EvoluSep.xlsx"

# 1. Carregar mapeamento da aba CIclo
df_ciclo = pd.read_excel(caminho_evolu, sheet_name='CIclo', header=4)
col_ciclo_dias = [c for c in df_ciclo.columns if 'DIAS_GERA' in c.upper()][0]
df_ciclo['Filial_str'] = df_ciclo['CLIENTE'].dropna().astype(int).astype(str)
map_ciclo = df_ciclo.set_index('Filial_str')[col_ciclo_dias].to_dict()

# Adicionar Barreirinhas 359 como QUA-SAB
map_ciclo['359'] = 'QUA-SAB'

# 2. Carregar e atualizar cadastro
df_cad = pd.read_excel(caminho_cad)
df_cad['Filial_str'] = df_cad['Filial'].astype(str)

# Atualizar Loja 251
idx_251 = df_cad[df_cad['Filial_str'] == '251'].index
if len(idx_251) > 0:
    df_cad.loc[idx_251, 'Rota Padrão'] = 'Rota 3 - BR-222 / Baixo Parnaíba'
    df_cad.loc[idx_251, 'Observação'] = 'Corredor 3 (BR-222 / Baixo Parnaíba - Parnaíba/PI). Ciclo TER-SEX.'
    print("[OK] Loja 251 atualizada para 'Rota 3 - BR-222 / Baixo Parnaíba'.")

# Atualizar Loja 211
idx_211 = df_cad[df_cad['Filial_str'] == '211'].index
if len(idx_211) > 0:
    df_cad.loc[idx_211, 'Restrição Veículo'] = 'Somente Carreta'
    df_cad.loc[idx_211, 'Observação'] = 'Polo rota CE / Piripiri (Transferência Somente Carreta).'
    print("[OK] Loja 211 atualizada para 'Somente Carreta'.")

# Inserir coluna Dias Geração
df_cad['Dias Geração'] = df_cad['Filial_str'].map(map_ciclo).fillna("Não Informado")

# Reordenar colunas
cols = [
    'Filial', 'Nome Loja', 'Região', 'Rota Padrão', 'Dias Geração', 'Restrição Veículo',
    'Bairro', 'Cidade', 'UF', 'Endereço / Referência', 'Latitude', 'Longitude', 'Observação'
]
df_cad = df_cad[[c for c in cols if c in df_cad.columns]]

# Salvar
df_cad.to_excel(caminho_cad, index=False)
print(f"[OK] Base de endereços salva com sucesso com coluna 'Dias Geração' ({len(df_cad)} lojas).")
