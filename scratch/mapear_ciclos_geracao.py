# -*- coding: utf-8 -*-
import sys, os
sys.stdout.reconfigure(encoding='utf-8')
import pandas as pd
import openpyxl

df_cad = pd.read_excel('Cadastro_Lojas_Enderecos.xlsx')
df_ciclo = pd.read_excel('Acomp_EvoluSep.xlsx', sheet_name='CIclo', header=4)
col_dias = [c for c in df_ciclo.columns if 'DIAS_GERA' in c.upper()][0]
df_ciclo['Filial'] = df_ciclo['CLIENTE'].dropna().astype(int).astype(str)
map_ciclo = df_ciclo.set_index('Filial')[col_dias].to_dict()

df_cad['Filial_str'] = df_cad['Filial'].astype(str)
df_cad['DIAS_GERACAO'] = df_cad['Filial_str'].map(map_ciclo)

int_cad = df_cad[df_cad['Região'].str.upper() == 'INTERIOR']
for r_nome, grp in int_cad.groupby('Rota Padrão'):
    print('='*75)
    print(r_nome)
    for _, row in grp.iterrows():
        f = row['Filial']
        n = str(row['Nome Loja'])[:25]
        c = str(row['DIAS_GERACAO'])
        cid = str(row['Cidade'])
        print(f"   Loja {f:<4} ({n:<25}) | Ciclo: {c:<15} | Cidade: {cid}")
