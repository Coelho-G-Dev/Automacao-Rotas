# -*- coding: utf-8 -*-
import sys
sys.path.append(r"c:\Users\gabri\Desktop\Separacao")
import automacao
import pandas as pd
import itertools
import time

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

df_lojas = df_frios.groupby(['REGIAO_PADRAO', 'ROTA_PADRAO', 'FILIAL_PADRAO'], as_index=False)['PESOBRUTO'].sum()

def otimizar_veiculos_rota_pura(rota_nome: str, regiao: str, pool_lojas: list, base_enderecos: dict):
    pool = list(pool_lojas)
    veiculos_dedicados = []
    
    # Tratamento especial de cargas gigantes (> 14t) como Loja 96
    for it in list(pool):
        if it["filial_orig"] == "96" and it["peso"] > 14000:
            pool.remove(it)
            veiculos_dedicados.append({
                "rota_padrao": rota_nome,
                "regiao": regiao,
                "veiculo": "Truck",
                "peso": 12000.0,
                "filiais": ["96 (Carga 1)"],
                "stores": [{"filial": "96 (Carga 1)", "filial_orig": "96", "peso": 12000.0, "rota": rota_nome}],
                "tipo_alocacao": "Dedicado Loja 96 (12t)"
            })
            pool.append({
                "filial": "96 (Sobra)",
                "filial_orig": "96",
                "peso": it["peso"] - 12000.0,
                "rota": rota_nome
            })
            
    # Branch and bound para maximizar peso alocado
    melhor_veiculos = []
    melhor_sobras = list(pool)
    melhor_peso_total = 0.0
    
    def buscar(rem_pool, veics_acum):
        nonlocal melhor_veiculos, melhor_sobras, melhor_peso_total
        p_acum = sum(v["peso"] for v in veics_acum)
        if p_acum > melhor_peso_total:
            melhor_peso_total = p_acum
            melhor_veiculos = list(veics_acum)
            melhor_sobras = list(rem_pool)
        elif p_acum == melhor_peso_total and len(veics_acum) < len(melhor_veiculos):
            melhor_veiculos = list(veics_acum)
            melhor_sobras = list(rem_pool)
            
        n = len(rem_pool)
        if n == 0:
            return
            
        indices = list(range(n))
        candidatos = []
        for k in range(min(automacao.LIMITE_MAX_LOJAS_POR_ROTA, n), 0, -1):
            for combo in itertools.combinations(indices, k):
                grp = [rem_pool[i] for i in combo]
                p = sum(s["peso"] for s in grp)
                f_list = [s["filial_orig"] for s in grp]
                v = automacao.selecionar_menor_veiculo(p, f_list, regiao, base_enderecos)
                if v is not None:
                    candidatos.append((grp, v, p, combo))
                    
        # Ordenar candidatos priorizando maior peso
        candidatos.sort(key=lambda x: x[2], reverse=True)
        
        for grp, v, p, combo in candidatos[:12]:
            novo_rem = [rem_pool[i] for i in indices if i not in combo]
            novo_v = {
                "rota_padrao": rota_nome,
                "regiao": regiao,
                "veiculo": v,
                "peso": p,
                "filiais": [s["filial"] for s in grp],
                "stores": grp,
                "tipo_alocacao": "Rota Pura"
            }
            buscar(novo_rem, veics_acum + [novo_v])
            
    t0 = time.time()
    buscar(pool, [])
    t1 = time.time()
    
    return veiculos_dedicados + melhor_veiculos, melhor_sobras, (t1 - t0)

print("\n=== TESTE DE PERFORMANCE E RESULTADOS POR ROTA PURA ===")
total_veics = 0
total_p_veics = 0.0
total_p_sobra = 0.0

for rota in sorted(df_lojas['ROTA_PADRAO'].unique()):
    sub = df_lojas[df_lojas['ROTA_PADRAO'] == rota]
    reg = sub['REGIAO_PADRAO'].iloc[0]
    lojas = [{"filial": r.FILIAL_PADRAO, "filial_orig": r.FILIAL_PADRAO, "peso": float(r.PESOBRUTO), "rota": rota} for _, r in sub.iterrows()]
    veics, sobras, elapsed = otimizar_veiculos_rota_pura(rota, reg, lojas, cad)
    
    total_veics += len(veics)
    p_v = sum(v["peso"] for v in veics)
    p_s = sum(s["peso"] for s in sobras)
    total_p_veics += p_v
    total_p_sobra += p_s
    
    print(f"Rota {rota:12s} ({reg:8s}) | Veículos: {len(veics)} ({p_v:9,.2f} kg) | Pendente: {len(sobras)} ({p_s:9,.2f} kg) | Tempo: {elapsed:.3f}s")
    for v in veics:
        paradas = automacao.ordenar_paradas(v["filiais"], reg, cad)
        print(f"   -> {v['veiculo']:7s} ({v['peso']:9,.2f} kg) | Paradas ({len(paradas)}): {', '.join(paradas)}")

print(f"\nTOTAL: {total_veics} veículos | {total_p_veics:,.2f} kg alocados | {total_p_sobra:,.2f} kg pendentes")
