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

df_lojas = df_frios.groupby(['REGIAO_PADRAO', 'ROTA_PADRAO', 'FILIAL_PADRAO'], as_index=False)['PESOBRUTO'].sum()

print("=== TESTANDO ALOCAÇÃO PURA ESTRITA (ZERO MISTURA DE ROTAS) ===")

def resolver_rota_pura(rota_nome, regiao, lojas_pool, base_enderecos):
    """
    Tenta encontrar a melhor combinação de veículos para uma rota pura específica.
    Cada veículo gerado DEVE conter apenas lojas desta rota e respeitar os limites de peso e restrições.
    Retorna (veiculos_gerados, lojas_restantes).
    """
    veiculos = []
    pool = list(lojas_pool)
    
    # 1. Tratar desdobramento especial (ex: Loja 96 se > 14t)
    for it in list(pool):
        if it["filial_orig"] == "96" and it["peso"] > 14000:
            pool.remove(it)
            # Carga 1: 12.000 kg Truck dedicado
            veiculos.append({
                "rota_padrao": rota_nome,
                "regiao": regiao,
                "veiculo": "Truck",
                "peso": 12000.0,
                "filiais": ["96 (Carga 1)"],
                "stores": [{"filial": "96 (Carga 1)", "filial_orig": "96", "peso": 12000.0, "rota": rota_nome}]
            })
            # Sobra de 96 volta para o pool da Rota 4
            pool.append({
                "filial": "96 (Sobra)",
                "filial_orig": "96",
                "peso": it["peso"] - 12000.0,
                "rota": rota_nome
            })
            
    # Função para buscar combinações gulosas/ótimas dentro do pool da rota
    # Queremos maximizar o peso alocado em veículos válidos
    melhor_solucao = []
    melhor_sobra = list(pool)
    
    # Algoritmo de busca recursiva para encontrar a combinação de veículos que maximiza peso alocado
    def buscar_veiculos(rem_pool, veics_acum):
        nonlocal melhor_solucao, melhor_sobra
        # Se alocou mais peso ou mesmo peso com menos veículos
        peso_alocado_acum = sum(v["peso"] for v in veics_acum)
        peso_alocado_melhor = sum(v["peso"] for v in melhor_solucao)
        if peso_alocado_acum > peso_alocado_melhor:
            melhor_solucao = list(veics_acum)
            melhor_sobra = list(rem_pool)
        elif peso_alocado_acum == peso_alocado_melhor and len(veics_acum) < len(melhor_solucao):
            melhor_solucao = list(veics_acum)
            melhor_sobra = list(rem_pool)
            
        n = len(rem_pool)
        if n == 0:
            return
            
        indices = list(range(n))
        # Tentar formar 1 veículo com k lojas (1 a 4)
        encontrou_algum = False
        # Ordenar tentativas por tamanho decrescente de peso
        candidatos = []
        for k in range(min(automacao.LIMITE_MAX_LOJAS_POR_ROTA, n), 0, -1):
            for combo in itertools.combinations(indices, k):
                grp = [rem_pool[i] for i in combo]
                p = sum(s["peso"] for s in grp)
                f_list = [s["filial_orig"] for s in grp]
                v = automacao.selecionar_menor_veiculo(p, f_list, regiao, base_enderecos)
                if v is not None:
                    candidatos.append((grp, v, p, combo))
                    
        # Ordenar candidatos priorizando veículos maiores / maior peso alocado
        candidatos.sort(key=lambda x: x[2], reverse=True)
        
        for grp, v, p, combo in candidatos[:8]: # Podar para manter performance
            novo_rem = [rem_pool[i] for i in indices if i not in combo]
            novo_v = {
                "rota_padrao": rota_nome,
                "regiao": regiao,
                "veiculo": v,
                "peso": p,
                "filiais": [s["filial"] for s in grp],
                "stores": grp
            }
            buscar_veiculos(novo_rem, veics_acum + [novo_v])

    buscar_veiculos(pool, [])
    
    return veiculos + melhor_solucao, melhor_sobra

# Testar para todas as rotas
rotas_veiculos = []
todas_sobras = []

for rota in sorted(df_lojas['ROTA_PADRAO'].unique()):
    sub = df_lojas[df_lojas['ROTA_PADRAO'] == rota]
    reg = sub['REGIAO_PADRAO'].iloc[0]
    lojas = [{"filial": r.FILIAL_PADRAO, "filial_orig": r.FILIAL_PADRAO, "peso": float(r.PESOBRUTO), "rota": rota} for _, r in sub.iterrows()]
    
    veics, sobras = resolver_rota_pura(rota, reg, lojas, cad)
    for idx_v, v in enumerate(veics, 1):
        v["rota_execucao"] = f"{rota} - Veículo {idx_v}"
        rotas_veiculos.append(v)
    for s in sobras:
        todas_sobras.append(s)

print(f"\nTOTAL DE VEÍCULOS GERADOS: {len(rotas_veiculos)}")
peso_total_veics = sum(v["peso"] for v in rotas_veiculos)
peso_total_sobras = sum(s["peso"] for s in todas_sobras)
print(f"PESO TOTAL EM VEÍCULOS: {peso_total_veics:,.2f} kg")
print(f"PESO TOTAL AGUARDANDO PISO: {peso_total_sobras:,.2f} kg")
print(f"SOMA GERAL: {peso_total_veics + peso_total_sobras:,.2f} kg (Original: 158,885.32 kg)\n")

print("--- VEÍCULOS FORMADOS (TODOS 100% PUROS) ---")
for v in rotas_veiculos:
    paradas = automacao.ordenar_paradas(v["filiais"], v["regiao"], cad)
    paradas_str = ", ".join(paradas)
    detalhes_p = [f"L{s['filial']} ({s['peso']:,.2f}kg)" for s in v["stores"]]
    print(f"[{v['regiao']}] {v['rota_execucao']} | {v['veiculo']:7s} | {v['peso']:9,.2f} kg | Lojas ({len(v['filiais'])}): {paradas_str}")
    print(f"      Pesos: {', '.join(detalhes_p)}")

print("\n--- LOJAS AGUARDANDO GERAÇÃO (BATER PISO) ---")
for s in sorted(todas_sobras, key=lambda x: (x['rota'], -x['peso'])):
    valvo = automacao.obter_menor_veiculo_alvo([s['filial_orig']], cad, s.get('regiao', 'CAPITAL'))
    print(f"[{s['rota']}] Loja {s['filial']:10s}: {s['peso']:8,.2f} kg | Alvo: {valvo['tipo']} ({valvo['peso_minimo']:,.0f} kg)")
