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

# 1. Agrupar dados por loja aplicando a regra do usuário
lojas_processadas = []
for (reg, rota, f), grp in df_frios.groupby(['REGIAO_PADRAO', 'ROTA_PADRAO', 'FILIAL_PADRAO']):
    p_total = float(grp['PESOBRUTO'].sum())
    p_hoje = float(grp[grp[col_dias] == '01-DE HOJE']['PESOBRUTO'].sum())
    p_prazo = float(grp[grp[col_dias] == '02-NO PRAZO']['PESOBRUTO'].sum())
    p_urgente = float(grp[grp[col_dias] == '03-URGENTE']['PESOBRUTO'].sum())
    
    tem_hoje = p_hoje > 0
    if tem_hoje:
        status_dias = "01-DE HOJE"
        tipo_demanda = "Pedido do Dia (Obrigatório)"
        # Se tem status 1 e 2/3, agrupa todo o peso
        obs_demanda = "Pedido do Dia (Peso total agrupado 01-Hoje + 02/03)" if (p_prazo > 0 or p_urgente > 0) else "Pedido do Dia (100% Hoje)"
    elif p_urgente > 0:
        status_dias = "03-URGENTE"
        tipo_demanda = "Complemento de Peso (Urgente)"
        obs_demanda = "Complemento de Peso (03-Urgente)"
    else:
        status_dias = "02-NO PRAZO"
        tipo_demanda = "Complemento de Peso (No Prazo)"
        obs_demanda = "Complemento de Peso (02-No Prazo)"
        
    lojas_processadas.append({
        "regiao": reg,
        "rota": rota,
        "filial": f,
        "filial_orig": f,
        "peso": p_total,
        "peso_hoje": p_hoje,
        "peso_prazo": p_prazo,
        "peso_urgente": p_urgente,
        "tem_hoje": tem_hoje,
        "status_dias": status_dias,
        "tipo_demanda": tipo_demanda,
        "obs_demanda": obs_demanda
    })

def otimizar_rota_com_dias_carregamento(rota_nome: str, regiao: str, pool_lojas: list, base_enderecos: dict):
    pool = list(pool_lojas)
    veiculos_dedicados = []
    
    # Loja 96 (se > 14t)
    for it in list(pool):
        if it["filial_orig"] == "96" and it["peso"] > 14000:
            pool.remove(it)
            veiculos_dedicados.append({
                "rota_padrao": rota_nome,
                "regiao": regiao,
                "veiculo": "Truck",
                "peso": 12000.0,
                "filiais": ["96 (Carga 1)"],
                "stores": [{**it, "filial": "96 (Carga 1)", "peso": 12000.0}],
                "tipo_alocacao": "Dedicado Loja 96 (12t)"
            })
            pool.append({
                **it,
                "filial": "96 (Sobra)",
                "peso": it["peso"] - 12000.0
            })

    # Cargas exclusivas >= 24t
    for s in list(pool):
        if s["peso"] >= 24000:
            v_tipo = automacao.selecionar_menor_veiculo(s["peso"], [s["filial_orig"]], regiao, base_enderecos)
            if v_tipo is not None:
                veiculos_dedicados.append({
                    "rota_padrao": rota_nome,
                    "regiao": regiao,
                    "veiculo": v_tipo,
                    "peso": s["peso"],
                    "filiais": [s["filial"]],
                    "stores": [s],
                    "tipo_alocacao": "Carga Exclusiva / Dedicada"
                })
                pool.remove(s)

    # Busca otimizada:
    # Prioridade 1: Maximizar peso de lojas OBRIGATÓRIAS (tem_hoje == True)
    # Prioridade 2: Maximizar peso total do veículo (usando complementares para atingir piso/meta)
    melhor_veiculos = []
    melhor_sobras = list(pool)
    melhor_score = (-1.0, -1.0)
    
    def buscar(rem_pool, veics_acum):
        nonlocal melhor_veiculos, melhor_sobras, melhor_score
        
        # Calcular score: (peso_obrigatorio_acumulado, peso_total_acumulado)
        p_obrig_acum = sum(sum(s["peso"] for s in v["stores"] if s["tem_hoje"]) for v in veics_acum)
        p_tot_acum = sum(v["peso"] for v in veics_acum)
        score = (p_obrig_acum, p_tot_acum)
        
        if score > melhor_score:
            melhor_score = score
            melhor_veiculos = list(veics_acum)
            melhor_sobras = list(rem_pool)
        elif score == melhor_score and len(veics_acum) < len(melhor_veiculos):
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
                
                # Regra: Na Capital, um veículo só deve ser aberto se contiver pelo menos uma loja obrigatória (tem_hoje)
                # ou se a região for Interior. Lojas puramente complementares não devem abrir veículo sozinhas!
                if regiao == "CAPITAL" and not any(s["tem_hoje"] for s in grp):
                    continue
                    
                p = sum(s["peso"] for s in grp)
                f_list = [s["filial_orig"] for s in grp]
                v = automacao.selecionar_menor_veiculo(p, f_list, regiao, base_enderecos)
                if v is not None:
                    p_obrig = sum(s["peso"] for s in grp if s["tem_hoje"])
                    candidatos.append((grp, v, p, p_obrig, combo))
                    
        # Ordenar candidatos priorizando maior peso obrigatório e depois maior peso total
        candidatos.sort(key=lambda x: (x[3], x[2]), reverse=True)
        
        for grp, v, p, p_obrig, combo in candidatos[:12]:
            novo_rem = [rem_pool[i] for i in indices if i not in combo]
            has_compl = any(not s["tem_hoje"] for s in grp)
            tipo_aloc = "Rota Pura (com complemento de peso)" if has_compl else "Rota Pura (100% Pedido do Dia)"
            novo_v = {
                "rota_padrao": rota_nome,
                "regiao": regiao,
                "veiculo": v,
                "peso": p,
                "filiais": [s["filial"] for s in grp],
                "stores": grp,
                "tipo_alocacao": tipo_aloc
            }
            buscar(novo_rem, veics_acum + [novo_v])
            
    buscar(pool, [])
    return veiculos_dedicados + melhor_veiculos, melhor_sobras

# Teste em todas as rotas
df_lojas_proc = pd.DataFrame(lojas_processadas)
for rota in sorted(df_lojas_proc['rota'].unique()):
    sub = df_lojas_proc[df_lojas_proc['rota'] == rota]
    reg = sub['regiao'].iloc[0]
    lojas = sub.to_dict('records')
    veics, sobras = otimizar_rota_com_dias_carregamento(rota, reg, lojas, cad)
    print(f"\n--- {rota} ({reg}) ---")
    print(f"Veículos gerados: {len(veics)}")
    for v in veics:
        lojas_str = ", ".join([f"L{s['filial']} ({s['peso']:,.2f}kg - {s['status_dias']})" for s in v['stores']])
        print(f"   {v['veiculo']:7s} | {v['peso']:9,.2f} kg | {v['tipo_alocacao']} | Lojas: {lojas_str}")
    if sobras:
        sobras_str = ", ".join([f"L{s['filial']} ({s['peso']:,.2f}kg - {s['status_dias']})" for s in sobras])
        print(f"   Pendentes ({len(sobras)}): {sobras_str}")
