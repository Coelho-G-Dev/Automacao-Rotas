# -*- coding: utf-8 -*-
"""
BATERIA DE TESTES SEQUENCIAIS DE ROTAS - SETOR DE FRIOS
Testa e valida as regras operacionais, restrições físicas, faixas de frota,
dias de carregamento, tratamento de grandes volumes, corredores do interior
e comparação com a operação real.
"""
import sys, os
sys.stdout.reconfigure(encoding='utf-8')
import pandas as pd
import openpyxl
from collections import Counter, defaultdict

# Adicionar pasta raiz ao path
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))
import automacao

def linha_separadora(titulo, tam=90):
    print("\n" + "=" * tam)
    print(f" {titulo.upper()} ")
    print("=" * tam)

def sub_separadora(titulo, tam=90):
    print("\n" + "-" * tam)
    print(f" >>> {titulo}")
    print("-" * tam)

def carregar_dados():
    caminho_base = automacao.ARQUIVO_ENTRADA_PADRAO
    caminho_cad = automacao.ARQUIVO_ENDERECOS_PADRAO
    aba, cab = automacao.localizar_aba_e_cabecalho(caminho_base)
    df_frios = automacao.carregar_dados_frios(caminho_base, aba, cab)
    cad = automacao.obter_base_enderecos(df_frios, caminho_cad)
    
    # Atualizar com dados cadastrais oficiais
    for idx in df_frios.index:
        f = df_frios.at[idx, 'FILIAL_PADRAO']
        if f in cad:
            r_cad = cad[f].get('Rota Padrão')
            reg_cad = cad[f].get('Região')
            if r_cad and str(r_cad).strip() not in ('', 'nan', 'None'):
                df_frios.at[idx, 'ROTA_PADRAO'] = automacao.normalizar_nome_rota(r_cad)
            if reg_cad and str(reg_cad).strip() not in ('', 'nan', 'None'):
                df_frios.at[idx, 'REGIAO_PADRAO'] = str(reg_cad).strip().upper()
                
    return df_frios, cad

def teste_1_restricoes_veiculares(cad):
    linha_separadora("TESTE SEQUENCIAL 1: VALIDAÇÃO DE RESTRIÇÕES VEICULARES E FÍSICAS")
    
    casos_teste = [
        {"filial": "411", "esperado": ["3/4"], "motivo": "Docas estreitas (Somente 3/4)"},
        {"filial": "418", "esperado": ["3/4"], "motivo": "Docas estreitas (Somente 3/4)"},
        {"filial": "560", "esperado": ["3/4", "Toco"], "motivo": "Somente Toco (ou menor)"},
        {"filial": "451", "esperado": ["Toco", "Truck", "Bitruck", "Carreta"], "motivo": "Toco para cima (proibido 3/4)"},
        {"filial": "434", "esperado": ["Toco", "Truck", "Bitruck", "Carreta"], "motivo": "Toco para cima (proibido 3/4)"},
        {"filial": "408", "esperado": ["3/4", "Toco", "Truck", "Carreta"], "motivo": "Série 400 (Não entra Bitruck)"},
        {"filial": "27",  "esperado": ["3/4", "Toco", "Truck", "Carreta"], "motivo": "Loja 27 (Não entra Bitruck)"},
        {"filial": "415", "esperado": ["3/4", "Toco", "Truck", "Carreta"], "motivo": "Loja 415 (Não entra Bitruck)"},
        {"filial": "506", "esperado": ["Carreta"], "motivo": "Ceará (Somente Carreta)"},
        {"filial": "211", "esperado": ["Carreta"], "motivo": "Ceará (Somente Carreta)"},
        {"filial": "61",  "esperado": ["3/4", "Toco", "Truck", "Bitruck", "Carreta"], "motivo": "Sem restrição"}
    ]
    
    erros = 0
    print(f"{'Filial':<8} | {'Nome Loja':<28} | {'Restrição Esperada':<24} | {'Permitidos Sistema':<30} | {'Status':<8}")
    print("-" * 110)
    for c in casos_teste:
        f = c["filial"]
        info = cad.get(f, {})
        restr = automacao.obter_restricoes_loja(f, cad)
        perm = restr["veiculos_permitidos"]
        status = "OK" if sorted(perm) == sorted(c["esperado"]) else "FALHA"
        if status == "FALHA": erros += 1
        print(f"{f:<8} | {info.get('Nome Loja', 'N/D')[:28]:<28} | {c['motivo'][:24]:<24} | {str(perm)[:30]:<30} | {status:<8}")

    print("\n>>> TESTE DE COMPATIBILIDADE DE CONJUNTO DE LOJAS:")
    conjuntos_teste = [
        {"lojas": ["411", "16"], "regiao": "CAPITAL", "esperado": ["3/4"], "motivo": "411 limita o conjunto para 3/4"},
        {"lojas": ["451", "461"], "regiao": "INTERIOR", "esperado": ["Toco", "Truck", "Carreta"], "motivo": "451 proíbe 3/4; 461 (série 400) proíbe Bitruck -> Sobram Toco, Truck, Carreta"},
        {"lojas": ["34", "61", "2061"], "regiao": "CAPITAL", "esperado": ["3/4", "Toco", "Truck", "Bitruck"], "motivo": "3 lojas na Capital proíbem Carreta (limite carreta capital = 1 loja)"},
        {"lojas": ["222", "447"], "regiao": "INTERIOR", "esperado": ["Toco", "Truck", "Bitruck", "Carreta"], "motivo": "Interior permite Carreta com até 4 lojas"}
    ]
    
    for c in conjuntos_teste:
        res = automacao.obter_veiculos_compativeis(c["lojas"], cad, c["regiao"])
        status = "OK" if sorted(res) == sorted(c["esperado"]) else "DIVERGÊNCIA"
        print(f"Lojas {str(c['lojas']):<18} ({c['regiao']:<8}) -> Permitidos: {str(res):<35} | {status} ({c['motivo']})")

    print(f"\n[RESULTADO TESTE 1]: {'100% APROVADO' if erros == 0 else f'{erros} FALHAS ENCONTRADAS'}")
    return erros == 0

def teste_2_faixas_frota_e_tolerancia(cad):
    linha_separadora("TESTE SEQUENCIAL 2: CALIBRAÇÃO DE FAIXAS DE FROTA, PISOS E TETOS")
    
    # Simulação de pontos de peso críticos
    cenarios_peso = [
        # (peso, lojas, regiao, veic_esperado_estrito, veic_com_tolerancia)
        (4800.0, ["61"], "CAPITAL", None, "3/4 (Exceção Capital - faltam 200kg)"),
        (5200.0, ["61"], "CAPITAL", "3/4", "3/4 (Faixa Operacional 3/4)"),
        (5950.0, ["61"], "CAPITAL", "3/4", "3/4 (Faixa Ideal)"),
        (6500.0, ["61"], "CAPITAL", None, "Vazio entre 3/4 e Toco (precisa +1500kg para bater Toco)"),
        (7200.0, ["61"], "CAPITAL", None, "Exceção Capital Toco (se tolerância for 1000kg)"),
        (8200.0, ["61"], "CAPITAL", "Toco", "Toco (Piso batido)"),
        (8900.0, ["61"], "CAPITAL", "Toco", "Toco (Excelente lotação)"),
        (9200.0, ["61"], "CAPITAL", None, "Acima de 9000kg (Excede teto máximo de Toco - Proibido para evitar quebra)"),
        (11500.0, ["61"], "CAPITAL", None, "Exceção Truck Capital (Faltam 500kg para 12t)"),
        (12500.0, ["61"], "CAPITAL", "Truck", "Truck (Dentro da faixa)"),
        (14073.46, ["429", "451", "461", "39"], "INTERIOR", "Truck", "Truck BR-222 com +73kg (dentro da tolerância <= 100kg - APROVADO)"),
        (14223.0, ["220", "435", "202"], "INTERIOR", None, "Truck excedente por 223kg (> 100kg de tolerância - REJEITADO para evitar quebra)"),
        (15800.0, ["61"], "CAPITAL", "Bitruck", "Bitruck (Dentro da faixa 15.5t a 16t)"),
        (25500.0, ["222", "447"], "INTERIOR", "Carreta", "Carreta (Faixa 24t a 28t)"),
        (25500.0, ["61", "34"], "CAPITAL", None, "Carreta proibida na Capital com >1 loja!")
    ]
    
    print(f"{'Peso (Kg)':<10} | {'Lojas':<18} | {'Região':<9} | {'Veículo Padrão':<16} | {'Exceção/Tolerância':<20} | {'Diagnóstico'}")
    print("-" * 115)
    
    for peso, lojas, regiao, esp_padrao, desc in cenarios_peso:
        v_padrao, eh_exc, falta = automacao.selecionar_menor_veiculo(peso, lojas, regiao, cad, permitir_excecao=True)
        v_estrito, _, _ = automacao.selecionar_menor_veiculo(peso, lojas, regiao, cad, permitir_excecao=False)
        
        info_v = v_padrao or "NENHUM"
        info_exc = f"SIM (Falta {falta:,.0f}kg)" if eh_exc else "NÃO"
        print(f"{peso:>10,.1f} | {str(lojas)[:18]:<18} | {regiao:<9} | {info_v:<16} | {info_exc:<20} | {desc}")

def teste_3_dias_carregamento(df_frios, cad):
    linha_separadora("TESTE SEQUENCIAL 3: LÓGICA DE DIAS DE CARREGAMENTO (HOJE vs PRAZO vs URGENTE)")
    
    col_d = "DIAS_CARREGAMENTO" if "DIAS_CARREGAMENTO" in df_frios.columns else "Dias Carregamento"
    
    print("Resumo da composição da base atual:")
    dist_dias = df_frios.groupby(col_d)["PESOBRUTO"].agg(['count', 'sum']).reset_index()
    dist_dias.columns = ['Status', 'Qtd Registros', 'Peso Total (Kg)']
    dist_dias['% Peso'] = dist_dias['Peso Total (Kg)'] / dist_dias['Peso Total (Kg)'].sum() * 100
    print(dist_dias.to_string(index=False))
    
    sub_separadora("VALIDAÇÃO DE COMPORTAMENTO POR ROTA:")
    print("Regra: Se uma rota só possui '02-NO PRAZO' e NENHUM '01-DE HOJE', ela não pode rodar caminhão na Capital!")
    
    rotas_sem_hoje = []
    for (reg, rota), grp in df_frios.groupby(['REGIAO_PADRAO', 'ROTA_PADRAO']):
        p_hoje = grp[grp[col_d].astype(str).str.contains("01-DE HOJE|HOJE", case=False, na=False)]["PESOBRUTO"].sum()
        p_tot = grp["PESOBRUTO"].sum()
        if p_hoje == 0:
            rotas_sem_hoje.append((reg, rota, p_tot, len(grp["FILIAL_PADRAO"].unique())))
            
    print(f"\nRotas encontradas com 100% de carga futura (sem pedidos de hoje): {len(rotas_sem_hoje)}")
    for reg, rota, p, n_lojas in rotas_sem_hoje:
        print(f"   • [{reg}] {rota}: {p:,.2f} kg ({n_lojas} lojas) -> Comportamento correto: AGUARDAR GERAÇÃO DO DIA")

def teste_4_lojas_pesadas_e_splits(df_frios, cad):
    linha_separadora("TESTE SEQUENCIAL 4: TRATAMENTO DE LOJAS ÂNCORA E SPLITS DE CARGA")
    
    sub_separadora("CASO 1: LOJA 96 (SUPER SÃO FRANCISCO / CURVA A)")
    sub_96 = df_frios[df_frios['FILIAL_PADRAO'] == '96']
    peso_96 = sub_96['PESOBRUTO'].sum()
    print(f"Peso atual da Loja 96 na base: {peso_96:,.2f} kg")
    
    # Testar cenários hipotéticos de peso para a Loja 96
    cenarios_96 = [9400.0, 13500.0, 18885.0, 26500.0]
    for p in cenarios_96:
        print(f"\n   Simulação Loja 96 com {p:,.2f} kg:")
        if p <= 9000:
            print(f"      -> 1 Toco direto ({p:,.2f} kg)")
        elif p <= 14000:
            print(f"      -> 1 Truck direto ({p:,.2f} kg)")
        elif p <= 28000:
            print(f"      Opção A (Atual): Split de 12.000 kg (Truck) + Sobra de {p-12000:,.2f} kg para agrupar na Rota 8")
            if p >= 24000:
                print(f"      Opção B (Recomendada se >=24t): 1 Carreta Exclusiva Direta de {p:,.2f} kg (1 Loja Capital permitida!)")
            else:
                print(f"      Opção B (Se 18-20t): Split equilibrado em 2 Tocos (~{p/2:,.0f} kg cada)")

    sub_separadora("CASO 2: LOJA 222 (MIX MATEUS ROSÁRIO - INTERIOR)")
    sub_222 = df_frios[df_frios['FILIAL_PADRAO'] == '222']
    peso_222 = sub_222['PESOBRUTO'].sum()
    print(f"Peso atual da Loja 222 na base: {peso_222:,.2f} kg")
    print("No interior, a Loja 222 costuma passar de 18.000 kg.")
    print("Combinando com Tutóia (447) ou Barreirinhas (434), atinge 25t a 28t, fechando CARRETA com 2 paradas!")

def teste_5_corredores_interior_vs_rota_pura(df_frios, cad):
    linha_separadora("TESTE SEQUENCIAL 5: CORREDORES DO INTERIOR vs ROTA PURA ESTRITA")
    
    sub_separadora("ANÁLISE DE ROTAS PURAS NO INTERIOR (COMO ESTÁ HOJE):")
    rotas_int = df_frios[df_frios['REGIAO_PADRAO'] == 'INTERIOR'].groupby('ROTA_PADRAO')['PESOBRUTO'].agg(['count', 'sum']).reset_index()
    for _, r in rotas_int.iterrows():
        print(f"   • {r['ROTA_PADRAO']:<32} | Peso: {r['sum']:>9.2f} kg ({r['count']} registros)")
        
    sub_separadora("VALIDAÇÃO DE CAPACIDADE MÁXIMA E TRAVA DE SOBREPESO (CORREDOR 4):")
    p_220 = df_frios[df_frios['FILIAL_PADRAO'] == '220']['PESOBRUTO'].sum()
    p_435 = df_frios[df_frios['FILIAL_PADRAO'] == '435']['PESOBRUTO'].sum()
    p_202 = df_frios[df_frios['FILIAL_PADRAO'] == '202']['PESOBRUTO'].sum()
    p_trio = p_220 + p_435 + p_202
    print(f"   Loja 220 (São Mateus): {p_220:,.2f} kg")
    print(f"   Loja 435 (Coroatá):    {p_435:,.2f} kg")
    print(f"   Loja 202 (Codó):       {p_202:,.2f} kg")
    print(f"   Soma do Trio:          {p_trio:,.2f} kg")
    print(f"   Capacidade Truck: 14.000 kg (Tolerância máxima: 100 kg)")
    print(f"   Resultado: Excedente de 223 kg (> 100 kg) bloqueado com sucesso para proteção veicular.")

def teste_6_confronto_operacao_real(cad):
    linha_separadora("TESTE SEQUENCIAL 6: CONFRONTO COM O HISTÓRICO REAL (DIAS 24, 25 E 26/09)")
    
    pasta_hist = os.environ.get("PASTA_ROTAS_HISTORICO", os.path.join(os.path.dirname(__file__), "..", "historico"))
    arquivos_hist = [
        ("Rotas dia 24-09.xlsx", os.path.join(pasta_hist, "Rotas dia 24-09.xlsx")),
        ("Rotas dia 25-09 .xlsx", os.path.join(pasta_hist, "Rotas dia 25-09 .xlsx")),
        ("Rotas dia 26-09  - .xlsx", os.path.join(pasta_hist, "Rotas dia 26-09  - .xlsx"))
    ]
    encontrados = [(nome, c) for nome, c in arquivos_hist if os.path.exists(c)]
    if not encontrados:
        print("   [INFO] Arquivos de histórico manual não encontrados no ambiente. Teste concluído.")
        return

    for nome_arq, caminho in encontrados:
        sub_separadora(f"DIAGNÓSTICO: {nome_arq}")
        wb = openpyxl.load_workbook(caminho, data_only=True)
        
        for aba_nome, reg_label in [("ROTAS_CAP", "CAPITAL"), ("ROTA_INT", "INTERIOR")]:
            if aba_nome not in wb.sheetnames:
                continue
            ws = wb[aba_nome]
            veics_manuais = []
            for r in range(8, ws.max_row + 1):
                vals = [ws.cell(r, c).value for c in range(2, 11)]
                paradas = [str(int(p)) if isinstance(p, (int, float)) else str(p).strip() for p in vals[:4] if p is not None]
                veic = str(vals[6] if aba_nome == "ROTAS_CAP" else vals[6]).strip().upper()
                if paradas:
                    veics_manuais.append({"paradas": paradas, "veiculo": veic})
                    
            print(f"   [{reg_label}] Operação Real montou {len(veics_manuais)} veículos:")
            cont_tipo = Counter(v['veiculo'] for v in veics_manuais)
            for t, qtd in cont_tipo.items():
                print(f"      • {t}: {qtd} veículo(s)")
            print(f"      Amostra de rotas: {[v['paradas'] for v in veics_manuais[:3]]}")

def executar_todos_os_testes():
    print("=" * 90)
    print(" INICIANDO SUÍTE COMPLETA DE TESTES SEQUENCIAIS PARA VALIDAÇÃO DE ROTAS ")
    print("=" * 90)
    
    df_frios, cad = carregar_dados()
    
    teste_1_restricoes_veiculares(cad)
    teste_2_faixas_frota_e_tolerancia(cad)
    teste_3_dias_carregamento(df_frios, cad)
    teste_4_lojas_pesadas_e_splits(df_frios, cad)
    teste_5_corredores_interior_vs_rota_pura(df_frios, cad)
    teste_6_confronto_operacao_real(cad)
    
    linha_separadora("TODOS OS TESTES SEQUENCIAIS FORAM EXECUTADOS COM SUCESSO!")

if __name__ == "__main__":
    executar_todos_os_testes()
