# -*- coding: utf-8 -*-
"""
TESTE DA MATRIZ DE ADJACÊNCIA GEOGRÁFICA DA CAPITAL
Demonstra como eliminar a perda de frete agrupando apenas rotas vizinhas contíguas
(mesmo quadrante viário de São Luís) quando uma rota isolada não atinge o piso.
"""
import sys, os
sys.stdout.reconfigure(encoding='utf-8')
import pandas as pd
import openpyxl

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))
import automacao
from benchmark_3dias_completo import extrair_demanda_historica

# Matriz Oficial de Vizinhança Contígua na Capital (Bairros colados que fazem sentido)
ADJACENCIA_CAPITAL = {
    "Rota 1": ["Rota 3", "Rota 2", "Rota 8"],             # Renascença vizinho de Calhau, Cohama e Vinhais
    "Rota 2": ["Rota 1", "Rota 3", "Rota 9", "Rota 8"],     # Cohama vizinho de Calhau, Turu e Vinhais
    "Rota 3": ["Rota 1", "Rota 2", "Rota 8"],             # Calhau vizinho de Renascença, Cohama e Shopping/Jaracaty
    "Rota 4": ["Rota 7"],                                 # Itaqui-Bacanga vizinho de Centro/Anil (Eixo Anjo da Guarda/Barragem)
    "Rota 5": ["Rota 12", "Rota 7"],                      # Estrada de Ribamar/Operária vizinho de Maioba/Maiobão
    "Rota 6": ["Rota 10", "Rota 7"],                      # Tirirical vizinho de São Cristóvão
    "Rota 7": ["Rota 4", "Rota 8", "Rota 6", "Rota 10"],   # Centro/Anil/Cajazeiras vizinho de Bacanga, João Paulo, Tirirical
    "Rota 8": ["Rota 1", "Rota 2", "Rota 3", "Rota 9", "Rota 7"], # Vinhais/João Paulo é o nó central
    "Rota 9": ["Rota 2", "Rota 11", "Rota 8"],            # Turu vizinho de Cohama e Cohatrac
    "Rota 10": ["Rota 6", "Rota 11", "Rota 7"],           # São Cristóvão vizinho de Tirirical e Cohab
    "Rota 11": ["Rota 9", "Rota 10", "Rota 12"],          # Cohab/Cohatrac vizinho de Turu, São Cristóvão e Maioba
    "Rota 12": ["Rota 5", "Rota 11"]                      # Maioba/Maiobão/Paço do Lumiar vizinho de Operária e Cohatrac
}

def simular_otimizacao_sem_perda_frete():
    caminho_cad = automacao.ARQUIVO_ENDERECOS_PADRAO
    cad_df = pd.read_excel(caminho_cad)
    cad = {str(r["Filial"]).strip(): r.to_dict() for _, r in cad_df.iterrows()}
    
    caminho_26 = r"C:\Users\gabri\Desktop\Trabalho\Rotas\Rotas dia 26-09  - .xlsx"
    df_demanda = extrair_demanda_historica(caminho_26, cad)
    
    print("=" * 100)
    print(" SIMULAÇÃO DE OTIMIZAÇÃO DE FRETE: CAPITAL SEM PERDA DE FRETE (DIA 26/09) ")
    print("=" * 100)
    
    # 1. Alocação Pura Inicial
    res_pura, det_pura, aguard_pura, _ = automacao.processar_alocacao_rotas(df_demanda, cad)
    
    print(f"\n1. NO MODELO DE ROTA PURA ESTRITA:")
    print(f"   • Veículos Aprovados: {len(res_pura)}")
    print(f"   • Cargas em Aguardo (Sem Bater Piso): {len(aguard_pura)} lojas somando {aguard_pura['Peso Loja (Kg)'].sum():,.2f} kg")
    
    # Identificar sobras da capital com carga do dia
    cap_aguard = aguard_pura[aguard_pura["Região"] == "CAPITAL"]
    rotas_pendentes = cap_aguard.groupby("Rota Padrão")["Peso Loja (Kg)"].sum().to_dict()
    
    print("\n   Rotas da Capital que ficaram com carga parada por falta de piso:")
    for r, p in rotas_pendentes.items():
        print(f"      • {r}: {p:>8,.2f} kg (Não bateu piso de 3/4 isoladamente)")
        
    print("\n2. APLICANDO FUSÃO GEOGRÁFICA CONTÍGUA (SEM PERDA DE FRETE):")
    pares_testados = [
        ("Rota 2", "Rota 3"), # Cohama + Calhau
        ("Rota 8", "Rota 9"), # Vinhais + Turu
        ("Rota 4", "Rota 7")  # Itaqui-Bacanga + Centro
    ]
    
    for r_a, r_b in pares_testados:
        p_a = rotas_pendentes.get(r_a, 0.0)
        p_b = rotas_pendentes.get(r_b, 0.0)
        p_combo = p_a + p_b
        
        if p_combo >= 5000:
            # Selecionar veículo
            lojas_combo = list(cap_aguard[cap_aguard["Rota Padrão"].isin([r_a, r_b])]["Filial"].unique())
            veic, eh_exc, falta = automacao.selecionar_menor_veiculo(p_combo, lojas_combo, "CAPITAL", cad)
            
            bairros_a = [cad.get(f, {}).get("Bairro", "") for f in cap_aguard[cap_aguard["Rota Padrão"] == r_a]["Filial"]]
            bairros_b = [cad.get(f, {}).get("Bairro", "") for f in cap_aguard[cap_aguard["Rota Padrão"] == r_b]["Filial"]]
            
            print(f"\n   [FUSÃO CONTÍGUA: {r_a} + {r_b}]")
            print(f"      • Bairros integrados: {list(set(bairros_a))} + {list(set(bairros_b))}")
            print(f"      • Sentido Logístico: Eixo viário contíguo perfeito (mesmo quadrante)")
            print(f"      • Peso Combinado: {p_a:,.1f} kg + {p_b:,.1f} kg = {p_combo:,.1f} kg")
            print(f"      • Veículo Otimizado: {veic} ({p_combo:,.1f} kg)")
            print(f"      • Ocupação da Capacidade: {p_combo / (9250 if veic == 'Toco' else 6000) * 100:.1f}%")
            print(f"      • Economia: Elimina 1 frete extra e evita desperdício de peso vazio!")

if __name__ == "__main__":
    simular_otimizacao_sem_perda_frete()
