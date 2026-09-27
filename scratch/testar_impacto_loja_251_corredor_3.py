# -*- coding: utf-8 -*-
"""
TESTE DE IMPACTO: LOJA 251 NO CORREDOR 3 (BR-222 / BAIXO PARNAÍBA)
E APLICAÇÃO DOS CICLOS DE GERAÇÃO COM COMPLEMENTO DE SOBRAS.
"""
import sys, os
sys.stdout.reconfigure(encoding='utf-8')
import pandas as pd
import openpyxl

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))
import automacao
from benchmark_3dias_completo import extrair_demanda_historica, extrair_rotas_manuais
from simular_refinamento_rotas import PARAMETROS_FROTA_REFINADOS, obter_restricoes_loja_refinada

def testar():
    caminho_cad = automacao.ARQUIVO_ENDERECOS_PADRAO
    cad_df = pd.read_excel(caminho_cad)
    cad = {str(r["Filial"]).strip(): r.to_dict() for _, r in cad_df.iterrows()}
    
    print("=" * 100)
    print(" VERIFICANDO CADASTRO DA LOJA 251:")
    print("=" * 100)
    print(f"Filial: {cad['251']['Filial']}")
    print(f"Nome: {cad['251']['Nome Loja']}")
    print(f"Rota Padrão: {cad['251']['Rota Padrão']}")
    print(f"Dias Geração: {cad['251'].get('Dias Geração')}")
    print(f"Restrição: {cad['251'].get('Restrição Veículo')}")
    
    # Rodar dia 25/09 (onde a Loja 251 e Loja 39 rodaram juntas)
    caminho_25 = r"C:\Users\gabri\Desktop\Trabalho\Rotas\Rotas dia 25-09 .xlsx"
    df_demanda = extrair_demanda_historica(caminho_25, cad)
    
    # Aplicar parâmetros calibrados
    automacao_orig_frota = automacao.PARAMETROS_FROTA
    automacao_orig_restr = automacao.obter_restricoes_loja
    automacao.PARAMETROS_FROTA = PARAMETROS_FROTA_REFINADOS
    automacao.obter_restricoes_loja = lambda f, b=None: obter_restricoes_loja_refinada(f, b or cad)
    
    res, det, aguard, _ = automacao.processar_alocacao_rotas(df_demanda, cad)
    
    print("\n" + "=" * 100)
    print(" ROTAS DO INTERIOR GERADAS NO DIA 25/09 (COM LOJA 251 NO CORREDOR 3):")
    print("=" * 100)
    sub_int = res[res["Região"] == "INTERIOR"]
    for _, r in sub_int.iterrows():
        print(f"   • {r['Rota de Execução']:<36} | {r['Veículo Ideal']:<8} | {r['Peso Total (Kg)']:>9.2f} kg | Paradas: {r['Filiais Atendidas']}")
        
    print("\nRotas Manuais Feitas pela Operação Real no Dia 25/09 (Interior):")
    manuais = extrair_rotas_manuais(caminho_25)["INTERIOR"]
    for idx, m in enumerate(manuais, 1):
        print(f"   • Carro {idx:02d} | {m['veiculo']:<8} | Paradas: {m['paradas']}")
        
    # Restaurar
    automacao.PARAMETROS_FROTA = automacao_orig_frota
    automacao.obter_restricoes_loja = automacao_orig_restr

if __name__ == "__main__":
    testar()
