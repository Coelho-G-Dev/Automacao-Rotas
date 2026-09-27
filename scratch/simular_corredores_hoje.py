import sys, os
sys.stdout.reconfigure(encoding='utf-8')
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))
import pandas as pd
import itertools
from automacao import (
    localizar_aba_e_cabecalho, carregar_dados_frios, obter_base_enderecos,
    PARAMETROS_FROTA, obter_veiculos_compativeis
)

aba, row = localizar_aba_e_cabecalho('Acomp_EvoluSep.xlsx')
df = carregar_dados_frios('Acomp_EvoluSep.xlsx', aba, row)
base = obter_base_enderecos(df, 'Cadastro_Lojas_Enderecos.xlsx')

col_d = "DIAS_CARREGAMENTO" if "DIAS_CARREGAMENTO" in df.columns else "Dias Carregamento"

# Agrupar lojas do interior com dados de 26/09
lojas_int = {}
for idx, row_f in df.iterrows():
    f = row_f['FILIAL_PADRAO']
    cad = base.get(f, {})
    reg = cad.get('Região', '').strip().upper()
    if reg == 'INTERIOR':
        p = float(row_f['PESOBRUTO'])
        dias = str(row_f.get(col_d, ''))
        eh_hoje = '01-DE HOJE' in dias or 'HOJE' in dias
        if f not in lojas_int:
            lojas_int[f] = {
                'filial': f, 'nome': cad.get('Nome Loja', f), 'cidade': cad.get('Cidade', ''),
                'peso_total': 0.0, 'peso_hoje': 0.0, 'tem_hoje': False
            }
        lojas_int[f]['peso_total'] += p
        if eh_hoje:
            lojas_int[f]['peso_hoje'] += p
            lojas_int[f]['tem_hoje'] = True

print(f"Total de Lojas do Interior no dia 26/09: {len(lojas_int)}")

# Aplicar corredores
from definir_corredores_interior import CORREDORES_INTERIOR

# Testar alocação por corredor
print("\n" + "=" * 90)
print("SIMULAÇÃO DE VEÍCULOS POR CORREDOR COM AS REGRAS OFICIAIS")
print("=" * 90)

faixas_veiculos = {
    "Toco": (8000, 9000),
    "Truck": (12000, 14000),
    "Bitruck": (15500, 16000),
    "Carreta": (24000, 28000)
}

def avaliar_veiculo(peso, lojas):
    # Carreta no interior: max 4 lojas
    for v_nome in ["Toco", "Truck", "Bitruck", "Carreta"]:
        f_min, f_max = faixas_veiculos[v_nome]
        # tolerância leve de 500kg para o piso se for truck/bitruck
        if f_min - 500 <= peso <= f_max:
            if v_nome == "Carreta" and len(lojas) > 4:
                continue
            return v_nome
    return None

# 1. Corredor 1: Lençóis
print("\n>>> CORREDOR 1: LENÇÓIS / LITORAL NORTE (Sem Santa Rita)")
lojas_lencois = [lojas_int[f] for f in ["222", "424", "434", "537", "447"] if f in lojas_int]
for l in lojas_lencois:
    print(f"   • Loja {l['filial']} ({l['cidade']}): {l['peso_total']:,.2f} kg (Hoje: {l['peso_hoje']:,.2f} kg)")

# Testar se [424, 447] fecha Truck
p_424_447 = lojas_int['424']['peso_total'] + lojas_int['447']['peso_total']
v_424_447 = avaliar_veiculo(p_424_447, ['424', '447'])
print(f"   --> Par [424, 447] (Rosário + Tutóia): {p_424_447:,.2f} kg -> Veículo: {v_424_447}")

# Testar [222, 434] (Mix Rosário + Barreirinhas)
p_222_434 = lojas_int['222']['peso_total'] + lojas_int['434']['peso_total']
print(f"   --> Par [222, 434] (Rosário + Barreirinhas): {p_222_434:,.2f} kg -> Total: 30t (Gera Carreta 24-28t com split de Rosário!)")

# 2. Corredor 3: BR-222 / Baixo Parnaíba
print("\n>>> CORREDOR 3: BR-222 / BAIXO PARNAÍBA (Vargem Grande, Urbano Santos, Chapadinha)")
lojas_222 = [lojas_int[f] for f in ["451", "461", "39"] if f in lojas_int]
for l in lojas_222:
    print(f"   • Loja {l['filial']} ({l['cidade']}): {l['peso_total']:,.2f} kg (Hoje: {l['peso_hoje']:,.2f} kg)")
p_451_461 = lojas_int['451']['peso_total'] + lojas_int['461']['peso_total']
v_451_461 = avaliar_veiculo(p_451_461, ['451', '461'])
print(f"   --> Par [451, 461] (Vargem Grande + Urbano Santos): {p_451_461:,.2f} kg -> Veículo: {v_451_461}")

# 3. Corredor 4: Centro / Mearim / Cocais
print("\n>>> CORREDOR 4: CENTRO / MEARIM / COCAIS (São Mateus, Coroatá, Codó)")
lojas_centro = [lojas_int[f] for f in ["220", "435", "202"] if f in lojas_int]
for l in lojas_centro:
    print(f"   • Loja {l['filial']} ({l['cidade']}): {l['peso_total']:,.2f} kg (Hoje: {l['peso_hoje']:,.2f} kg)")
p_centro = sum(l['peso_total'] for l in lojas_centro)
v_centro = avaliar_veiculo(p_centro, ['220', '435', '202'])
print(f"   --> Trio [220, 435, 202] (São Mateus + Coroatá + Codó): {p_centro:,.2f} kg -> Veículo: {v_centro}")

# 4. Corredor 2: Baixada
print("\n>>> CORREDOR 2: BAIXADA MARANHENSE (Com possibilidade de Santa Rita 429 e Miranda 439)")
lojas_baixada = [lojas_int[f] for f in ["457", "459", "99", "429"] if f in lojas_int]
for l in lojas_baixada:
    print(f"   • Loja {l['filial']} ({l['cidade']}): {l['peso_total']:,.2f} kg (Hoje: {l['peso_hoje']:,.2f} kg)")
p_baixada = sum(l['peso_total'] for l in lojas_baixada)
print(f"   --> Total Baixada hoje com Santa Rita: {p_baixada:,.2f} kg (Fica aguardando bater peso mínimo de 8.000 kg)")
