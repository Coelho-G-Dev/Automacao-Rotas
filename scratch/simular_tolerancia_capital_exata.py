import sys, os
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))
import itertools
import pandas as pd
from automacao import (
    obter_base_enderecos, carregar_dados_frios, localizar_aba_e_cabecalho,
    normalizar_nome_rota, PARAMETROS_FROTA, obter_veiculos_compativeis
)

aba, row = localizar_aba_e_cabecalho('Acomp_EvoluSep.xlsx')
df = carregar_dados_frios('Acomp_EvoluSep.xlsx', aba, row)
base = obter_base_enderecos(df, 'Cadastro_Lojas_Enderecos.xlsx')

for idx in df.index:
    f = df.at[idx, 'FILIAL_PADRAO']
    if f in base:
        r_cad = base[f].get('Rota Padrão')
        reg_cad = base[f].get('Região')
        if r_cad and str(r_cad).strip() not in ('', 'nan', 'None'):
            df.at[idx, 'ROTA_PADRAO'] = normalizar_nome_rota(r_cad)
        if reg_cad and str(reg_cad).strip() not in ('', 'nan', 'None'):
            df.at[idx, 'REGIAO_PADRAO'] = str(reg_cad).strip().upper()

def selecionar_veiculo_com_tolerancia(peso: float, filiais: list = None, regiao: str = "CAPITAL"):
    veiculos_validos = obter_veiculos_compativeis(filiais or [], base, regiao)
    if regiao == "CAPITAL":
        lojas_unicas = set(str(f).split()[0].replace("(Parcial)", "").strip() for f in (filiais or []))
        if len(lojas_unicas) > 1 and "Carreta" in veiculos_validos:
            veiculos_validos = [v for v in veiculos_validos if v != "Carreta"]

    # 1. Tentar faixa estrita ideal primeiro
    for v in PARAMETROS_FROTA:
        if v["tipo"] in veiculos_validos and regiao in v.get("regioes", ["CAPITAL", "INTERIOR"]):
            if v["peso_minimo"] <= peso <= v["capacidade"]:
                return v["tipo"], False, 0.0

    # 2. Se for CAPITAL, aplicar tolerância de 500 a 1000 kg para piso (3/4 a partir de 5000 kg)
    # NUNCA excedendo a capacidade máxima
    if regiao == "CAPITAL":
        for v in PARAMETROS_FROTA:
            if v["tipo"] in veiculos_validos and "CAPITAL" in v.get("regioes", ["CAPITAL", "INTERIOR"]):
                if v["tipo"] == "3/4":
                    piso_tol = 4950.0 # tolerância de ~500kg para 5.500kg
                else:
                    piso_tol = v["peso_minimo"] - 1000.0 # tolerância de até 1.000kg
                
                # NUNCA exceder a capacidade
                if piso_tol <= peso <= v["capacidade"]:
                    falta = max(0.0, v["peso_minimo"] - peso)
                    return v["tipo"], True, falta

    return None, False, 0.0

# Testar para cada rota da Capital
lojas_processadas = []
col_d = "DIAS_CARREGAMENTO"
for (reg, rota, f), grp in df.groupby(["REGIAO_PADRAO", "ROTA_PADRAO", "FILIAL_PADRAO"]):
    p_total = float(grp["PESOBRUTO"].sum())
    p_hoje = float(grp[grp[col_d].astype(str).str.contains("01-DE HOJE|HOJE", case=False, na=False)]["PESOBRUTO"].sum())
    p_prazo = float(grp[grp[col_d].astype(str).str.contains("02-NO PRAZO|PRAZO", case=False, na=False)]["PESOBRUTO"].sum())
    p_urgente = float(grp[grp[col_d].astype(str).str.contains("03-URGENTE|URGENTE", case=False, na=False)]["PESOBRUTO"].sum())
    tem_hoje = (p_hoje > 0)
    lojas_processadas.append({
        "regiao": reg, "rota": rota, "filial": f, "filial_orig": f,
        "peso": p_total, "peso_hoje": p_hoje, "tem_hoje": tem_hoje
    })

df_lojas = pd.DataFrame(lojas_processadas)
pares_rotas = df_lojas[["regiao", "rota"]].drop_duplicates().copy()
pares_rotas["ordem_reg"] = pares_rotas["regiao"].apply(lambda r: 0 if "CAPITAL" in str(r).upper() else 1)
pares_rotas["num_rota"] = pares_rotas["rota"].apply(lambda r: int(''.join(filter(str.isdigit, str(r)))) if any(c.isdigit() for c in str(r)) else 999)
pares_rotas = pares_rotas.sort_values(by=["ordem_reg", "num_rota", "rota"])

print("\n" + "=" * 95)
print("SIMULAÇÃO DE ROTEIRIZAÇÃO COM REGRA DE 3/4 A 5000 KG E TOLERÂNCIA DE 500-1000 KG NA CAPITAL")
print("=" * 95)

for _, r_par in pares_rotas.iterrows():
    reg = r_par["regiao"]
    r_nome = r_par["rota"]
    sub = df_lojas[(df_lojas["regiao"] == reg) & (df_lojas["rota"] == r_nome)]
    lojas = sub.to_dict("records")
    
    # Tratamento 96
    pool = list(lojas)
    veics_ded = []
    for it in list(pool):
        if it["filial_orig"] == "96" and it["peso"] > 14000:
            pool.remove(it)
            veics_ded.append({
                "veiculo": "Truck", "peso": 12000.0,
                "filiais": ["96 (Carga 1)"], "excecao": False, "falta": 0.0
            })
            pool.append({**it, "filial": "96 (Sobra)", "peso": it["peso"] - 12000.0})

    # Combinatória
    n = len(pool)
    indices = list(range(n))
    candidatos = []
    for k in range(min(4, n), 0, -1):
        for combo in itertools.combinations(indices, k):
            grp = [pool[i] for i in combo]
            if reg == "CAPITAL" and not any(s.get("tem_hoje", False) for s in grp):
                continue
            p = sum(s["peso"] for s in grp)
            f_list = [s["filial_orig"] for s in grp]
            v, exc, falta = selecionar_veiculo_com_tolerancia(p, f_list, reg)
            if v is not None:
                p_obrig = sum(s["peso"] for s in grp if s.get("tem_hoje", True))
                # Priorizar: mais pedidos do dia, menor exceção, maior peso
                candidatos.append((grp, v, p, p_obrig, exc, falta, combo))

    # Greedy search
    candidatos.sort(key=lambda x: (x[3], not x[4], x[2]), reverse=True)
    usados = set()
    veics_rota = list(veics_ded)
    for grp, v, p, p_obrig, exc, falta, combo in candidatos:
        if not any(i in usados for i in combo):
            for i in combo:
                usados.add(i)
            veics_rota.append({
                "veiculo": v, "peso": p, "filiais": [s["filial"] for s in grp],
                "excecao": exc, "falta": falta
            })

    sobras = [pool[i] for i in indices if i not in usados]
    
    for v in veics_rota:
        status_txt = f"EXCEÇÃO CAPITAL (-{v['falta']:,.0f} kg do piso ideal)" if v["excecao"] else "FAIXA IDEAL"
        print(f"[{reg:<8}] {r_nome:<15} | Veículo: {v['veiculo']:<7} | Peso: {v['peso']:>9.2f} kg | Status: {status_txt:<35} | Lojas: {v['filiais']}")
    
    if sobras:
        p_sob = sum(s["peso"] for s in sobras)
        print(f"   --> SOBRAS {r_nome}: {p_sob:,.2f} kg - Lojas: {[s['filial'] for s in sobras]}")
