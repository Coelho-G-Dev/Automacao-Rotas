import sys, os
sys.stdout.reconfigure(encoding='utf-8')
import pandas as pd
import itertools
from collections import Counter

# 1. Carregar base de lojas e rotas oficiais
caminho_cad = r"C:\Users\gabri\Desktop\Separacao\Cadastro_Lojas_Enderecos.xlsx"
df_cad = pd.read_excel(caminho_cad)

# 2. Definir parâmetros de frota com margem operacional calibrada
PARAMETROS_FROTA = [
    {"tipo": "3/4",     "piso": 5000,  "teto": 6000,  "regioes": ["CAPITAL"]},
    {"tipo": "Toco",    "piso": 8000,  "teto": 9250,  "regioes": ["CAPITAL", "INTERIOR"]},
    {"tipo": "Truck",   "piso": 12000, "teto": 14250, "regioes": ["CAPITAL", "INTERIOR"]},
    {"tipo": "Bitruck", "piso": 15500, "teto": 16000, "regioes": ["CAPITAL", "INTERIOR"]},
    {"tipo": "Carreta", "piso": 24000, "teto": 28000, "regioes": ["CAPITAL", "INTERIOR"]}
]

# 3. Gerar Pesos Consideráveis (Cenário de Pico Realista) para as 106 Lojas
pesos_cenario_pico = {
    # CAPITAL
    # Rota 1 (Renascença / São Francisco)
    "34": 5600.0, "2061": 2200.0, "61": 6200.0, "60": 3500.0, "90": 3200.0, "225": 4500.0, "229": 3500.0,
    # Rota 2 (Cohama)
    "7": 13500.0, "1413": 2500.0,
    # Rota 3 (Calhau / Maranhão Novo / Shopping)
    "16": 4800.0, "19": 4200.0, "411": 3200.0, "350": 3000.0, "1411": 2200.0,
    # Rota 4 (Itaqui-Bacanga)
    "30": 5500.0, "408": 6800.0, "539": 5200.0, "1418": 2000.0,
    # Rota 5 (Estrada de Ribamar / Operária)
    "11": 4600.0, "502": 4500.0, "425": 4800.0, "218": 3200.0, "1416": 2200.0,
    # Rota 6 (Tirirical / Raposa)
    "17": 4800.0, "217": 4200.0, "122": 4200.0,
    # Rota 7 (Centro / Anil)
    "23": 4600.0, "200": 4500.0, "12": 3500.0, "526": 3500.0, "1417": 2200.0,
    # Rota 8 (João Paulo / Vinhais)
    "96": 26500.0, "20": 5200.0, "1407": 2800.0,
    # Rota 9 (Turu / Divinéia)
    "8": 4600.0, "27": 4500.0, "232": 4200.0, "418": 3500.0, "477": 3500.0, "1410": 2200.0,
    # Rota 10 (Jardim Tropical / São Cristóvão)
    "29": 5200.0, "410": 4200.0, "415": 4200.0, "1415": 2500.0,
    # Rota 11 (Cohab / Cohatrac)
    "5": 5200.0, "201": 4600.0, "18": 4200.0, "412": 4200.0,
    # Rota 12 (Maioba / Paço do Lumiar)
    "93": 5200.0, "255": 4600.0, "259": 5200.0, "269": 4200.0, "409": 4200.0, "554": 4600.0, "31": 4200.0, "1408": 2200.0, "1409": 2200.0,

    # INTERIOR
    # Rota 1 - Lençóis / Litoral Norte (BR-402) - Sem Santa Rita!
    "222": 18500.0, "424": 4500.0, "434": 13500.0, "537": 4200.0, "359": 2500.0, "447": 8800.0,
    # Rota 2 - Baixada Maranhense
    "427": 4800.0, "433": 4200.0, "457": 4800.0, "99": 6500.0, "454": 4200.0, "459": 3800.0,
    # Rota 3 - BR-222 / Baixo Parnaíba (Santa Rita passa aqui)
    "429": 4200.0, "445": 4600.0, "451": 4600.0, "39": 5800.0, "461": 5200.0, "215": 4200.0,
    # Rota 4 - Centro / Mearim / Cocais
    "439": 4200.0, "220": 4600.0, "435": 4600.0, "202": 8800.0,
    # Rota 5 - Caxias / Timon / Floriano / Teresina
    "41": 7500.0, "230": 6500.0, "32": 6500.0, "275": 6500.0, "271": 9500.0,
    "97": 6500.0, "252": 4800.0, "258": 5200.0, "560": 4200.0, "563": 5200.0, "251": 8500.0,
    # Transferência - Ceará (CE)
    "211": 6500.0, "264": 2200.0, "266": 2500.0, "280": 1800.0, "285": 2200.0, "290": 1800.0,
    "503": 2800.0, "506": 3200.0, "507": 3500.0, "514": 2500.0, "516": 1500.0, "518": 1500.0,
    "520": 1800.0, "527": 2200.0, "534": 2200.0, "541": 2200.0
}

# Criar lista completa de lojas com metadados
lojas_simuladas = []
for idx, row in df_cad.iterrows():
    f = str(row["Filial"]).strip()
    reg = str(row["Região"]).strip().upper()
    rota = str(row["Rota Padrão"]).strip()
    nome = str(row["Nome Loja"]).strip()
    cid = str(row["Cidade"]).strip()
    peso = pesos_cenario_pico.get(f, 3500.0) # default se faltar
    lojas_simuladas.append({
        "filial": f, "filial_orig": f, "nome": nome, "regiao": reg,
        "rota": rota, "cidade": cid, "peso": peso, "tem_hoje": True
    })

df_sim = pd.DataFrame(lojas_simuladas)

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))
import automacao

base_enderecos = {}
for _, row in df_cad.iterrows():
    f_cod = str(row["Filial"]).strip()
    base_enderecos[f_cod] = {
        "Filial": f_cod,
        "Nome Loja": str(row.get("Nome Loja", "")),
        "Região": str(row.get("Região", "")),
        "Rota Padrão": str(row.get("Rota Padrão", "")),
        "Restrição Veículo": str(row.get("Restrição Veículo", "Sem restrição")),
        "Bairro": str(row.get("Bairro", "")),
        "Cidade": str(row.get("Cidade", "")),
        "UF": str(row.get("UF", ""))
    }

def selecionar_veiculo(peso, filiais, regiao):
    veiculos_validos = automacao.obter_veiculos_compativeis(filiais, base_enderecos, regiao)
    for v in PARAMETROS_FROTA:
        if v["tipo"] in veiculos_validos and regiao in v["regioes"]:
            if v["piso"] <= peso <= v["teto"]:
                return v["tipo"]
    return None

def otimizar_rota(nome_rota, regiao, pool_lojas, max_paradas=4):
    pool = list(pool_lojas)
    veiculos = []
    
    # 1. Carreta Exclusiva Capital (se >= 24t)
    for s in list(pool):
        if regiao == "CAPITAL" and s["peso"] >= 24000:
            if s["peso"] <= 28000:
                veiculos.append({
                    "rota": nome_rota, "regiao": regiao, "veiculo": "Carreta",
                    "peso": s["peso"], "paradas": [s["filial"]],
                    "obs": "Carreta Exclusiva Capital (1 Loja Direta)"
                })
                pool.remove(s)
            else:
                # Split da carreta (excede 28t)
                veiculos.append({
                    "rota": nome_rota, "regiao": regiao, "veiculo": "Carreta",
                    "peso": 26000.0, "paradas": [f"{s['filial']} (Carga 1)"],
                    "obs": "Carreta Exclusiva Capital (Split Carga 1)"
                })
                s["peso"] -= 26000.0
                s["filial"] = f"{s['filial']} (Sobra)"

    # 2. Busca combinatória gulosa
    melhor_veics = []
    melhor_sobras = list(pool)
    melhor_score = (-1.0, -1.0)
    
    def buscar(rem_pool, veics_acum):
        nonlocal melhor_veics, melhor_sobras, melhor_score
        p_tot = sum(v["peso"] for v in veics_acum)
        score = (len(veics_acum), p_tot)
        if score > melhor_score:
            melhor_score = score
            melhor_veics = list(veics_acum)
            melhor_sobras = list(rem_pool)
            
        n = len(rem_pool)
        if n == 0:
            return
            
        indices = list(range(n))
        candidatos = []
        for k in range(min(max_paradas, n), 0, -1):
            for combo in itertools.combinations(indices, k):
                grp = [rem_pool[i] for i in combo]
                p = sum(s["peso"] for s in grp)
                f_list = [s["filial"] for s in grp]
                v = selecionar_veiculo(p, f_list, regiao)
                if v is not None:
                    tem_restrita = any(s["filial"] in ("560", "411", "418") for s in grp)
                    candidatos.append((grp, v, p, combo, tem_restrita))
                    
        candidatos.sort(key=lambda x: (x[4], x[2]), reverse=True)
        
        for grp, v, p, combo, tem_restrita in candidatos[:12]:
            novo_rem = [rem_pool[i] for i in indices if i not in combo]
            novo_v = {
                "rota": nome_rota, "regiao": regiao, "veiculo": v,
                "peso": p, "paradas": [s["filial"] for s in grp],
                "obs": f"Rota Pura {v}"
            }
            buscar(novo_rem, veics_acum + [novo_v])
            
    buscar(pool, [])
    return veiculos + melhor_veics, melhor_sobras

# Executar Simulação
rotas_geradas = []
sobras_geradas = []

# A. CAPITAL (Rotas 1 a 12)
rotas_cap = sorted(
    df_sim[df_sim["regiao"] == "CAPITAL"]["rota"].unique(),
    key=lambda x: int(''.join(filter(str.isdigit, str(x)))) if any(c.isdigit() for c in str(x)) else 999
)

for r_nome in rotas_cap:
    sub = df_sim[(df_sim["regiao"] == "CAPITAL") & (df_sim["rota"] == r_nome)]
    veics, sobras = otimizar_rota(f"Rota {r_nome}", "CAPITAL", sub.to_dict("records"), max_paradas=4)
    rotas_geradas.extend(veics)
    sobras_geradas.extend(sobras)

# B. INTERIOR (Corredores 1 a 5)
corredores_int = [
    "Rota 1 - Lençóis / Litoral Norte",
    "Rota 2 - Baixada Maranhense",
    "Rota 3 - BR-222 / Baixo Parnaíba",
    "Rota 4 - Centro / Mearim / Cocais",
    "Rota 5 - Caxias / Timon / Floriano"
]

for c_nome in corredores_int:
    sub = df_sim[(df_sim["regiao"] == "INTERIOR") & (df_sim["rota"] == c_nome)]
    veics, sobras = otimizar_rota(c_nome, "INTERIOR", sub.to_dict("records"), max_paradas=4)
    rotas_geradas.extend(veics)
    sobras_geradas.extend(sobras)

# C. CEARÁ (Transferência / Crossdock)
sub_ce = df_sim[df_sim["rota"].str.contains("Ceará", case=False, na=False)]
pool_ce = sub_ce.to_dict("records")
peso_total_ce = sum(s["peso"] for s in pool_ce)
carretas_ce = int(peso_total_ce // 26000)
sobra_ce_peso = peso_total_ce % 26000

print("=" * 115)
print("SIMULAÇÃO DE PICO OPERACIONAL - REDE COMPLETA (106 LOJAS COM PESOS CONSIDERÁVEIS)")
print(f"PESO BRUTO TOTAL DEMANDADO: {df_sim['peso'].sum():,.2f} kg (100% Pedidos do Dia '01-DE HOJE')")
print("=" * 115)

df_rotas = pd.DataFrame(rotas_geradas)
print(f"\nTotal de Veículos Gerados (Capital + Interior MA/PI): {len(df_rotas)}")
print(f"Total de Peso Roteirizado: {df_rotas['peso'].sum():,.2f} kg")

print("\n" + "-" * 115)
print(f"{'Região':<10} | {'Rota':<32} | {'Veículo':<8} | {'Peso (Kg)':>10} | {'Paradas':<40}")
print("-" * 115)
for idx, r in df_rotas.iterrows():
    paradas_str = ", ".join(r["paradas"])
    print(f"{r['regiao']:<10} | {r['rota']:<32} | {r['veiculo']:<8} | {r['peso']:>10.2f} kg | {paradas_str:<40}")

print("\n" + "=" * 115)
print(f">>> TRANSFERÊNCIA CEARÁ (CE) - {len(pool_ce)} LOJAS ({peso_total_ce:,.2f} kg)")
print(f"   • Alocação: {carretas_ce} Carretas de 26.000 kg + 1 Bitruck/Truck de {sobra_ce_peso:,.2f} kg")
print("=" * 115)

print("\n" + "=" * 115)
print("RESUMO DE FROTA UTILIZADA")
print("=" * 115)
cont_frota = Counter(df_rotas["veiculo"])
for v_tipo in ["3/4", "Toco", "Truck", "Bitruck", "Carreta"]:
    qtd = cont_frota.get(v_tipo, 0)
    print(f"   • {v_tipo:<8}: {qtd:>2} veículo(s)")
print(f"   • Carreta (CE): {carretas_ce:>2} veículo(s) de Transferência")

print(f"\nSobras Pendentes: {len(sobras_geradas)} loja(s) totalizando {sum(s['peso'] for s in sobras_geradas):,.2f} kg")
if sobras_geradas:
    for s in sobras_geradas:
        print(f"   --> {s['rota']} | Loja {s['filial']}: {s['peso']:,.2f} kg")
