import sys, os
sys.stdout.reconfigure(encoding='utf-8')
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))
import pandas as pd
import itertools
from collections import Counter
import automacao

caminho_cad = r"C:\Users\gabri\Desktop\Separacao\Cadastro_Lojas_Enderecos.xlsx"
df_cad = pd.read_excel(caminho_cad)

base_enderecos = {}
for _, row in df_cad.iterrows():
    f_cod = str(row["Filial"]).strip()
    base_enderecos[f_cod] = {
        "Filial": f_cod, "Nome Loja": str(row.get("Nome Loja", "")),
        "Região": str(row.get("Região", "")), "Rota Padrão": str(row.get("Rota Padrão", "")),
        "Restrição Veículo": str(row.get("Restrição Veículo", "Sem restrição"))
    }

PARAMETROS_FROTA = [
    {"tipo": "3/4",     "piso": 5000,  "teto": 6000,  "regioes": ["CAPITAL"]},
    {"tipo": "Toco",    "piso": 8000,  "teto": 9250,  "regioes": ["CAPITAL", "INTERIOR"]},
    {"tipo": "Truck",   "piso": 12000, "teto": 14250, "regioes": ["CAPITAL", "INTERIOR"]},
    {"tipo": "Bitruck", "piso": 15500, "teto": 16000, "regioes": ["CAPITAL", "INTERIOR"]},
    {"tipo": "Carreta", "piso": 24000, "teto": 28000, "regioes": ["CAPITAL", "INTERIOR"]}
]

# Lojas de Rota 5
lojas_r5 = [
    {"filial": "41",  "peso": 7500.0},
    {"filial": "230", "peso": 6500.0},
    {"filial": "32",  "peso": 6500.0},
    {"filial": "275", "peso": 6500.0},
    {"filial": "271", "peso": 9500.0},
    {"filial": "97",  "peso": 6500.0},
    {"filial": "252", "peso": 4800.0},
    {"filial": "258", "peso": 5200.0},
    {"filial": "560", "peso": 4200.0}, # SOMENTE TOCO!
    {"filial": "563", "peso": 5200.0},
    {"filial": "251", "peso": 8500.0}
]

def selecionar_veiculo(peso, filiais, regiao="INTERIOR"):
    veiculos_validos = automacao.obter_veiculos_compativeis(filiais, base_enderecos, regiao)
    for v in PARAMETROS_FROTA:
        if v["tipo"] in veiculos_validos and regiao in v["regioes"]:
            if v["piso"] <= peso <= v["teto"]:
                return v["tipo"]
    return None

# Gerar candidatos
n = len(lojas_r5)
indices = list(range(n))
candidatos = []
for k in range(min(4, n), 0, -1):
    for combo in itertools.combinations(indices, k):
        grp = [lojas_r5[i] for i in combo]
        p = sum(s["peso"] for s in grp)
        f_list = [s["filial"] for s in grp]
        v = selecionar_veiculo(p, f_list, "INTERIOR")
        if v is not None:
            # Priorizar combos que resolvem lojas restritas (como 560)
            tem_560 = any(s["filial"] == "560" for s in grp)
            candidatos.append((grp, v, p, combo, tem_560))

# Testar busca com prioridade para 560
# Se o combo tem 560, ele é avaliado com alta prioridade para não deixar 560 órfã
candidatos.sort(key=lambda x: (x[4], x[2]), reverse=True)

usados = set()
veics_escolhidos = []
for grp, v, p, combo, tem_560 in candidatos:
    if not any(i in usados for i in combo):
        for i in combo:
            usados.add(i)
        veics_escolhidos.append({
            "veiculo": v, "peso": p, "paradas": [s["filial"] for s in grp]
        })

print("=== ROTA 5 COM PRIORIDADE PARA LOJA 560 (SOMENTE TOCO) ===")
for idx, v in enumerate(veics_escolhidos, 1):
    print(f"Carro {idx}: {v['veiculo']:<7} | Peso: {v['peso']:>9.2f} kg | Paradas: {v['paradas']}")

sobras = [lojas_r5[i] for i in indices if i not in usados]
print("Sobras:", [(s['filial'], s['peso']) for s in sobras])
