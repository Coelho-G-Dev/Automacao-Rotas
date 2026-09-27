import os, sys
sys.stdout.reconfigure(encoding='utf-8')
import pandas as pd
import openpyxl

caminho_cad = r"C:\Users\gabri\Desktop\Separacao\Cadastro_Lojas_Enderecos.xlsx"
df_cad = pd.read_excel(caminho_cad)

mapa_interior = {
    # 1. Rota 1 - Lençóis / Litoral Norte (BR-402) - Sem Santa Rita
    "222": ("Rota 1 - Lençóis / Litoral Norte", "Início de Rota BR-402 (Lençóis / Barreirinhas / Tutóia). NÃO vai para Baixada."),
    "424": ("Rota 1 - Lençóis / Litoral Norte", "Início de Rota BR-402 (Lençóis / Barreirinhas / Tutóia)."),
    "434": ("Rota 1 - Lençóis / Litoral Norte", "Pólo Barreirinhas (BR-402)."),
    "537": ("Rota 1 - Lençóis / Litoral Norte", "Pólo Barreirinhas (BR-402)."),
    "447": ("Rota 1 - Lençóis / Litoral Norte", "Final de Rota Tutóia (BR-402 / Litoral)."),
    "359": ("Rota 1 - Lençóis / Litoral Norte", "Dark Store Barreirinhas (BR-402)."),

    # 2. Rota 2 - Baixada Maranhense
    "427": ("Rota 2 - Baixada Maranhense", "Entrada Baixada (Arari). Conecta com Miranda 439 e Santa Rita 429."),
    "433": ("Rota 2 - Baixada Maranhense", "Baixada (Viana)."),
    "457": ("Rota 2 - Baixada Maranhense", "Baixada (São Bento). NÃO mistura com Rosário."),
    "99":  ("Rota 2 - Baixada Maranhense", "Pólo Baixada (Pinheiro)."),
    "454": ("Rota 2 - Baixada Maranhense", "Baixada (Santa Helena)."),
    "459": ("Rota 2 - Baixada Maranhense", "Baixada / Litoral Ocidental (Cururupu)."),

    # 3. Rota 3 - BR-222 / Baixo Parnaíba (Santa Rita passa aqui, Baixada e Centro)
    "429": ("Rota 3 - BR-222 / Baixo Parnaíba", "Passagem BR-135: Adequa em Baixada (Rota 2), BR-222 (Rota 3) ou Centro (Rota 4). NUNCA Lençóis."),
    "445": ("Rota 3 - BR-222 / Baixo Parnaíba", "BR-135 / Entroncamento BR-222 (Itapecuru-Mirim)."),
    "451": ("Rota 3 - BR-222 / Baixo Parnaíba", "BR-222 (Vargem Grande)."),
    "39":  ("Rota 3 - BR-222 / Baixo Parnaíba", "Pólo BR-222 (Chapadinha)."),
    "461": ("Rota 3 - BR-222 / Baixo Parnaíba", "Final de Rota BR-222 / MA (Urbano Santos)."),
    "215": ("Rota 3 - BR-222 / Baixo Parnaíba", "Leste / MA (Coelho Neto)."),

    # 4. Rota 4 - Centro / Mearim / Cocais
    "439": ("Rota 4 - Centro / Mearim / Cocais", "Entroncamento BR-135/222: Faz rota com Centro ou qualquer loja sentido Arari 427."),
    "220": ("Rota 4 - Centro / Mearim / Cocais", "BR-135 Médio Mearim (São Mateus do Maranhão)."),
    "435": ("Rota 4 - Centro / Mearim / Cocais", "MA-020 (Coroatá)."),
    "202": ("Rota 4 - Centro / Mearim / Cocais", "Pólo Cocais BR-316 (Codó)."),

    # 5. Rota 5 - Caxias / Timon / Floriano / Teresina
    "41":  ("Rota 5 - Caxias / Timon / Floriano", "Pólo BR-316 (Mix Caxias)."),
    "230": ("Rota 5 - Caxias / Timon / Floriano", "Pólo BR-316 (Mix Caxias)."),
    "32":  ("Rota 5 - Caxias / Timon / Floriano", "Divisa PI (Mix Timon)."),
    "275": ("Rota 5 - Caxias / Timon / Floriano", "Divisa PI (Mix Timon Alvorada)."),
    "271": ("Rota 5 - Caxias / Timon / Floriano", "Final de Rota Sul PI (Floriano)."),
    "97":  ("Rota 5 - Caxias / Timon / Floriano", "Teresina - PI."),
    "252": ("Rota 5 - Caxias / Timon / Floriano", "Teresina - PI."),
    "258": ("Rota 5 - Caxias / Timon / Floriano", "Teresina - PI."),
    "560": ("Rota 5 - Caxias / Timon / Floriano", "Teresina - PI."),
    "563": ("Rota 5 - Caxias / Timon / Floriano", "Teresina - PI."),
    "251": ("Rota 5 - Caxias / Timon / Floriano", "Litoral PI (Parnaíba)."),

    # 6. Transferência - Ceará (CE)
    "211": ("Transferência - Ceará (CE)", "Pólo rota CE / Piripiri."),
    "264": ("Transferência - Ceará (CE)", "Mix Tianguá - CE."),
    "266": ("Transferência - Ceará (CE)", "Mix Sobral - CE."),
    "280": ("Transferência - Ceará (CE)", "Mix Itapipoca - CE."),
    "285": ("Transferência - Ceará (CE)", "Super Crateús - CE."),
    "290": ("Transferência - Ceará (CE)", "Super Quixeramobim - CE."),
    "503": ("Transferência - Ceará (CE)", "Mix Maracanaú - CE."),
    "506": ("Transferência - Ceará (CE)", "Mix Henrique Jorge - Fortaleza CE."),
    "507": ("Transferência - Ceará (CE)", "Mix Juazeiro do Norte - CE."),
    "514": ("Transferência - Ceará (CE)", "Mix Maranguape - CE."),
    "516": ("Transferência - Ceará (CE)", "Mix Canindé - CE."),
    "518": ("Transferência - Ceará (CE)", "Mix Russas - CE."),
    "520": ("Transferência - Ceará (CE)", "Mix Aracati - CE."),
    "527": ("Transferência - Ceará (CE)", "Mix Zé Walter - Fortaleza CE."),
    "534": ("Transferência - Ceará (CE)", "Mix Caucaia - CE."),
    "541": ("Transferência - Ceará (CE)", "Mix Messejana - Fortaleza CE.")
}

atualizados = 0
for idx, row in df_cad.iterrows():
    f = str(row["Filial"]).strip()
    if f in mapa_interior:
        r_nova, obs_nova = mapa_interior[f]
        df_cad.at[idx, "Região"] = "INTERIOR"
        df_cad.at[idx, "Rota Padrão"] = r_nova
        df_cad.at[idx, "Observação"] = obs_nova
        atualizados += 1

# Inserir Loja 359 se não existir
if "359" not in df_cad["Filial"].astype(str).values:
    nova_linha = {
        "Filial": "359",
        "Nome Loja": "359 - ARMAZEM MATEUS S A DARK STORE BARREIRINHAS",
        "Região": "INTERIOR",
        "Rota Padrão": "Rota 1 - Lençóis / Litoral Norte",
        "Restrição Veículo": "Sem restrição",
        "Bairro": "Centro",
        "Cidade": "Barreirinhas",
        "UF": "MA",
        "Endereço / Referência": "Barreirinhas - MA",
        "Latitude": -2.7489,
        "Longitude": -42.8256,
        "Observação": "Dark Store Barreirinhas (BR-402)."
    }
    df_cad = pd.concat([df_cad, pd.DataFrame([nova_linha])], ignore_index=True)
    print("[OK] Loja 359 cadastrada com sucesso!")

df_cad.to_excel(caminho_cad, index=False)
print(f"[OK] {atualizados} lojas do interior atualizadas no {caminho_cad}!")
